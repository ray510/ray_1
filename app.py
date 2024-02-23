from flask import Flask, render_template
import json
import plotly
import plotly.graph_objs as go
import json
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
from plotly.subplots import make_subplots
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
from datetime import datetime, timedelta
import asyncio
app = Flask(__name__)


class FileChangeHandler(FileSystemEventHandler):
    def on_modified(self, event):
        global selected_strikes  # 如果 selected_strikes 在函数外定义，需要声明为 global
        if event.src_path == r"C:\Users\kbl\Desktop\WEB\New_Vega.json":  # 使用完整路径
            print("File changed, reloading data and updating chart...")
            load_and_update_chart()  # 调用加载数据和更新图表的函数

def start_monitoring(path):
    event_handler = FileChangeHandler()
    observer = Observer()
    observer.schedule(event_handler, path=path, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

def load_and_update_chart():
    selected_strikes = ["16500"]
    current_date = datetime.now()

    # 加载 JSON 数据
    with open('New_Vega_1.json', 'r') as file:
        vega_data = json.load(file)
    with open('rec.json', 'r') as file:
        volume_data = json.load(file)

    # 数据处理
    times, vega_values, hsi_index, volume_times, volumes = [], {}, [], [], []
    for strike in selected_strikes:
        vega_values[strike] = []

    for entry in vega_data:
        time_entry = pd.to_datetime(entry['Time'], dayfirst=True)
        if time_entry.date() == current_date.date():
            times.append(time_entry)
            hsi_index.append(float(entry['HSI_Index'].replace(',', '')))
            for strike in selected_strikes:
                strike_vega_values = dict(zip(entry['Strike'].split('|'), entry['Vega'].split('|')))
                vega_values[strike].append(float(strike_vega_values.get(strike, 'nan')))

    for entry in volume_data:
        volume_time_entry = pd.to_datetime(entry['Time'], dayfirst=True)
        if volume_time_entry.date() == current_date.date():
            volume_times.append(volume_time_entry)
            volumes.append(int(entry['Volume']))

    # 创建 DataFrame
    df = pd.DataFrame(list(zip(times, hsi_index, *vega_values.values())), columns=['Time', 'HSI_Index'] + selected_strikes)
    df.set_index('Time', inplace=True)

    # Vega 差分与信号标记
    vega_diffs = df[selected_strikes].diff()
    std_devs = vega_diffs.std()
    multiplier = 0.14
    thresholds = {strike: {'buy': -multiplier * std_dev, 'sell': multiplier * std_dev} for strike, std_dev in std_devs.items()}
    trades = []
    for strike in selected_strikes:
        df[f'{strike}_Vega_diff'] = df[strike].diff()
        std_dev = df[f'{strike}_Vega_diff'].std()
        threshold_buy = -multiplier * std_dev
        threshold_sell = multiplier * std_dev

        df[f'{strike}_buy_signal'] = (df[f'{strike}_Vega_diff'] < threshold_buy)
        df[f'{strike}_sell_signal'] = (df[f'{strike}_Vega_diff'] > threshold_sell)

        # 模拟交易逻辑（简化版）
        # 注意：实际应用中需要更复杂的逻辑
    # 标记买卖点
    #buy_signals, sell_signals = {strike: [] for strike in selected_strikes}, {strike: [] for strike in selected_strikes}
    #buy_signals_text, sell_signals_text = {strike: [] for strike in selected_strikes}, {strike: [] for strike in selected_strikes}

    # 标记买卖点
    buy_signals, sell_signals = {strike: [] for strike in selected_strikes}, {strike: [] for strike in selected_strikes}
    buy_signals_text, sell_signals_text = {strike: [] for strike in selected_strikes}, {strike: [] for strike in selected_strikes}
    all_signals = []

    # 创建图表
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.1, specs=[[{"secondary_y": True}], [{}]])
    fig.add_trace(go.Scatter(x=df.index, y=df['HSI_Index'], mode='lines', name='HSI Index'), row=1, col=1, secondary_y=False)
    for strike in selected_strikes:
        for time, vega_diff in vega_diffs[strike].items():
            if vega_diff < thresholds[strike]['buy']:
                buy_signals[strike].append(time)
                buy_signals_text[strike].append(f"Time: {time}<br>Vega: {df.loc[time, strike]}<br>HSI Index: {df.loc[time, 'HSI_Index']}")
            elif vega_diff > thresholds[strike]['sell']:
                sell_signals[strike].append(time)
                sell_signals_text[strike].append(f"Time: {time}<br>Vega: {df.loc[time, strike]}<br>HSI Index: {df.loc[time, 'HSI_Index']}")
                
        for time, text in zip(buy_signals[strike], buy_signals_text[strike]):
            all_signals.append({
            "time": time,
            "type": "BUY",
            "details": text,
            "Quanity": 1
        })
    # 处理卖出信号
        for time, text in zip(sell_signals[strike], sell_signals_text[strike]):
            all_signals.append({
                "time": time,
                "type": "SELL",
                "details": text,
                "Quanity": 1
            })
    all_signals.sort(key=lambda x: x['time'])
    
    # 转换成JSON格式
    signals_json = json.dumps(all_signals, indent=4, default=str)
    # 写入到文件
    #with open('signals.json', 'w') as file:
        #file.write(signals_json)

    fig.add_trace(go.Bar(x=volume_times, y=volumes, name='Volume'), row=2, col=1)
    fig.add_trace(go.Scatter(x=df.index, y=df[strike], mode='lines', name=f'Vega {strike}'), row=1, col=1, secondary_y=True)
    fig.add_trace(go.Scatter(x=df.index[df[f'{strike}_buy_signal']], y=df[strike][df[f'{strike}_buy_signal']], mode='markers', marker_symbol='triangle-up', marker_color='green', name=f'Buy {strike}'), row=1, col=1, secondary_y=True)
    fig.add_trace(go.Scatter(x=df.index[df[f'{strike}_sell_signal']], y=df[strike][df[f'{strike}_sell_signal']], mode='markers', marker_symbol='triangle-down', marker_color='red', name=f'Sell {strike}'), row=1, col=1, secondary_y=True)
    # 更新图表布局
    fig.update_layout(title='HSI Index and Vega Values with Volume', template="plotly_dark")
    fig.update_yaxes(title_text="HSI Index / Vega Value", row=1, col=1, secondary_y=False)
    fig.update_yaxes(title_text="Volume", row=2, col=1)

    # 显示图表
    #fig.show()

    # 继续绘图逻辑...
    graph_json = json.dumps(fig, cls=plotly.utils.PlotlyJSONEncoder)

    # 在这里，您可能需要将总盈亏以某种形式展示或返回
    #print("Total PnL:", total_pnl)

    # 如果这个函数应该返回图表的JSON，确保最后返回正确的值
    return graph_json

@app.route('/graph-data')
def graph_data():
    graph_json = load_and_update_chart()  # 确保这个函数返回图表的JSON数据
    return graph_json

@app.route('/')
def index():
    graph_json = load_and_update_chart()
    return render_template('index.html', graph_json=graph_json)

if __name__ == '__main__':
    app.run(debug=True)