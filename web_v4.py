from playwright.sync_api import sync_playwright, Page
import logging
import time
import os
from typing import Optional
from datetime import datetime,timedelta
from dateutil.relativedelta import relativedelta
import schedule
from pathlib import Path
from dateutil.relativedelta import relativedelta
import sys
import win32com.client as win32
import pythoncom

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class Config:
    INITIAL_URL = "http://ap1.dchl.org/tips/_security/login.jsp"
    POPUP_URL = "http://ap1.dchl.org/tips/Login.do"
    TIMEOUTS = {'element': 15000, 'popup': 10000, 'download': 20000}
    SELECTORS = {
        'login': {
            'username': 'input[name="LOGINID"]',
            'password': 'input[name="PWD"]',
            'submit': 'input[name="login"][type="submit"]'
        },
        'navigation': {
            'sales_menu': '#A0',
            'home': 'a[href="/tips/index.jsp"][target="_top"]'
        },
        'audit': {
            'date': 'input[name="SS03"]',
            'format_select': 'select[name="OUT_TYPE"]',
            'a_button': 'input.FunButton[value="A"][accesskey="A"]',
            'back_link': 'a[href="javascript:history.go(-1)"]'
        },
        'monthly_uncollect': {
            'date': 'input[name="SS03"]',
            'format_select': 'select[name="OUT_TYPE"]',
            'b_button': 'input.FunButton[value="B"][accesskey="B"]',
            'back_link': 'a[href="javascript:history.go(-1)"]',
            'start_date': 'input[name="SA13B"]',
            'end_date': 'input[name="SA13E"]'
        },
        'uncollected': {
            'd_button': 'input.FunButton[value="D"][accesskey="D"]',
            'start_date': 'input[name="SA13B"]',
            'end_date': 'input[name="SA13E"]',
        },
        'collections': {
            'e_button': 'input.FunButton[value="E"][accesskey="E"]',
            'date': 'input[name="ST07B"]',
            'format_select': 'select[name="OUT_TYPE"]',
        },
        'exchange_invoice': {
            'd_button': 'input.FunButton[value="D"][accesskey="D"]',
            'date': 'input[name="SA13"]',
            'back_link': 'a[href="javascript:history.go(-1)"]'
        },
        'inventory': {
            'menu': '#A1',
            'd_button': 'input.FunButton[value="D"][accesskey="D"]'
        },
        'inventory_csv': {
            'menu': '#A1',
            'd_button': 'input.FunButton[value="D"][accesskey="D"]',
            'format_select': 'select[name="OUT_TYPE"]'
        },
        'tv_export': {
            'basic_data_menu': '#A3',
            'c_button': 'input.FunButton[value="C"][accesskey="C"]',
            'page_row_count': 'input[name="pageRowCount"]',
            'max_row_count': 'input[name="maxRowCount"]',
            'confirm_button': 'input[type="button"][value="Confirm"]'
        },
        'pdf_download': {
            'c_button': 'input.FunButton[value="C"][accesskey="C"]'
        }
    }
    FUNCTIONS = {
        'inventory': {
            'menu': 'INV',
            'sub_menu': 'IN4'
        },
        'tv_export': {
            'menu': 'BAS',
            'sub_menu': 'BA2'
        }
    }

def close_popup_during_download(page):
    try:
        # 獲取所有打開的頁面
        pages = page.context.pages
        for popup in pages:
            if popup != page:  # 不是主頁面
                # 檢查是否是下載消息窗口
                if "Download Message" in popup.title():
                    logging.info("檢測到下載消息窗口")
                    # 由於沒有明確的關閉按鈕，我們直接關閉這個頁面
                    popup.close()
                    logging.info("已關閉下載消息窗口")
                    return True
        return False
    except Exception as e:
        logging.error(f"處理彈出窗口時發生錯誤: {str(e)}")
        return False
    
def wait_for_popup(page: Page) -> Optional[Page]:
    try:
        with page.expect_popup(timeout=Config.TIMEOUTS['popup']) as popup_info:
            page.goto(Config.INITIAL_URL)
        popup_page = popup_info.value
        logging.info(f"彈出窗口已打開，URL: {popup_page.url}")
        return popup_page
    except Exception as e:
        logging.error(f"等待彈出窗口時發生錯誤: {str(e)}")
        return None

def login_system(page: Page, username: str, password: str) -> Optional[Page]:
    try:
        logging.info(f"正在登錄，當前頁面 URL: {page.url}")
        page.wait_for_load_state('networkidle')
        
        selectors = Config.SELECTORS['login']
        for selector in selectors.values():
            page.wait_for_selector(selector, state="visible", timeout=Config.TIMEOUTS['element'])
        
        page.fill(selectors['username'], username)
        time.sleep(0.5)
        page.fill(selectors['password'], password)
        time.sleep(0.5)
        with page.expect_navigation():
            page.click(selectors['submit'])
        
        page.wait_for_load_state('networkidle')
        logging.info(f"登錄完成，當前頁面 URL: {page.url}")
        return page
    except Exception as e:
        logging.error(f"登錄過程中發生錯誤: {str(e)}")
        page.screenshot(path="login_error.png")
        return None

def return_to_home(page: Page):
    """返回系統主頁"""
    try:
        logging.info("返回主頁...")
        home_selector = Config.SELECTORS['navigation']['home']
        page.wait_for_selector(home_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        page.evaluate('''() => {
            lock();
            window.top.location.href = '/tips/index.jsp';
        }''')
        page.wait_for_load_state('networkidle')
        time.sleep(1)
        logging.info("已返回主頁")
    except Exception as e:
        logging.error(f"返回主頁時發生錯誤: {str(e)}")
        raise

def print_daily_product_audit(login_page: Page, download_path: str, target_date: str) -> bool:
    try:
        logging.info("開始下載 Daily Product Audit 報表...")

        logging.info("點擊 Sales 菜單...")
        sales_selector = Config.SELECTORS['navigation']['sales_menu']
        login_page.wait_for_selector(sales_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.evaluate('P1("A0")')
        time.sleep(1)
        
        logging.info("點擊 Sales Management...")
        login_page.evaluate('processfunction("SOF","SO1")')
        time.sleep(1)
        
        frame = login_page.frame_locator('iframe[name="functionPage"]')
        a_button_selector = Config.SELECTORS['audit']['a_button']
        frame.locator(a_button_selector).wait_for(state="visible", timeout=Config.TIMEOUTS['element'])
        frame.locator(a_button_selector).click()
        
        login_page.wait_for_load_state('networkidle')
        time.sleep(1)
        
        date_selector = Config.SELECTORS['audit']['date']
        login_page.wait_for_selector(date_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.fill(date_selector, target_date)
        
        select_selector = Config.SELECTORS['audit']['format_select']
        login_page.wait_for_selector(select_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.select_option(select_selector, 'xls')
        
        login_page.evaluate('process("6")')
        time.sleep(2)

        report_frame = login_page.frame_locator('iframe[name="reportWin"]')
        back_selector = Config.SELECTORS['audit']['back_link']

        if report_frame.locator(back_selector).is_visible(timeout=1500):
            logging.info("需要返回上一頁")
            report_frame.locator(back_selector).click()
            time.sleep(0.5)
            logging.info("已返回上一頁")
            return False  # 直接返回，不再繼續執行下載操作

        # 只有在沒有看到"返回"按鈕時才執行下載操作
        with login_page.expect_download(timeout=Config.TIMEOUTS['download']) as download_info:
            logging.info("點擊 Confirm 按鈕...")
            confirm_button = login_page.locator("input[type='button'][value='Confirm'][onclick=\"process('6');\"]")
            confirm_button.click()

            try:
                download = download_info.value
                logging.info("下載已開始，等待完成...")

                download_file_name = f"Daily_Product_Audit.xls"
                download_file_path = os.path.join(download_path, download_file_name)
                
                download.save_as(download_file_path)
                
                max_wait_time = 20
                start_time = time.time()
                while time.time() - start_time < max_wait_time:
                    if os.path.exists(download_file_path):
                        file_size = os.path.getsize(download_file_path)
                        if file_size > 0:
                            logging.info(f"下載成功！文件大小: {file_size} bytes")
                            return True
                    time.sleep(1)

                logging.error("文件下載超時或文件大小為0")
                return False

            except TimeoutError:
                logging.error("下載超時")
                return False
            
    except Exception as e:
        logging.error(f"執行過程中發生錯誤: {str(e)}")
        return False

    return False

def print_monthly_uncollect(login_page: Page, download_path: str, target_date: str) -> bool:
    try:
        logging.info("開始下載 Monthly Uncollect 報表...")

        date_obj = datetime.strptime(target_date, "%Y%m%d")
        prev_month = date_obj - relativedelta(months=1)
        target_date_1 = prev_month.strftime("%Y%m%d")
        
        logging.info("點擊 Sales 菜單...")
        sales_selector = Config.SELECTORS['navigation']['sales_menu']
        login_page.wait_for_selector(sales_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.evaluate('P1("A0")')
        time.sleep(1)

        logging.info("點擊 Sales Management...")
        login_page.evaluate('processfunction("SOF","SO1")')
        time.sleep(1)
        
        logging.info("等待並點擊 B 按鈕...")
        frame = login_page.frame_locator('iframe[name="functionPage"]')
        b_button_selector = Config.SELECTORS['monthly_uncollect']['b_button']
        frame.locator(b_button_selector).wait_for(state="visible", timeout=Config.TIMEOUTS['element'])
        frame.locator(b_button_selector).click()
        
        login_page.wait_for_load_state('networkidle')
        time.sleep(1)
        
        logging.info(f"填寫日期範圍: {target_date_1} 到 {target_date}")
        start_date_selector = Config.SELECTORS['monthly_uncollect']['start_date']
        login_page.wait_for_selector(start_date_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.fill(start_date_selector, target_date_1)
        time.sleep(0.1)
        
        end_date_selector = Config.SELECTORS['monthly_uncollect']['end_date']
        login_page.wait_for_selector(end_date_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.fill(end_date_selector, target_date)
        
        logging.info("開始處理報表...")
        login_page.evaluate('process("6")')
        time.sleep(2)
        
        report_frame = login_page.frame_locator('iframe[name="reportWin"]')
        back_selector = Config.SELECTORS['monthly_uncollect']['back_link']
        
        special_back_button = login_page.locator("#go_back")
        if special_back_button.is_visible(timeout=1500):
            logging.info("檢測到特殊的返回按鈕，點擊返回")
            special_back_button.click()
            time.sleep(1)
            logging.info("沒有可下載的報表")
            return False

        logging.info("開始下載報表...")
        with login_page.expect_download(timeout=Config.TIMEOUTS['download']) as download_info:
            login_page.evaluate('process("6")')
            time.sleep(2)  # 給系統時間來響應

            report_frame = login_page.frame_locator('iframe[name="reportWin"]')
            back_selector = Config.SELECTORS['monthly_uncollect']['back_link']

            if report_frame.locator(back_selector).is_visible(timeout=1500):
                logging.info("需要返回上一頁")
                report_frame.locator(back_selector).click()
                login_page.wait_for_load_state('networkidle')
                return False

            download = download_info.value
            download_file_name = "Month_Uncollect.xls"  # 使用固定文件名
            download_file_path = os.path.join(download_path, download_file_name)

            download.save_as(download_file_path)
            logging.info(f"文件已嘗試保存到: {download_file_path}")
            if close_popup_during_download(login_page):
                logging.info("已處理下載消息窗口")

            if os.path.exists(download_file_path):
                file_size = os.path.getsize(download_file_path)
                logging.info(f"下載成功！文件大小: {file_size} bytes")
                return True
            else:
                logging.error("文件下載失敗或找不到文件")
                return False

    except Exception as e:
        logging.error(f"執行過程中發生錯誤: {str(e)}")
        return False
    finally:
        try:
            time.sleep(1)
            return_to_home(login_page)
        except Exception as e:
            logging.error(f"返回主頁失敗: {str(e)}")

def print_uncollected_order_detail(login_page: Page, download_path: str, target_date: str) -> bool:
    try:
        logging.info("開始下載 Uncollected Order Detail 報表...")

        logging.info("點擊 Collection Management...")
        login_page.evaluate('processfunction("SOF","SO2")')
        time.sleep(1)

        frame = login_page.frame_locator('iframe[name="functionPage"]')
        d_button_selector = Config.SELECTORS['uncollected']['d_button']
        frame.locator(d_button_selector).wait_for(state="visible", timeout=Config.TIMEOUTS['element'])
        frame.locator(d_button_selector).click()

        login_page.wait_for_load_state('networkidle')
        time.sleep(1)

        start_date_selector = Config.SELECTORS['uncollected']['start_date']
        login_page.wait_for_selector(start_date_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.fill(start_date_selector, target_date)
        time.sleep(0.1)

        end_date_selector = Config.SELECTORS['uncollected']['end_date']
        login_page.wait_for_selector(end_date_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.fill(end_date_selector, target_date)

        logging.info("點擊第一個 Confirm 按鈕...")
        confirm_button = login_page.locator("input.BTN_PWR[type='button'][value='Confirm'][onclick=\"process('6');\"]")
        confirm_button.click()

        login_page.wait_for_load_state('networkidle')
        time.sleep(2)

        report_frame = login_page.frame_locator('iframe[name="reportWin"]')
        back_selector = Config.SELECTORS['uncollected']['back_link']

        if report_frame.locator(back_selector).is_visible(timeout=1500):
            logging.info("需要返回上一頁")
            report_frame.locator(back_selector).click()
            time.sleep(0.5)
            logging.info("已返回上一頁")
            return False  # 直接返回，不再繼續執行下載操作

        # 只有在沒有看到"返回"按鈕時才執行下載操作
        with login_page.expect_download(timeout=Config.TIMEOUTS['download']) as download_info:
            logging.info("點擊第二個 Confirm 按鈕...")
            confirm_button = login_page.locator("input[type='button'][value='Confirm'][onclick=\"process('6');\"]")
            confirm_button.click()

            try:
                download = download_info.value
                logging.info("下載已開始，等待完成...")

                download_file_name = f"Uncollected_Order_Detail_{target_date}.xls"
                download_file_path = os.path.join(download_path, download_file_name)
                
                download.save_as(download_file_path)
                
                max_wait_time = 20
                start_time = time.time()
                while time.time() - start_time < max_wait_time:
                    if os.path.exists(download_file_path):
                        file_size = os.path.getsize(download_file_path)
                        if file_size > 0:
                            logging.info(f"下載成功！文件大小: {file_size} bytes")
                            return True
                    time.sleep(1)

                logging.error("文件下載超時或文件大小為0")
                return False

            except TimeoutError:
                logging.error("下載超時")
                return False
            
    except Exception as e:
        logging.error(f"執行過程中發生錯誤: {str(e)}")
        return False

    return False

def print_collections(login_page: Page, download_path: str, target_date: str) -> bool:
    try:
        logging.info("開始下載 Print Collections 報表...")

        logging.info("切換到 iframe...")
        frame = login_page.frame_locator('iframe[name="functionPage"]')
        
        logging.info("點擊 E 按鈕...")
        e_button_selector = Config.SELECTORS['collections']['e_button']
        frame.locator(e_button_selector).wait_for(state="visible", timeout=Config.TIMEOUTS['element'])
        frame.locator(e_button_selector).click()
        
        login_page.wait_for_load_state('networkidle')
        time.sleep(1)
        
        logging.info(f"填寫日期: {target_date}")
        date_selector = Config.SELECTORS['collections']['date']
        login_page.wait_for_selector(date_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.fill(date_selector, target_date)
        
        logging.info("選擇 Excel 格式...")
        select_selector = Config.SELECTORS['collections']['format_select']
        login_page.wait_for_selector(select_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.select_option(select_selector, 'xls')
        
        logging.info("開始處理報表...")
        login_page.evaluate('process("6")')
        time.sleep(2)

        # 等待頁面加載
        login_page.wait_for_load_state('networkidle')
        time.sleep(1)

        # 檢查是否出現特殊的返回按鈕
        special_back_button = login_page.locator("#go_back")
        if special_back_button.is_visible(timeout=1500):
            logging.info("檢測到特殊的返回按鈕，點擊返回")
            special_back_button.click()
            time.sleep(1)
            logging.info("沒有可下載的報表")
            return False

        # 如果沒有特殊返回按鈕，進行下載
        logging.info("開始下載報表...")
        with login_page.expect_download(timeout=Config.TIMEOUTS['download']) as download_info:
            try:
                login_page.evaluate('process("6")')
                time.sleep(2)
                download = download_info.value
                download_file_name = f"Collections.xls"
                download_file_path = os.path.join(download_path, download_file_name)
                
                download.save_as(download_file_path)
                
                max_wait_time = 20
                start_time = time.time()
                while time.time() - start_time < max_wait_time:
                    if close_popup_during_download(login_page):
                        logging.info("已處理下載消息窗口")
                    if os.path.exists(download_file_path):
                        file_size = os.path.getsize(download_file_path)
                        if file_size > 0:
                            logging.info(f"下載成功！文件大小: {file_size} bytes")
                            return True
                    time.sleep(1)

                logging.error("文件下載超時或文件大小為0")
                return False

            except TimeoutError:
                logging.error("下載超時")
                return False

    except Exception as e:
        logging.error(f"執行過程中發生錯誤: {str(e)}")
        return False

    return False

def print_exchange_invoice(login_page: Page, download_path: str, target_date: str) -> bool:
    try:
        logging.info("開始處理 Exchange Invoice 報表...")

        logging.info("點擊 Exchange Invoice...")
        login_page.evaluate('processfunction("SOF","SO3")')
        time.sleep(1)

        logging.info("切換到 iframe...")
        frame = login_page.frame_locator('iframe[name="functionPage"]')

        logging.info("點擊 D 按鈕...")
        d_button_selector = Config.SELECTORS['exchange_invoice']['d_button']
        frame.locator(d_button_selector).wait_for(state="visible", timeout=Config.TIMEOUTS['element'])
        frame.locator(d_button_selector).click()

        login_page.wait_for_load_state('networkidle')
        time.sleep(1)

        logging.info(f"填寫日期: {target_date}")
        date_selector = Config.SELECTORS['exchange_invoice']['date']
        login_page.wait_for_selector(date_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.fill(date_selector, target_date)

        logging.info("開始處理報表...")
        login_page.evaluate('process("6")')
        time.sleep(2)

        # 檢查是否出現特殊的返回按鈕
        special_back_button = login_page.locator("#go_back")
        if special_back_button.is_visible(timeout=1500):
            logging.info("檢測到特殊的返回按鈕，點擊返回")
            special_back_button.click()
            time.sleep(1)
            logging.info("沒有可下載的報表")
            return False

        # 如果沒有特殊返回按鈕，進行下載
        logging.info("開始下載報表...")
        with login_page.expect_download(timeout=Config.TIMEOUTS['download']) as download_info:
            try:
                login_page.evaluate('process("6")')
                time.sleep(2)
                download = download_info.value
                download_file_name = f"Exchange_Invoice.xls"
                download_file_path = os.path.join(download_path, download_file_name)
                
                download.save_as(download_file_path)
                
                max_wait_time = 20
                start_time = time.time()
                while time.time() - start_time < max_wait_time:
                    if close_popup_during_download(login_page):
                        logging.info("已處理下載消息窗口")
                    if os.path.exists(download_file_path):
                        file_size = os.path.getsize(download_file_path)
                        if file_size > 0:
                            logging.info(f"下載成功！文件大小: {file_size} bytes")
                            return True
                    time.sleep(1)

                logging.error("文件下載超時或文件大小為0")
                return False

            except TimeoutError:
                logging.error("下載超時")
                return False

    except Exception as e:
        logging.error(f"執行過程中發生錯誤: {str(e)}")
        return False

    return False

def print_inventory_excel(login_page: Page, download_path: str) -> bool:
    try:
        logging.info("开始下载 Inventory Excel 报表...")

        logging.info("点击 Inventory 菜单...")
        inventory_selector = Config.SELECTORS['inventory']['menu']
        login_page.wait_for_selector(inventory_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.evaluate('P1("A1")')
        time.sleep(1)

        logging.info("点击 Inventory Reports...")
        func_config = Config.FUNCTIONS['inventory']
        login_page.evaluate(f'processfunction("{func_config["menu"]}","{func_config["sub_menu"]}")')
        time.sleep(1)

        frame = login_page.frame_locator('iframe[name="functionPage"]')

        logging.info("点击 D 按钮...")
        d_button_selector = Config.SELECTORS['inventory']['d_button']
        frame.locator(d_button_selector).wait_for(state="visible", timeout=Config.TIMEOUTS['element'])
        frame.locator(d_button_selector).click()

        login_page.wait_for_load_state('networkidle')
        time.sleep(1)

        logging.info("开始处理报表...")
        with login_page.expect_download(timeout=Config.TIMEOUTS['download']) as download_info:
            login_page.evaluate('process("6")')

            try:
                download = download_info.value
                current_date = datetime.now().strftime("%Y%m%d")
                download_file_name = f"Inventory.xls"
                download_file_path = os.path.join(download_path, download_file_name)

                download.save_as(download_file_path)
                logging.info(f"文件已尝试保存到: {download_file_path}")

                if os.path.exists(download_file_path):
                    file_size = os.path.getsize(download_file_path)
                    logging.info(f"下载成功！文件大小: {file_size} bytes")
                    if file_size == 0:
                        logging.warning("警告：下载的文件大小为0字节")
                        return False
                    return True
                else:
                    logging.error("文件下载失败或找不到文件")
                    return False
            except Exception as e:
                logging.error(f"下载过程中发生错误: {str(e)}")
                return False

    except Exception as e:
        logging.error(f"执行过程中发生错误: {str(e)}")
        return False
    finally:
        try:
            return_to_home(login_page)
        except Exception as e:
            logging.error(f"返回主页失败: {str(e)}")

def download_inventory_pdf(login_page: Page, download_path: str) -> bool:
    try:
        logging.info("開始下載 Inventory PDF 報告...")

        frame = login_page.frame_locator('iframe[name="functionPage"]')

        logging.info("點擊 C 按鈕...")
        c_button_selector = Config.SELECTORS['pdf_download']['c_button']
        frame.locator(c_button_selector).wait_for(state="visible", timeout=Config.TIMEOUTS['element'])
        frame.locator(c_button_selector).click()

        login_page.wait_for_load_state('networkidle')
        time.sleep(1)

        logging.info("點擊 Confirm 按鈕...")
        login_page.evaluate('process("6")')

        login_page.wait_for_load_state('networkidle')
        time.sleep(3)

        form_data = login_page.evaluate('''() => {
            const form = document.forms['IN4R745f'];
            const formData = {};
            for (let element of form.elements) {
                if (element.name) {
                    formData[element.name] = element.value;
                }
            }
            formData['OUT_TYPE'] = 'pdf';
            formData['actionCode'] = '6';
            formData['alias'] = '6';
            return formData;
        }''')

        headers = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Accept': 'application/pdf',
            'Cookie': '; '.join([f"{c['name']}={c['value']}" for c in login_page.context.cookies()])
        }

        logging.info("發送 POST 請求下載 PDF...")
        response = login_page.context.request.post(
            login_page.url,
            form=form_data,
            headers=headers
        )

        if response.ok:
            filename = f"inventory_report.pdf"
            download_file_path = os.path.join(download_path, filename)

            with open(download_file_path, 'wb') as f:
                f.write(response.body())

            if os.path.exists(download_file_path):
                file_size = os.path.getsize(download_file_path)
                logging.info(f"PDF 文件已保存，大小: {file_size} bytes")
                
                with open(download_file_path, 'rb') as f:
                    if f.read(4).startswith(b'%PDF'):
                        logging.info("成功下載 PDF 文件")
                        return True
                    else:
                        logging.warning("警告：保存的文件可能不是 PDF 格式")
                        return False
            else:
                logging.error("文件未能成功保存")
                return False
        else:
            logging.error(f"請求失敗，狀態碼: {response.status}")
            return False

    except Exception as e:
        logging.error(f"執行 PDF 下載過程中發生錯誤: {str(e)}")
        return False
    finally:
        try:
            return_to_home(login_page)
        except Exception as e:
            logging.error(f"返回主頁失敗: {str(e)}")

def inventory_csv(login_page: Page, download_path: str) -> bool:
    try:
        logging.info("開始下載 Inventory CSV 報告...")

        frame = login_page.frame_locator('iframe[name="functionPage"]')

        logging.info("點擊 C 按鈕...")
        c_button_selector = Config.SELECTORS['pdf_download']['c_button']
        frame.locator(c_button_selector).wait_for(state="visible", timeout=Config.TIMEOUTS['element'])
        frame.locator(c_button_selector).click()

        login_page.wait_for_load_state('networkidle')
        time.sleep(1)

        logging.info("選擇 CSV 格式...")
        select_selector = Config.SELECTORS['inventory_csv']['format_select']
        login_page.wait_for_selector(select_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.select_option(select_selector, 'csv')

        logging.info("點擊 Confirm 按鈕...")
        with login_page.expect_download(timeout=Config.TIMEOUTS['download']) as download_info:
            login_page.evaluate('process("6")')
            time.sleep(2)

            back_selector = Config.SELECTORS['audit']['back_link']
            if login_page.locator(back_selector).is_visible(timeout=1500):
                logging.info("需要返回上一頁")
                login_page.locator(back_selector).click()
                login_page.wait_for_load_state('networkidle')
                return False

            try:
                download = download_info.value
                download_file_name = f"inventory.csv"
                download_file_path = os.path.join(download_path, download_file_name)

                download.save_as(download_file_path)
                logging.info(f"文件已嘗試保存到: {download_file_path}")

                if os.path.exists(download_file_path):
                    file_size = os.path.getsize(download_file_path)
                    logging.info(f"下載成功！文件大小: {file_size} bytes")
                    return True
                else:
                    logging.error("文件下載失敗或找不到文件")
                    return False
            except Exception as e:
                logging.error(f"下載過程中發生錯誤: {str(e)}")
                return False

    except Exception as e:
        logging.error(f"執行過程中發生錯誤: {str(e)}")
        return False
    finally:
        try:
            return_to_home(login_page)
        except Exception as e:
            logging.error(f"返回主頁失敗: {str(e)}")

def wait_for_save_button(page, timeout=60000):
    try:
        save_button_selector = "input[type='button'][name='save'][value='Save'][onclick=\"process('10')\"]"
        page.wait_for_selector(save_button_selector, state='visible', timeout=timeout)
        logging.info("Save 按鈕已出現")
        return True
    except TimeoutError:
        logging.warning("等待 Save 按鈕出現超時")
        return False
    
def export_tv_data(login_page: Page, download_path: str) -> bool:
    try:
        logging.info("開始執行 TV 數據導出...")

        logging.info("點擊 Basic Data 選單...")
        basic_data_selector = Config.SELECTORS['tv_export']['basic_data_menu']
        login_page.wait_for_selector(basic_data_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.evaluate('P1("A3")')
        time.sleep(1)

        logging.info("點擊 TV...")
        func_config = Config.FUNCTIONS['tv_export']
        login_page.evaluate(f'processfunction("{func_config["menu"]}","{func_config["sub_menu"]}")')
        time.sleep(1)

        frame = login_page.frame_locator('iframe[name="functionPage"]')

        logging.info("點擊 C 按鈕...")
        c_button_selector = Config.SELECTORS['tv_export']['c_button']
        frame.locator(c_button_selector).wait_for(state="visible", timeout=Config.TIMEOUTS['element'])
        frame.locator(c_button_selector).click()

        login_page.wait_for_load_state('networkidle')
        time.sleep(1)

        logging.info("填寫 PageRowCount 和 MaxRowCount...")
        page_row_selector = Config.SELECTORS['tv_export']['page_row_count']
        max_row_selector = Config.SELECTORS['tv_export']['max_row_count']
        login_page.wait_for_selector(page_row_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.fill(page_row_selector, '9999')
        login_page.wait_for_selector(max_row_selector, state="visible", timeout=Config.TIMEOUTS['element'])
        login_page.fill(max_row_selector, '9999')

        logging.info("點擊第一個 Confirm 按鈕...")
        login_page.evaluate('process("6")')

        logging.info("等待數據加載...")
        wait_for_save_button(login_page)

        logging.info("點擊導出 Excel 按鈕...")
        with login_page.expect_download(timeout=30000) as download_info:
            login_page.evaluate('process("25")')

            try:
                download = download_info.value
                current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
                download_file_name = f"TV.xls"
                download_file_path = os.path.join(download_path, download_file_name)

                download.save_as(download_file_path)
                logging.info(f"文件已嘗試保存到: {download_file_path}")

                if os.path.exists(download_file_path):
                    file_size = os.path.getsize(download_file_path)
                    logging.info(f"下載成功！文件大小: {file_size} bytes")
                    if file_size == 0:
                        logging.warning("警告：下載的文件大小為0字節")
                        return False
                    return True
                else:
                    logging.error("文件下載失敗或找不到文件")
                    return False
            except Exception as e:
                logging.error(f"下載過程中發生錯誤: {str(e)}")
                return False

    except Exception as e:
        logging.error(f"執行過程中發生錯誤: {str(e)}")
        return False
    finally:
        try:
            return_to_home(login_page)
        except Exception as e:
            logging.error(f"返回主頁失敗: {str(e)}")

def get_desktop_path(company_name: str) -> str:
    """
    獲取桌面路徑並創建公司資料夾
    
    Args:
        company_name: 公司名稱
        
    Returns:
        str: 完整的下載路徑
    """
    # 獲取桌面路徑
    desktop_path = r'\\files01-wtc.kmml.local\ON-Warehouse\各場庫存及容量\python_data\test'
    
    # 在桌面創建下載資料夾
    download_folder = os.path.join(desktop_path, "download_data", company_name)
    
    # 確保資料夾存在
    os.makedirs(download_folder, exist_ok=True)
    
    return download_folder

def create_company_folder(base_path, company):
    company_path = os.path.join(base_path, company)
    if not os.path.exists(company_path):
        os.makedirs(company_path)
    return company_path

####### 轉換file Start #######
def initialize_excel():
    """初始化Excel應用程序"""
    try:
        # 初始化COM
        pythoncom.CoInitialize()
        
        # 首先嘗試使用 gencache
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
        except:
            # 如果gencache失敗，使用普通的Dispatch
            excel = win32.Dispatch('Excel.Application')
        
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel
    except Exception as e:
        print(f"初始化Excel時發生錯誤: {e}")
        return None

def safe_quit_excel(excel):
    """安全地關閉Excel應用程序"""
    try:
        if excel:
            excel.DisplayAlerts = False
            excel.Quit()
            pythoncom.CoUninitialize()
    except:
        pass

def excel_save_multiple_files(file_mappings):
    """
    批量另存為Excel文件到指定路徑
    
    參數:
    file_mappings (list): 包含源文件和目標文件路徑的字典列表
    """
    excel = None
    current_workbook = None
    
    try:
        # 初始化Excel
        excel = initialize_excel()
        if not excel:
            raise Exception("無法初始化Excel應用程序")
        
        # 遍歷並處理每個文件
        for mapping in file_mappings:
            input_file = mapping['input']
            output_file = mapping['output']
            
            try:
                # 確保輸入文件路徑是有效的
                input_path = Path(input_file)
                if not input_path.exists():
                    print(f"文件不存在: {input_file}")
                    continue
                
                # 確保輸出目錄存在
                output_path = Path(output_file)
                output_path.parent.mkdir(parents=True, exist_ok=True)
                
                # 關閉之前的工作簿（如果有）
                if current_workbook:
                    try:
                        current_workbook.Close(False)
                    except:
                        pass
                
                # 打開新的工作簿
                current_workbook = excel.Workbooks.Open(str(input_path))
                
                # 檢查是否需要覆蓋現有文件
                if output_path.exists():
                    try:
                        os.remove(str(output_path))
                    except Exception as e:
                        print(f"無法刪除現有文件 {output_file}: {e}")
                        continue
                
                # 另存為
                current_workbook.SaveAs(str(output_path))
                print(f"文件 {input_file} 已成功另存為 {output_file}")
                
            except Exception as file_error:
                print(f"處理文件 {input_file} 時發生錯誤: {file_error}")
                
            finally:
                # 確保工作簿被關閉
                if current_workbook:
                    try:
                        current_workbook.Close(False)
                    except:
                        pass
                current_workbook = None
    
    except Exception as e:
        print(f"發生全局錯誤: {e}")
        
    finally:
        # 關閉當前工作簿
        if current_workbook:
            try:
                current_workbook.Close(False)
            except:
                pass
        
        # 安全地關閉Excel
        safe_quit_excel(excel)

def register_excel_com():
    """註冊Excel COM組件"""
    try:
        import win32com.client.gencache
        # 強制重新生成gencache
        win32com.client.gencache.EnsureModule('{00020813-0000-0000-C000-000000000046}', 0, 1, 9)
        return True
    except Exception as e:
        print(f"註冊Excel COM組件時發生錯誤: {e}")
        return False
    

########  轉換file _ END #########

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def run(playwright):
    #if datetime.today().weekday() == 0:  # 0 代表星期一
    #    logging.info("今天是星期一，程式不執行")
    #    return

    base_download_path = r'\\files01-wtc.kmml.local\ON-Warehouse\各場庫存及容量\python_data\download_data'
    companies = ["BV", "BD", "TS", "TD", "MM", "FR", "WH", "ED", "EF", "ES", "EB", "SM"]  # 可以根据需要添加更多公司
    all_results = {}

    for company in companies:
        logging.info(f"\n開始處理公司: {company}")
        company_start_time = time.time()
        successful_tasks = 0
        total_tasks = 0
        company_results = []
        
        company_download_path = create_company_folder(base_download_path, company)
        
        browser = playwright.chromium.launch(channel="msedge", headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        username = f"{company}.OM079"
        password = "0000"
        target_date = (datetime.today() - timedelta(days=1)).strftime('%Y%m%d')

        popup_page = wait_for_popup(page)
        if popup_page:
            logged_in_page = login_system(popup_page, username, password)
            if logged_in_page:
                logging.info("登錄成功，開始執行下載任務")
                
                total_tasks += 1
                start_time = time.time()
                success_daily = print_daily_product_audit(logged_in_page, company_download_path, target_date)
                end_time = time.time()
                company_results.append({"task": "Daily Product Audit", "success": success_daily, "duration": end_time - start_time})
                if success_daily:
                    logging.info("Daily Product Audit 報表下載成功")
                    successful_tasks += 1
                else:
                    logging.warning("Daily Product Audit 報表下載失敗或不需要下載")
                if success_daily:
                    return_to_home(logged_in_page)

                time.sleep(1.5)
                total_tasks += 1
                start_time = time.time()
                # 下載 Monthly Uncollect 報表
                success_monthly = print_monthly_uncollect(logged_in_page, company_download_path, target_date)
                end_time = time.time()
                company_results.append({"task": "Monthly Uncollect", "success": success_monthly, "duration": end_time - start_time})
                if success_monthly:
                    logging.info("Monthly Uncollect 報表下載成功")
                    successful_tasks += 1
                else:
                    logging.warning("Monthly Uncollect 報表下載失敗或不需要下載")
                
                time.sleep(1.5)

                total_tasks += 1
                start_time = time.time()
                # 下載 Uncollected Order Detail 報表
                success_uncollected = print_uncollected_order_detail(logged_in_page, company_download_path, target_date)
                end_time = time.time()
                company_results.append({"task": "Uncollected Order Detail", "success": success_uncollected, "duration": end_time - start_time})
                if success_uncollected:
                    logging.info("Uncollected Order Detail 報表下載成功")
                    successful_tasks += 1
                else:
                    logging.warning("Uncollected Order Detail 報表下載失敗或不需要下載")
                return_to_home(logged_in_page)

                time.sleep(1.5)

                total_tasks += 1
                start_time = time.time()
                # 下載 Print Collections 報表
                success_collections = print_collections(logged_in_page, company_download_path, target_date)
                end_time = time.time()
                company_results.append({"task": "Print Collections", "success": success_collections, "duration": end_time - start_time})
                if success_collections:
                    logging.info("Print Collections 報表下載成功")
                    successful_tasks += 1
                else:
                    logging.warning("Print Collections 報表下載失敗或不需要下載")
                return_to_home(logged_in_page)
                
                time.sleep(1.5)

                total_tasks += 1
                start_time = time.time()
                # 下載 Exchange Invoice 報表
                success_exchange = print_exchange_invoice(logged_in_page, company_download_path, target_date)
                end_time = time.time()
                company_results.append({"task": "Exchange Invoice", "success": success_exchange, "duration": end_time - start_time})
                if success_exchange:
                    logging.info("Exchange Invoice 報表下載成功")
                    successful_tasks += 1
                else:
                    logging.warning("Exchange Invoice 報表下載失敗或不需要下載")
                return_to_home(logged_in_page)
                
                time.sleep(1.5)

                total_tasks += 1
                start_time = time.time()
                # 下載 Inventory Excel 報表
                success_inventory = print_inventory_excel(logged_in_page, company_download_path)
                end_time = time.time()
                company_results.append({"task": "Inventory Excel", "success": success_inventory, "duration": end_time - start_time})
                if success_inventory:
                    logging.info("Inventory Excel 報表下載成功")
                    successful_tasks += 1
                else:
                    logging.warning("Inventory Excel 報表下載失敗或不需要下載")

                time.sleep(1.5)

                total_tasks += 1
                start_time = time.time()
                # 下載 Inventory PDF 報表
                success_inventory_pdf = download_inventory_pdf(logged_in_page, company_download_path)
                end_time = time.time()
                company_results.append({"task": "Inventory PDF", "success": success_inventory_pdf, "duration": end_time - start_time})
                if success_inventory_pdf:
                    logging.info("Inventory PDF 報表下載成功")
                    successful_tasks += 1
                else:
                    logging.warning("Inventory PDF 報表下載失敗或不需要下載")
                
                time.sleep(1.5)
                
                total_tasks += 1
                start_time = time.time()
                # 下載 Inventory CSV 報表
                success_inventory_csv = inventory_csv(logged_in_page, company_download_path)
                end_time = time.time()
                company_results.append({"task": "Inventory CSV", "success": success_inventory_csv, "duration": end_time - start_time})
                if success_inventory_csv:
                    logging.info("Inventory CSV 報表下載成功")
                    successful_tasks += 1
                else:
                    logging.warning("Inventory CSV 報表下載失敗或不需要下載")

                time.sleep(1.5)

                total_tasks += 1
                start_time = time.time()
                success_tv_export = export_tv_data(logged_in_page, company_download_path)
                end_time = time.time()
                company_results.append({"task": "TV Export", "success": success_tv_export, "duration": end_time - start_time})
                if success_tv_export:
                    logging.info("TV 數據導出成功")
                    successful_tasks += 1
                else:
                    logging.warning("TV 數據導出失敗或不需要下載")

                logging.info("所有下載任務已完成")
            else:
                logging.error("登錄失敗")
        else:
            logging.error("無法打開彈出窗口")

        browser.close()
        logging.info("瀏覽器已關閉，程序執行完畢")
        company_end_time = time.time()
        company_total_time = company_end_time - company_start_time
        
        logging.info(f"\n公司 {company} 處理摘要:")
        logging.info(f"總執行時間: {company_total_time:.2f} 秒")
        logging.info(f"成功執行任務數: {successful_tasks}/{total_tasks}")
        all_results[company] = company_results

    # 輸出總體摘要
    logging.info("\n========= 總體執行結果摘要 =========")
    for company, results in all_results.items():
        if results:
            total_success = sum(1 for r in results if r['success'])
            total_time = sum(r['duration'] for r in results)
            logging.info(f"\n公司: {company}")
            logging.info(f"總執行時間: {round(total_time, 2)} 秒")
            logging.info(f"成功數量: {total_success}/{len(results)}")

    if not register_excel_com():
        print("無法註冊Excel COM組件，程序將退出")
        sys.exit(1)
    
    companies = ["BV", "BD", "TS", "TD", "MM", "FR", "WH", "ED", "EF", "ES", "EB"]
    
    # 生成文件映射
    file_mappings = []
    for company in companies:
        base_input = Path(f'//files01-wtc.kmml.local/ON-Warehouse/各場庫存及容量/python_data/download_data/{company}')
        base_output = Path(f'//files01-wtc.kmml.local/ON-Warehouse/各場庫存及容量/({company})資料庫更新')
        
        mappings = [
            {
                'input': str(base_input / 'Inventory.xls'),
                'output': str(base_output / f'({company})Inventory.xls')
            },
            {
                'input': str(base_input / 'TV.xls'),
                'output': str(base_output / f'({company})TV info maintenance.xls')
            },
            {
                'input': str(base_input / 'Daily_Product_Audit.xls'),
                'output': str(base_output / f'({company})每日銷售數.xls')
            },
            {
                'input': str(base_input / 'Month_Uncollect.xls'),
                'output': str(base_output / f'({company}) EDC客未取數(12個月).xls')
            },
            {
                'input': str(base_input / 'Uncollected.xls'),
                'output': str(base_output / '客未取貨.xls')
            },
            {
                'input': str(base_input / 'Collections.xls'),
                'output': str(base_output / 'Collection Detail.xls')
            },
            {
                'input': str(base_input / 'Exchange_Invoice.xls'),
                'output': str(base_output / '換貨紀錄.xls')
            }
        ]
        file_mappings.extend(mappings)
    
    excel_save_multiple_files(file_mappings)
if __name__ == "__main__":
    with sync_playwright() as playwright:
        run(playwright)