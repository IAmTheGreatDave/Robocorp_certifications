from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    browser.configure(
        slowmo=10,    
    )
    download_orders_file()
    open_order_page()
    get_orders()
    zip_all_receipts()

def open_order_page():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_orders_file():
    http = HTTP()
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)

def close_popup(page):
    try:
        page.click('button:text("OK")')
    except:
        pass #if the popup wouldn't appear for some reason

def check_for_errors(page, error_locator):
    while page.locator(error_locator).count() > 0:
        page.click('button:text("ORDER")')
        page.wait_for_timeout(500)



def save_scsreenshot(page,order_number, screenshot_locator):
    screenshot_element = page.locator(screenshot_locator)
    screenshot_element.screenshot(path=f"output/screenshots/screenshot_{order_number}.png") 
    # page.screenshot(path=f"output/receipts/screenshot_{order_number}.png")
    screenshot = f"output/screenshots/screenshot_{order_number}.png"
    return screenshot

def save_receipt_as_pdf(page, order_number):
    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, f"output/receipts/order_{order_number}.pdf")
    pdf_file = f"output/receipts/order_{order_number}.pdf"
    print(pdf_file)
    return pdf_file
    
def attach_screenshot_to_pdf(pdf_file, screenshot):
    screenshot_list = [screenshot]
    pdf = PDF()
    pdf.add_files_to_pdf(screenshot_list, pdf_file, append=True)

def zip_all_receipts():
    archive = Archive()
    archive.archive_folder_with_zip(folder="output/receipts", archive_name="output/zipped_receipts.zip")
    
    

def send_order(order):
    print(order)
    page = browser.page()
    page.set_viewport_size({'width':1920, 'height':1080})
    close_popup(page)
    page.select_option('#head',str(order['Head']))
    page.click(f"id=id-body-{order['Body']}")
    page.fill("input[placeholder='Enter the part number for the legs']", order['Legs'], timeout=2000)
    page.screenshot()
    page.fill('//*[@id="address"]', order['Address'])
    page.click('button:text("Preview")')
    page.click('button:text("ORDER")')
    error_locator = 'div[role="alert"].alert.alert-danger'
    
    check_for_errors(page,error_locator)
    screenshot_locator = "#robot-preview-image"
    screenshot = save_scsreenshot(page, order['Order number'], screenshot_locator)
    receipt = save_receipt_as_pdf(page, order['Order number'])
    attach_screenshot_to_pdf(receipt, screenshot)
    page.click('button:text("ORDER ANOTHER ROBOT")')
    page.wait_for_timeout(1000)

def get_orders():
    '''reads orders to table'''
    orders = Tables()
    orders_sheet = orders.read_table_from_csv('orders.csv', header=True)
    print(f'Amount of orders is: {len(orders_sheet)}')

    for row in orders_sheet:
        send_order(row)
    
    

    