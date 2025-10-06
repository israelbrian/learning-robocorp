from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
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
        slowmo=100,
    )
    open_website()
    download_excel_file()
    fill_form_with_excel_data()
    orders = get_orders()

def open_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)    

def fill_and_submit_the_sales_form(sales_rep):
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()
    page.select_option("#head", str(sales_rep["Order"]))
    # page.get_by_label("#head", sales_rep["Order"])
    # page.fill("#lastname", sales_rep["Last Name"])
    page.fill("#placeholder=Enter the part number for the legs", str(sales_rep["Sales"]))
    page.click("text=Submit")

def fill_form_with_excel_data():
    """Reads the excel file and fills in the sales form"""
    excel = Files()
    excel.open_workbook("orders.xlsx")
    worksheet = excel.read_worksheet_as_table("data", header=True)
    excel.close_workbook()

    for row in worksheet:
        fill_and_submit_the_sales_form(row)

def get_orders():
"""Baixe o arquivo de pedidos, leia-o como uma tabela e retorne o resultado"""  
    return 0