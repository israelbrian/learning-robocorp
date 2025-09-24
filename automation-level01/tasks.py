from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP

@task
def robot_spare_bin_python():
    """Insert the sales data for the week and export it as a PDF"""
    browser.configure(
        slowmo=600,
    )
    open_the_intranet_website()
    login()
    fill_and_submit_the_sales_form()
    download_excel_file()

def open_the_intranet_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/")

def login():
    """Fills in the login form and clicks the 'Login' button"""
    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")

def fill_and_submit_the_sales_form():
    """Fills in the sales data and click the 'Submit' button"""
    page = browser.page()
    page.fill("#firstname", "Maria")
    page.fill("#lastname", "araujo")
    page.fill("#salesresult", "10")
    page.select_option("#salestarget", "90000")
    page.click("text=Submit")

def download_excel_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)