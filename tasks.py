from robocorp.tasks import task
from robocorp import browser, http, excel
from RPA.PDF import PDF

@task
def robot_spare_bin_python():
    """Robot to enter weekly sales data into the RobotSpareBin Industries Intranet."""

    # Without this, nothing is visible as it is running headless
    browser.configure(
        slowmo=100,
        browser_engine="chromium",
        headless=False,
    )
    open_the_intranet_website()
    log_in()
    download_excel_file()
    fill_form_with_excel_data("SalesData.xlsx")
    screenshot_results()
    export_as_pdf()
    log_out()


def open_the_intranet_website():
    """Navigates to the given URL"""

    browser.goto("https://robotsparebinindustries.com/")

def log_in():
    """Fills in the login form and clicks the 'Log in' button"""

    page = browser.page()
    page.fill("#username", "maria")
    page.fill("#password", "thoushallnotpass")
    page.click("button:text('Log in')")

def download_excel_file():
    """Downloads data from the given URL"""

    # Set overwrite=True in case you want your file to be replaced on every download
    http.download(url="https://robotsparebinindustries.com/SalesData.xlsx", overwrite=True)

def fill_and_submit_form(sales_rep):
    """Fills in the sales data and click the 'Submit' button"""

    page = browser.page()

    page.fill("#firstname", sales_rep["First Name"])
    page.fill("#lastname", sales_rep["Last Name"])
    page.select_option("#salestarget", str(sales_rep["Sales Target"]))
    page.fill("#salesresult", str(sales_rep["Sales"]))
    page.click("text=Submit")

def fill_form_with_excel_data(excel_file):
    """Read data from excel and pass it row by row to fill_and_submit_form() function"""

    workbook = excel.open_workbook(excel_file)
    worksheet = workbook.worksheet("data")

    for row in worksheet.as_table(header=True):
        fill_and_submit_form(row)

def screenshot_results():
    """Take a screenshot of the page"""

    page = browser.page()
    page.screenshot(path="output/sales_summary.png")

def export_as_pdf():
    """Export the data to a pdf file"""

    page = browser.page()

    sales_results = page.locator("#sales-results")
    sales_results_html = sales_results.inner_html()

    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, "output/sales_results.pdf")

def log_out():
    """Presses the 'Log out' button"""
    page = browser.page()  
    page.click("text=Log out")
