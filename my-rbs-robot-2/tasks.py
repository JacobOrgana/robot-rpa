from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive
import os, shutil

@task
def order_robots_from_RobotSpareBin():
    """ Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images. """
    open_robot_order_website()
    download_csv_file()
    create_new_folder()
    get_orders()
    archive_receipts()
    shutil.rmtree(os.getcwd() + "\\output\\Orders")

def open_robot_order_website():
    """"Navigates to given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_csv_file():
    """Downloads excel file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/ orders.csv", overwrite=True)

def fill_form(orderrow):
    """Read data from csv and fill in the sales form"""
    page = browser.page()
    #page.fill("#Order number", orderrow["order number"])
    close_annoying_modal()
    page.select_option("#head", str(orderrow["Head"]))
    page.click(f"//input[@value='{str(orderrow['Body'])}']")
    page.fill("input.form-control[type='number']",str(orderrow["Legs"]))
    page.fill("#address", str(orderrow["Address"]))
    order_folder_output(orderrow["Order number"])
    order_click()
    receipt= store_receipt_as_pdf(orderrow["Order number"])
    robot= screenshot_robot(orderrow["Order number"])
    embed_screenshot_to_receipt(robot, receipt, orderrow["Order number"])
    page.click("button#order-another")


def  get_orders():
    """Import Data from CSV"""
    csv = Tables()
    orderslist = csv.read_table_from_csv(" orders.csv", header=True)
    for row in orderslist:
        fill_form(row)

def store_receipt_as_pdf(order_number):
    """Export the data to a pdf file"""
    page = browser.page()
    order_result_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    order_result = '../my-rbs-robot-2/output/Orders/'+order_number+'/Ordernumber_'+order_number+'.pdf'
    pdf.html_to_pdf(order_result_html, order_result)
    return order_result

def screenshot_robot(order_number):
    """Take a screenshot of the page"""
    page = browser.page()
    robot_preview = ['../my-rbs-robot-2/output/Orders/'+order_number+'/robot-image'+order_number+'.png']
    page.screenshot(path=robot_preview[0])
    return robot_preview

def close_annoying_modal():
    page =  browser.page()
    page.click("text=OK")

def embed_screenshot_to_receipt(screenshot, pdf_file, order_number):
    pdf= PDF()
    pdf.open_pdf(pdf_file)
    pdf.add_files_to_pdf(screenshot, pdf_file,True)
    pdf.close_pdf()

def order_folder_output(order_number):
    os.makedirs(os.getcwd() +"\\output\\Orders\\"+order_number, exist_ok= True)
    
def archive_receipts():
    zip= Archive()
    zip.archive_folder_with_zip("../my-rbs-robot-2/output/Orders/", "Orders.zip", recursive=True)

def create_new_folder():
    os.makedirs(os.getcwd() + "\\output\\Orders", exist_ok= True)
    
def order_click():
    page =  browser.page()
    page.click("button#order")
    if page.is_visible(".alert.alert-danger"):
        page.click ("button#order")

    if page.is_visible (".alert.alert-danger"):
        page.click ("button#order")
    
    if page.is_visible (".alert.alert-danger"):
        page.click ("button#order")

    if page.is_visible (".alert.alert-danger"):
        page.click ("button#order")