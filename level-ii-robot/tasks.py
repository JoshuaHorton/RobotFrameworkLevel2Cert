from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.Excel.Files import Files
'''from RPA.Browser import Browser'''
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
        slowmo=100,
    )
    open_robot_order_website()
    close_annoying_modal()
    download_csv_file()
    orders = get_orders()
    for row in orders:
        fill_and_submit_order_form(row)
    archive_receipts()



def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_csv_file():
    """Downloads csv file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)

def get_orders():
    '''Read orders.csv and return as a Table'''
    library = Tables()
    orders = library.read_table_from_csv(
        "orders.csv", 
        columns=["Order number", "Head", "Body", "Legs", "Address"]
        )
    return orders

def close_annoying_modal():
    """Clicks any of the non-functional buttons just to close the modal"""
    page = browser.page()
    page.click("button.btn.btn-dark")



def fill_and_submit_order_form(ord):
    '''Issues order request for one row from csv'''
    page = browser.page()

    page.select_option("#head", str(ord["Head"]))
    value = ord["Body"]
    for li in page.get_by_role("radio", name="body").all():
        if li.input_value() == value:
            li.set_checked(True)
            break
    page.get_by_placeholder("Enter the part number for the legs").fill(str(ord["Legs"]))
    page.fill("#address", ord["Address"])

    imageID = "files/" + str(ord["Order number"]) + ".png"
    grab_robot_image(imageID)

    try:
        page.click("button.btn.btn-primary")
        print("Clicked Order in fill_and_submit_order_form for " + str(ord["Order number"]))
    except:
        print("An Error Happened in fill_and_submit_order_form for " + str(ord["Order number"]) + "- re-trying once")
        page.click("button.btn.btn-primary")

    path_to_receipt = store_receipt_as_pdf(str(ord["Order number"]), page)
    embed_screenshot_to_receipt(imageID, path_to_receipt)


def grab_robot_image(filenameIN):
    '''Takes screenshot of created robot image'''
    '''   Considered switching to RPA.Browser which methods more easily allowed screenshots by element evidently'''
    loc = "#robot-preview-image"
    docID = filenameIN
    '''brwsr = Browser()
    brwsr.click_button("text=Preview")
    brwsr.capture_element_screenshot(loc, docID)'''

    page = browser.page()
    page.click("text=Preview")
    '''image = page.locator(loc).inner_html()'''
    page.screenshot(path=docID, full_page=True)
    '''page.screenshot(path=docID, full_page=True)'''
    '''page.screenshot(locator=loc, path=docID)'''
 
def store_receipt_as_pdf(ordnoIN, pageIN):
    '''Export the data to a pdf file'''
    docID = "output/receipts/" + ordnoIN + ".pdf"
    page = pageIN

    while True:
        try:
            receipt_html = page.locator("#receipt").inner_html()
            break
        except:
            print("Error in store_receipt_as_pdf for " + ordnoIN + ", re-clicking")
            page.click("button.btn.btn-primary")
   
    print(receipt_html)
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, docID)
    page.click("button.btn.btn-primary")
    close_annoying_modal()
    return docID

def embed_screenshot_to_receipt(screenshot, pdf_file):
    '''Merge the screenshot files together into the receipt'''
    pdf = PDF()
    list_of_files = [ screenshot ]
    pdf.add_files_to_pdf(files=list_of_files, target_document=pdf_file, append=True)

def archive_receipts():
    lib = Archive()
    lib.archive_folder_with_zip("output/receipts/", "receipts.zip")
