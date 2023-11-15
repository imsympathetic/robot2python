from robocorp.tasks import task
from robocorp import browser
from robocorp import http
from RPA.Excel.Files import Files
from RPA.PDF import PDF
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from RPA.Tables import Tables
import csv
import time
from RPA.Browser.Selenium import Selenium
import re
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
        slowmo=500,
    )
    download_excel_file()
    open_robot_order_website()
    fill_order_using_data_from_excel()
    #create_zip_file()
    #close_browser()
    archive_receipts()

def open_robot_order_website():
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    time.sleep(1)

def download_excel_file():
    http.download("https://robotsparebinindustries.com/orders.csv", overwrite=True)

def fill_order_using_data_from_excel():
    tables = Tables()
    orders = tables.read_table_from_csv("orders.csv")
    print(orders)

    for order in orders:
        fill_order_for_one_robot(order)

def fill_order_for_one_robot(order):
    # Set Selenium Speed: 0.2 sec
    page = browser.page()

    page.click("text=OK")
    
    page.select_option("#head", order["Head"])
    
    page.check(selector="#id-body-"+order["Body"])
    
    page.fill('//label[text()="3. Legs:"]/ancestor::div/input',order["Legs"])
    page.fill("#address", order["Address"])


    click_order()
    receipt_pdf = store_order_receipt_as_pdf(order)
    screenshot = screenshot_robot(order)
    embed_screenshot_to_pdf(screenshot, receipt_pdf)
    page.click("#order-another")
    time.sleep(1)

def click_order():
        page = browser.page()
        page.click("#order")
        retry_count=0

        while page.is_visible(selector='//div[@class="alert alert-danger"]',timeout=60) and retry_count <3:
            page.click("#order")
            retry_count=retry_count + 1
            time.sleep(2)

def store_order_receipt_as_pdf(order):
    page=browser.page()
    receipt_html=page.locator("#receipt").inner_html()
    pdf_path=f"./output/receipts/receipt" + order["Order number"] + ".pdf"
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, pdf_path)
    return pdf_path

def screenshot_robot(order):
    page = browser.page()
    screenshot_path="./output/receipts/screenshot"+order["Order number"]+ ".jpg"
    page.locator("#robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path


def embed_screenshot_to_pdf(screenshot, pdf_file):
    """Add screenshot to the pdf file"""
    pdf = PDF()
    pdf.add_files_to_pdf(files=[screenshot], target_document=pdf_file, append=True)

 

def close_browser():
    driver = webdriver.Chrome()
    driver.get('https://robotsparebinindustries.com/#/robot-order')
    driver.quit()

def archive_receipts():
    """ZIP receipt files"""
    lib = Archive()
    lib.archive_folder_with_zip('output/receipts', 'receipts.zip')