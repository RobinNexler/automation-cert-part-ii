import logging
import os
import shutil
from io import BytesIO

import pandas as pd
import requests
from robocorp import browser
from robocorp.tasks import task
from RPA.Archive import Archive
from RPA.PDF import PDF

# Configure browser to run in slow motion
browser.configure(
    screenshot="only-on-failure",
    headless=False,
    slowmo=100,  # 100ms delay between actions
)

ORDER_URL: str = "https://robotsparebinindustries.com/#/robot-order"
CSV_URL: str = "https://robotsparebinindustries.com/orders.csv"
MAX_RETRIES: int = 3


@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    try:
        orders = get_orders()
        fill_the_form_and_store_receipts(orders)
        archive_receipts()
    finally:
        cleanup()


def archive_receipts() -> None:
    """
    Create a ZIP archive of the receipt PDFs.

    This function creates a ZIP archive of all receipt PDFs stored in the output/receipts directory.
    The archive is placed in the output directory.
    """
    archive = Archive()
    archive.archive_folder_with_zip(
        folder="temp/receipts",
        archive_name="output/receipts.zip",
        recursive=True,
        include="*.pdf",
    )
    logging.info("PDF receipts archived to output/receipts.zip")


def fill_the_form_and_store_receipts(orders: pd.DataFrame) -> None:
    """
    Iterate through each order and fill out the robot order form.

    Args:
        orders: DataFrame containing robot order details from the CSV file
    """
    for _, row in orders.iterrows():
        os.makedirs("temp/receipts", exist_ok=True)
        os.makedirs("temp/receipts_images", exist_ok=True)
        browser.goto(ORDER_URL)
        close_annoying_modal()
        order_number = fill_the_form_for_one_order(row)
        image_path = save_receipt_as_image(order_number=order_number)
        store_receipt_as_pdf(order_number=order_number, image_path=image_path)
        browser.page().close()


def fill_the_form_for_one_order(row: pd.Series) -> str:
    """
    Fill out the form for a single robot order.

    Args:
        row: A pandas Series containing details for a single robot order
    """
    page = browser.page()
    # Select head from dropdown
    page.select_option("//select[@id='head']", str(row["Head"]))

    # Select body using radio buttons
    page.check(f"//input[@id='id-body-{row['Body']}']")

    # Enter legs in number input field
    page.fill("//div[contains(label, 'Legs')]//input", str(row["Legs"]))

    # Enter address in text field
    page.fill("//input[@id='address']", row["Address"])

    # Click the Order button
    click_order_button_with_retry(row)

    order_number = page.text_content("//p[@class='badge badge-success']")
    return order_number


def click_order_button_with_retry(row: pd.Series, count: int = 0) -> None:
    """
    Click the Order button with retry logic.
    """
    page = browser.page()
    page.click("//button[@id='order']")

    if count > MAX_RETRIES:
        logging.error(f"Failed to click order button for order {row}")
        return

    if page.locator("//div[@class='alert alert-danger']").is_visible():
        click_order_button_with_retry(count + 1)


def save_receipt_as_image(order_number: str) -> str:
    """
    Save a screenshot of the robot image.

    Args:
        order_number: The number of the order

    Returns:
        The path to the robot image file
    """
    page = browser.page()
    output_path = f"temp/receipts_images/robot_{order_number}.png"
    page.locator("//div[@id='robot-preview-image']").screenshot(path=output_path)
    return output_path


def store_receipt_as_pdf(order_number: str, image_path: str) -> None:
    """
    Store the HTML receipt as a PDF file and embed the robot image.

    Args:
        order_number: The number of the order
        image_path: The path to the robot image file
    """
    page = browser.page()
    pdf_path = f"temp/receipts/receipt_{order_number}.pdf"
    pdf_path_with_robot_image = (
        f"temp/receipts/receipt_{order_number}_with_robot_image.pdf"
    )

    # Get the HTML receipt element
    receipt_html = page.locator("//div[@id='receipt']").inner_html()

    # Create a PDF from the HTML receipt
    pdf = PDF()
    pdf.html_to_pdf(receipt_html, pdf_path)

    # Add the robot image to the PDF
    pdf.add_watermark_image_to_pdf(
        image_path=image_path,
        source_path=pdf_path,
        output_path=pdf_path_with_robot_image,
    )
    logging.info(f"Receipt for order {order_number} saved as PDF with robot image")


def get_orders() -> pd.DataFrame:
    """
    Downloads an Excel file from a remote URL and returns its contents as a pandas DataFrame.

    This function fetches the Excel file directly from the web, loads it into memory,
    and parses it without writing anything to disk. As a result, no temporary or permanent
    files are created. making it efficient and free of file system artifacts.

    Returns:
        pandas.DataFrame: A DataFrame containing the contents of the downloaded Excel file.

    Raises:
        requests.HTTPError: If the file could not be retrieved due to a network or server error.
    """
    response = requests.get(CSV_URL)
    response.raise_for_status()

    excel_data = pd.read_csv(BytesIO(response.content))
    return excel_data


def close_annoying_modal() -> None:
    """
    Close the consent dialog that appears when the page loads.
    Clicks the 'OK' button to dismiss the modal.
    """
    page = browser.page()
    page.click("//button[@class='btn btn-dark']")


def cleanup() -> None:
    """
    Remove temporary folders and files after the process is complete.
    """
    try:
        if os.path.exists("temp"):
            shutil.rmtree("temp")
            logging.info("Temporary files cleaned up successfully")
    except Exception as e:
        logging.error(f"Error cleaning up temporary files: {e}")


if __name__ == "__main__":
    cleanup()
