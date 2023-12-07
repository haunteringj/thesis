import os
import time
import tabula
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# function to extract tables from pdf files
def tabula_extraction():
    # directories for input and ouptut
    pdf_directory = '/Users/jay/Downloads'
    output_excel_file = 'combined_tables.xlsx'

    # create pandas excel writers to write different pdfs to different worksheets
    excel_writer = pd.ExcelWriter(output_excel_file, engine='xlsxwriter')

    # iterate through pdf files and extract tables
    for pdf_file in os.listdir(pdf_directory):
        if pdf_file.startswith("Event_") and pdf_file.endswith(".pdf"):
            pdf_file_path = os.path.join(pdf_directory, pdf_file)

            # tabula extracts tables from pdf
            tables = tabula.read_pdf(pdf_file_path, pages='all')

            # combine pdfs into combined dataframe
            combined_data = pd.DataFrame()

            # save combined data into single excel file
            for i, table in enumerate(tables):
                df = pd.DataFrame(table)
                combined_data = pd.concat([combined_data, df], ignore_index=True)

                sheet_name = pdf_file.split('.')[0]
                combined_data.to_excel(excel_writer, sheet_name=sheet_name, index=False)

    excel_writer.close()

    print("Extraction process complete.")     
    return  

# selenium table extractor 
def extract_eventId(driver): 
    # handle cookies and close popup
    time.sleep(2)
    cookies_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "img[src='assets/imgs/close.svg']"))
    )
    cookies_element.click()

    # !!! HARDCODED NUMBER OF REPORTS TO DOWNLOAD !!!
    num_reports = 2

    try: 
        # download n reports
        for i in range(num_reports):
        
            driver.execute_script("window.scrollTo(0, 0);")

            # click export element cog and download event list
            export_event_list = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"/html/body/app-root/div/app-public-interface/section/app-in-event-list-pi/div/app-in-event-list/div/mat-drawer-container/mat-drawer-content/div/div[2]/app-custom-table/div/div[1]/div[3]"))
            )
            export_event_list.click()
            download_event_list = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, f"/html/body/div[3]/div[2]/div/div/div/app-custom-export-button/button"))
            )
            download_event_list.click()

            # iterate to next page and continue downloading event lists
            next_page = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, f"/html/body/app-root/div/app-public-interface/section/app-in-event-list-pi/div/app-in-event-list/div/mat-drawer-container/mat-drawer-content/div/div[2]/mat-paginator/div/div/div[2]/button[2]"))
            )
            next_page.click()
        
    finally:
        download_dir = '/Users/jay/Downloads'

        # get csvs containing Event-lists 
        event_ids = []
        pattern = 'Event-list-*.csv'

        # get event lists from downloads file and iterate through them for event ids
        downloaded_files = os.listdir(download_dir)
        for file in downloaded_files:
            if file.endswith('.csv') and file.startswith('Event-list-'):
                file_path = os.path.join(download_dir, file)

                # read csv for event ids
                with open(file_path, 'r') as file:
                    lines = file.readlines()
                
                # put event id rows into dataframe 
                rows = [line.strip().split(';') for line in lines]
                df = pd.DataFrame(rows[1:], columns=rows[0])

                # put event id dataframe into list
                if 'eventId' in df.columns:
                    event_ids.extend(df['eventId'].tolist())
                else:
                    print("Column 'eventId' not found in file:", file)
        print('EventID extraction complete')
        return event_ids

# selenium downloads event pdf reports 
def download_event_PDF(driver, event_ids):

    # iterate through event ids to download repords 
    for event in event_ids:
        driver.get(f"https://wahis.woah.org/#/in-review/{event}?fromPage=event-dashboard-url")
    
        # click download btn
        download_report = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Download')]"))
            )
        download_report.click()

        # export page as pdf
        export_report = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Export')]"))
            )
        export_report.click()
        time.sleep(2)
    print('Event Reports download complete')


# start selenium webdriver
driver = webdriver.Chrome()
driver.get("https://wahis.woah.org/#/event-management")

# gather list of eventIds
event_list = extract_eventId(driver)

# download event PDFs using eventIds
download_event_PDF(driver, event_list)

# closer driver
driver.quit()

# extract tabular information
tabula_extraction()


# Other lasting issues:
# - delete downloaded reports and csvs after extraction
# - change num of reports to accept user input
# - make table detection broader
# - make extraction nicer