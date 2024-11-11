import logging
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time, os, telebot, threading
from decouple import config
import pandas as pd
from openpyxl import Workbook
from settings import CITIES, URL
from openpyxl.styles import Font, Alignment, PatternFill

# Initialize logger
def setup_logger():
    logger = logging.getLogger("JobScraperLogger")
    logger.setLevel(logging.DEBUG)

    file_handler = logging.FileHandler("job_scraper.log")
    file_handler.setLevel(logging.DEBUG)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger

logger = setup_logger()
token = config("TELEGRAM_TOKEN")
bot = telebot.TeleBot(token=token)
chat_id_storage = "chat_ids.txt"

# Start bot polling in a separate thread
def start_bot_polling():
    bot.infinity_polling()

# Select multiple cities
def select_cities(cities):
    try:
        jobSearchDropdown = WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "jobSearchFilters"))
        )
        time.sleep(3)
        select = Select(jobSearchDropdown)
        
        if select.is_multiple:
            for city in cities:
                select.select_by_value(value=city)
            selected_options = select.all_selected_options
            for option in selected_options:
                logger.info(f"Selected option: {option.get_attribute('value')}")
        else:
            logger.warning("Dropdown does not support multiple selections")
    except Exception as e:
        logger.error(f"Error in select_cities function: {e}")

# Extract job information
def extract_jobs(excel_file):
    jobs = []
    try:
        WebDriverWait(browser, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'job-list-item'))
        )
        if os.path.exists(excel_file):
            existing_jobs_df = pd.read_excel(excel_file, sheet_name="Jobs List")
            existing_req_ids = set(existing_jobs_df["Req ID"].astype(str))
        else:
            existing_req_ids = set()
        
        time.sleep(2)
        job_items = browser.find_elements(By.CLASS_NAME, "job-list-item")
        for job in job_items:
            title = job.find_element(By.TAG_NAME, "h3").text
            link = job.find_element(By.TAG_NAME, "a").get_attribute("href")
            address = job.find_element(By.TAG_NAME, "h4").text
            req_id = job.find_element(By.CLASS_NAME, "job-attribute").text.strip()
            if req_id not in existing_req_ids:
                jobs.append({
                    "Req ID": req_id,
                    "Title": title,
                    "Link": link,
                    "Address": address
                })
        logger.info(f"Extracted {len(jobs)} job listings.")
    except Exception as e:
        logger.error(f"Error extracting job details: {e}")
    
    return jobs

def create_excel(jobs, excel_file):
    new_jobs_df = pd.DataFrame(jobs)

    if os.path.exists(excel_file):
        existing_jobs_df = pd.read_excel(excel_file, sheet_name="Jobs List")
        combined_df = pd.concat([existing_jobs_df, new_jobs_df]).drop_duplicates(subset="Req ID", keep="first")
    else:
        combined_df = new_jobs_df

    with pd.ExcelWriter(excel_file, engine="openpyxl", mode="w") as writer:
        combined_df.to_excel(writer, index=False, sheet_name="Jobs List")
        worksheet = writer.sheets['Jobs List']

        header_font = Font(bold=True, color="FFFFFF")
        fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        alignment = Alignment(horizontal="center", vertical="center")

        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = fill
            cell.alignment = alignment

        for column in worksheet.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
    
    logger.info(f"Excel file '{excel_file}' updated successfully.")

@bot.message_handler(commands=['start'])
def initialize_bot(message):
    chat_id = message.chat.id
    with open(chat_id_storage, "a") as file:
        file.write(f"{chat_id}\n")
    bot.send_message(chat_id, "Welcome! You will receive job alerts from home depot automatically.")

def get_all_chat_ids():
    if os.path.exists(chat_id_storage):
        with open(chat_id_storage, "r") as file:
            chat_ids = {line.strip() for line in file.readlines()}
    else:
        chat_ids = set()
    return chat_ids

def send_job_updates(jobs):
    chat_ids = get_all_chat_ids()
    if not chat_ids:
        logger.info("No users have subscribed to job updates.")
        return

    for job in jobs:
        job_message = (
            f"📌 *{job['Title']}*\n\n"
            f"📍 *Address:* {job['Address']}\n"
            f"🔗 [View Job Posting]({job['Link']})\n\n"
            f"🆔 *Req ID:* {job['Req ID']}"
        )
        for chat_id in chat_ids:
            try:
                bot.send_message(chat_id, job_message, parse_mode="Markdown")
            except Exception as e:
                logger.error(f"Failed to send job notification to {chat_id}: {e}")

if __name__ == "__main__":
    # Start bot polling in a separate thread
    bot_thread = threading.Thread(target=start_bot_polling)
    bot_thread.start()

    # Start scraping loop
    while True:
        browser = webdriver.Chrome()
        browser.get(url=URL)
        cities = CITIES
        select_cities(cities=cities)
        excel_file = "jobs_list.xlsx"
        jobs = extract_jobs(excel_file)
        if jobs:
            create_excel(jobs, excel_file)
            send_job_updates(jobs)
        else:
            logger.warning("No jobs found in the given cities.")
        browser.quit()
        time.sleep(60)  # Run the scraping process every 1 Minute
