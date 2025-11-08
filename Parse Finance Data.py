from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException 
from selenium.webdriver.chrome.options import Options
import requests
from bs4 import BeautifulSoup
import pandas as pd
import random
import os

Company = ["6680 鑫創電子 (上櫃電腦及週邊設備業)", "8050 廣積 (上櫃電腦及週邊設備業)", "6922 宸曜 (上櫃電腦及週邊設備業)","3594 磐儀 (上櫃電腦及週邊設備業)", "2397 友通 (上市電腦及週邊設備業)",
           "8234 新漢 (上櫃電腦及週邊設備業)",  "2395 研華 (上市電腦及週邊設備業)", "6166 凌華 (上市電腦及週邊設備業)", "3088 艾訊 (上櫃電腦及週邊設備業)", "6579 研揚 (上市電腦及週邊設備業)", "3479 安勤 (上櫃電腦及週邊設備業)",
           "6414 樺漢 (上市電腦及週邊設備業)", "6570 維田 (上櫃電腦及週邊設備業)", "6245 立端 (上櫃通信網路業)", "3416 融通電 (上市電腦及週邊設備業)"]

# Set up headless Chrome options
options = Options()
options.add_argument("--headless")
options.add_argument("--disable-gpu")  # Disable GPU for headless mode
options.add_argument("--no-sandbox")   # Run without sandbox for CI environments
options.add_argument("--disable-dev-shm-usage")  # Prevent limited shared memory errors


# 自動安裝正確的 ChromeDriver 版本
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# Open the website
url = "https://mopsfin.twse.com.tw/"
driver.get(url)
time.sleep(3)  # Allow the page to load

# Dictionary to store results
dataframes = {}

# Loop over section IDs from a-2 to a-8
for i in range(2, 9):
    section_id = f"a-{i}"
    
    try:
        # Locate the section
        section = driver.find_element(By.ID, section_id)

        # Click to expand the section if not already open
        if "aria-expanded" in section.get_attribute("outerHTML"):
            expanded = section.get_attribute("aria-expanded")
            if expanded == "false":
                section.click()
                time.sleep(3)  # Wait for it to expand

        # Locate all buttons with the target class inside this section
        buttons = section.find_elements(By.CLASS_NAME, "compareItem")
        print(f"Found {len(buttons)} buttons in {section_id}")

        # Click each button
        for btn in buttons:
            button_name = btn.get_attribute("name")  # Get 'name' attribute
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(random.randint(3, 6))   # Wait for modal to load or response

             # 嘗試尋找 "setting-pop" 面板
            try:
                selection_panel = driver.find_element(By.ID, "setting-pop")
                if selection_panel.is_displayed():
                    print("Selection panel opened successfully!")

                    # 獲取公司輸入欄位
                    company_inputs = driver.find_elements(By.CLASS_NAME, "companyInput")
                    print(f"Found {len(company_inputs)} company input fields.")

                    # 限制填入數量（最多填入 10 個）
                    num_inputs = min(len(company_inputs), len(Company))

                    # 依序填入公司名稱
                    for i in range(num_inputs):
                        company_inputs[i].clear()
                        company_inputs[i].send_keys(Company[i])
                    
                    company_inputs[i].send_keys(Keys.ENTER)  # 按下 Enter 以選擇公司
                    time.sleep(3)  # 等待選擇生效

            except Exception as e:
                print("Selection panel not found or not displayed, proceeding to table extraction...")

            # Replace 'table-id' with the actual ID or class of the table you're waiting for
            table_locator = (By.ID, 'resultList1') 

            # Wait for up to 20 seconds for the table to become visible
            retries = 3
            for attempt in range(retries):
                try:
                    table = WebDriverWait(driver, 20).until(
                        EC.visibility_of_element_located(table_locator)
                    )
                    print("Table Shows Successfully!")
                    break  # If successful, exit the loop
                except TimeoutException:
                    if attempt < retries - 1:
                        print(f"Retrying ({attempt+1}/{retries})...")
                        time.sleep(2 ** attempt)  # Exponential backoff: wait longer each retry
                    else:
                        print("Failed after several retries.")

            # Extract the page source from the driver
            soup = BeautifulSoup(driver.page_source, "html.parser")

            # Find the table by its ID
            headtable = soup.find('table', {'id': 'resultList1'})

            # Extract the header row from the <thead> section 
            header_row = headtable .find('thead').find_all('th')

            # Extract the header values (strip and store the text)
            headers = [header.text.strip() for header in header_row]

            # Extract the body rows from the <tbody> section
            body_rows = headtable.find('tbody').find_all('tr')

            # Extract the body values (strip and store the text)
            body = [row.text.strip()  for row in body_rows]

            # Convert the data into a DataFrame
            df1 = pd.DataFrame(body, columns=headers)



            # Find the table by its ID
            table = soup.find('table', {'id': 'resultList2'})

            # Extract the header row from the <thead> section 
            header_row = table.find('thead').find_all('td')

            # Extract the header values (strip and store the text)
            headers = [header.text.strip() for header in header_row]

            # Extract the body rows from the <tbody> section
            body_rows = table.find('tbody').find_all('tr')

            # Extract the body values (strip and store the text)
            body = [[cell.text.strip() for cell in row.find_all('td')] for row in body_rows]

            # Convert the data into a DataFrame
            df2 = pd.DataFrame(body, columns=headers)


            # Assuming df is another DataFrame
            output = pd.merge(df1, df2, left_index=True, right_index=True)
            dataframes[button_name] = output

            # Create a list to store DataFrames with an added "Category" column
            df_list = []    

            for name, df in dataframes.items(): 
                df = df.reset_index(drop=True)  # 刪除原本的索引
                df["Category"] = name  # 新增 Category 欄位
                
                if "年度/季度" in df.columns:  
                    df = df.rename(columns={"年度/季度": "公司"})  # 重新命名欄位
                
                df_list.append(df)  # 加入列表

                # Concatenate all DataFrames
            final_df = pd.concat(df_list, ignore_index=True)

    except Exception as e:
        print(f"Error processing {section_id}: {e}")

# Close the driver after execution
driver.quit()

# Ensure the 'finance_data' directory exists
directory = "finance_data"
if not os.path.exists(directory):
    os.makedirs(directory)

# Define the file path to store the CSV
file_name = os.path.join(directory, "selected_companies_data.csv")

# Check if the file already exists
if os.path.exists(file_name):
    # Read existing data
    existing_df = pd.read_csv(file_name)
    
    # Find new rows that are not in existing_df
    combined_df = pd.concat([existing_df, final_df], ignore_index=True)
    
    # Drop duplicate rows based on specific columns (adjust columns as needed)
    combined_df = combined_df.drop_duplicates(subset=["公司", "Category"], keep="last")
else:
    # If file doesn't exist, use final_df as the initial dataset
    combined_df = final_df

# Save the updated DataFrame as a CSV
combined_df.to_csv(file_name, index=False)
print(f"Data updated successfully in {file_name}")