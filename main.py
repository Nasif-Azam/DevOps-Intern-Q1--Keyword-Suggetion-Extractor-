import time
import datetime
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

today = datetime.date.today()
today_name = today.strftime("%A")
sheet = pd.read_excel('Excel File/4BeatsQ1.xlsx', sheet_name=today_name)
print(f"Today is: {today_name}. Sheet Information: \n{sheet}")
keywords = sheet['Keywords Name'].values
print(f"The Keywords Are: {keywords}")

##### Set up WebDriver for Brave Browser #####
options = webdriver.ChromeOptions()
options.binary_location = r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe"
chrome_driver_path = r'C:\Users\mdnas\.cache\selenium\chromedriver\win64\128.0.6613.84\chromedriver.exe'
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)
print(driver.title)


##### Function to fetch Google suggestions #####
def getSuggestions(keyword):
    driver.get('https://www.google.com/?hl=en')
    search_box = driver.find_element(By.NAME, 'q')
    search_box.send_keys(keyword)
    time.sleep(2)
    suggestions = driver.find_elements(By.CSS_SELECTOR, '.wM6W7d')
    suggestionList = []
    for suggestion in suggestions:
        if len(suggestion.text) != 0:
            suggestionList.append(suggestion.text)
    return suggestionList


##### Store Shortest and Longest Suggestions #####
shortest_list = []
longest_list = []
for i, keyword in enumerate(keywords):
    all_suggestions = getSuggestions(keyword)
    if all_suggestions:
        shortest = min(all_suggestions, key=len)
        longest = max(all_suggestions, key=len)
        shortest_list.append(shortest)
        longest_list.append(longest)
        print(f"Keyword({i + 1}): {keyword}")
        print(f"Shortest Suggestions: {shortest}")
        print(f"Longest Suggestions: {longest}")
    else:
        print(f"No suggestions found for {keyword}")

##### Update Shortest & Longest Suggestions #####
sheet['Shortest Option'] = shortest_list
sheet['Longest Option'] = longest_list

##### Save To Another Excel File #####
# sheet.to_excel('New Excel.xlsx', sheet_name=today_name, index=False)
##### Save To Existing Excel File #####
with pd.ExcelWriter('Excel File/4BeatsQ1.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    sheet.to_excel(writer, sheet_name=today_name, index=False)

input("Press Enter To Close The Browser: ")
