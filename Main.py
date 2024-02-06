from ast import While
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from bs4 import BeautifulSoup
from time import sleep
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import datetime
import tkinter as tk
from tkinter import FIRST, messagebox
import requests


class WebScraper:
    def __init__(self):
        # Initialize the webdriver
        self.firefox_options = Options()
        self.firefox_options.headless = True
        self.driver = webdriver.Firefox(options=self.firefox_options)
        self.actions = ActionChains(self.driver)
        self.data_export = []

    def close_browser(self):
        self.driver.quit()

    def navigate_to_url(self, url):
        self.driver.get(url)
        sleep(10)  # Adjust sleep as needed

    def get_data_from_Firstpage(self):
        # Now that the content is loaded, get the HTML
        print('Render the dynamic content to static HTML')
        html = self.driver.page_source       
        print(' Parse the static HTML')
        soup = BeautifulSoup(html, "html.parser")
        for a_tag in soup.find_all('a', class_='pageResult--Title'):
            name = a_tag.text.strip()
            name = name.split(' - ')[0]
            link = a_tag['href']
            print("name: ", name)

            ## For each Profile:
            response = requests.get(link)
            if response.status_code == 200:
                soup_2 = BeautifulSoup(response.text, 'html.parser')
                # Check if the <i> tag with class "fa fa-users" exists
                users_icon = soup_2.find('i', class_='fa fa-users')
                if users_icon:
                    print("Type Profile Full")
                    info2_dict = {}
                    info2_dict["Name: "] = name

                    # Find the <h2> tag with the specified class
                    h2_tag = soup_2.find('h2', class_='hero-header font_museo500')
                    p_tag = soup_2.find('p', class_='btn-rg')
                    if h2_tag:
                        text_content = h2_tag.text.strip()
                        p_tag_text = p_tag.text.strip()
                        text_content_write = text_content + "\n" + p_tag_text
                        info2_dict['Position'] = text_content_write
                    else:
                        info2_dict['Position'] = "None"
                        print("The specified <h2> tag was not found.")

                    # Extract School / Department information
                    # Find the <p> tag with the <i> tag containing "fa fa-users"
                    school_department_tag = soup_2.find(lambda tag: tag.name == 'p' and tag.find('i', class_='fa fa-users') is not None)
                    if school_department_tag:
                        # Extract text from all <a> tags within the found <p> tag
                        links_text = ' | '.join(a.get_text(strip=True) for a in school_department_tag.find_all('a'))
                        
                        # Add the concatenated text to the dictionary
                        info2_dict["School / Department"] = links_text

                    # Extract Email information
                    email_tag = soup_2.find(lambda tag: tag.name == 'p' and tag.find('i', class_='fa fa-envelope-o') is not None)
                    if email_tag:
                        email_text = email_tag.find('a').get('href').replace('mailto:', '')
                        info2_dict["Email"] = email_text
                    self.data_export.append(info2_dict)


                else:
                    print("Type Profile Short")
                    # Initialize an empty dictionary to store the extracted information
                    info_dict = {}
                    info_dict["Name: "] = name

                    # Define the desired fields and corresponding tags
                    fields = ["Position", "School / Department", "Email"]

                    for field in fields:
                        # Find the <span> tag with the corresponding text
                        span_tag = soup_2.find('span', text=field + ':')
                        
                        # Check if the span_tag is found
                        if span_tag:
                            # Extract the text from the next sibling <span> tag
                            value = span_tag.find_next('span').text.strip()
                            # Add the field and value to the dictionary
                            info_dict[field] = value
                        else:
                            print("Fail catch")
                    self.data_export.append(info_dict)



            else:
                print(f"Failed to retrieve the page. Status code: {response.status_code}")


    def export_to_excel(self, file_name='output.xlsx'):
        # Create a DataFrame from the list of dictionaries
        df = pd.DataFrame(self.data_export)

        # Export DataFrame to Excel
        df.to_excel(file_name, index=False)
        print(f'Data exported to {file_name}')

# Example usage:
if __name__ == "__main__":
    ##Filter the option in Marketplace and copy the url
    scraper = WebScraper()
    url1 = "https://www.rmit.edu.au/search?searchtype=contacts&q=%22RMIT+University%22"
    scraper.navigate_to_url(url1)
    scraper.get_data_from_Firstpage()

    for i in range(2,20):
        url2 = "https://www.rmit.edu.au/search?q=%22RMIT+University%22&searchtype=contacts&current=" + str(i) + "&size=10"
        scraper.navigate_to_url(url2)
        scraper.get_data_from_Firstpage()
    
    # Export the data to Excel after the loop
    scraper.export_to_excel()

    scraper.close_browser()
