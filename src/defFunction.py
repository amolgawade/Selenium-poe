from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from openpyxl.worksheet import worksheet
from selenium import webdriver
from selenium .webdriver.common.by import By
import time
from selenium.webdriver.support.wait import WebDriverWait
from locators import HomePageLocators
import openpyxl
import pyperclip

file_name = "F:\\upsc_questions.xlsx"  # Note the double backslashes or use forward slashes


class PoeSite:
    def __init__(self):
        # Set up Chrome driver
        self.clipboard = None
        self.driver = webdriver.Chrome()

    def navigate_to_website(self, url):
        # Navigate to the website
        self.driver.get(url)

    def click_use_email_button(self):
        self.driver.find_element(*HomePageLocators.email_but).click()

    def click_on_email_box(self):
        # locate email input box
        self.driver.find_element(*HomePageLocators.email_input).click()
        self.driver.find_element(*HomePageLocators.email_input).send_keys("legendneverdie1947@gmail.com")

    def click_go_button(self):
        # Locate and click the go button
        self.driver.find_element(*HomePageLocators.go_butt).click()

    def click_login_button(self):
        self.driver.find_element(*HomePageLocators.login_butt).click()

    def wait(self, seconds):
        # Wait for the specified number of seconds
        time.sleep(1)

    def close_browser(self):
        # Close the browser
        self.driver.quit()

    def ask_questions(self):


        time.sleep(5)
        text_area = self.driver.find_element(*HomePageLocators.question_box).click()
        workbook = openpyxl.load_workbook(file_name)
        worksheet = workbook['Sheet1']
        max_row = worksheet.max_row

        for row_number in range(1, max_row):
            question = worksheet.cell(row=row_number, column=1).value

            self.driver.find_element(*HomePageLocators.question_box).send_keys(question)
            self.driver.find_element(*HomePageLocators.question_send).click()
            time.sleep(10)

            # wait for visible the like button
            wait = WebDriverWait(self.driver, timeout=100)
            index = 0 + row_number
            wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='__next']/div[1]/section/div/div/div[{}]/section/button[2]".format(index))))

            # hover to element
            time.sleep(0.5)
            element = self.driver.find_element(By.XPATH,"//*[@id='__next']/div[1]/section/div/div/div[{}]/div[2]/div[3]/div/div/button/div".format(index))
            actions = ActionChains(self.driver)
            actions.move_to_element(element).click()

            # copy button
            time.sleep(0.5)
            data = self.driver.find_element(By.XPATH, "//*[@id='__next']/div[1]/section/div/div/div[{}]/div[2]/div[3]/div/div/div/div/button[3]".format(index)).click()
            
            # copy data to clipboard
            self.clipboard = pyperclip
            self.clipboard.copy(data)




