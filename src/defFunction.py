from openpyxl import Workbook
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
         # Create a Workbook instance
        workbook = Workbook()
        sheet = workbook.active

        # Load the Excel file
        workbook = openpyxl.load_workbook(file_name)
        worksheet = workbook['Sheet1']
        max_row = worksheet.max_row
        self.driver.find_element(By.XPATH, "//*[@id='__next']/div[1]/section/footer/div/div[2]/textarea").click()
        # Loop through each row in the Excel file
        for row_number in range(1, max_row + 1):
            question = worksheet.cell(row=row_number, column=1).value

            # Click on the text area


            # Enter the question in the text area
            self.driver.find_element(By.XPATH, "//*[@id='__next']/div[1]/section/footer/div/div[2]/textarea").send_keys(question)

            # Click on the send button
            self.driver.find_element(By.XPATH, "//*[@id='__next']/div[1]/section/footer/div/div[3]/button").click()

            time.sleep(10)

            # Wait for the like button to be visible
            wait = WebDriverWait(self.driver, timeout=100)
            index = 0 + row_number
            wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='__next']/div[1]/section/div/div/div[{}]/section/button[2]".format(index))))

            wait = WebDriverWait(self.driver, timeout=20)
            hover_box = "//*[@id='__next']/div[1]/section/div/div/div[{}]/div[2]/div[2]/div[1]/div[1]".format(index)
            box = wait.until(EC.visibility_of_element_located((By.XPATH, hover_box)))
            actions = ActionChains(self.driver)
            actions.move_to_element(box).click().perform()

            # hover to drop dowm menu
            wait = WebDriverWait(self.driver, timeout=20)
            hover_xpath = "//*[@id='__next']/div[1]/section/div/div/div[{}]/div[2]/div[3]/div/div/button/div".format(index)
            hover = wait.until(EC.visibility_of_element_located((By.XPATH, hover_xpath)))

            actions = ActionChains(self.driver)
            actions.move_to_element(hover).click().perform()

            # Click on the copy button
            wait = WebDriverWait(self.driver, timeout=20)
            copy_but = "//*[@id='__next']/div[1]/section/div/div/div[{}]/div[2]/div[3]/div/div/div/div/button[3]".format(index)
            copy = wait.until(EC.visibility_of_element_located((By.XPATH, copy_but)))
            copy.click()

            # Copy data from clipboard
            clipboard = pyperclip
            copied_data = clipboard.paste()

            # Write data to Excel file
            worksheet.cell(row=row_number, column=2).value = copied_data
            # Reset clipboard
            clipboard.copy('')

        # Save the Excel file
        workbook.save(file_name)





