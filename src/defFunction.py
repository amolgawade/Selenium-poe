from openpyxl import Workbook
from selenium.common import TimeoutException
from selenium.webdriver import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from openpyxl.worksheet import worksheet
from selenium import webdriver
from selenium.webdriver.common.by import By
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
        self.driver.maximize_window()
        # Set implicit wait to 10 seconds
        self.driver.implicitly_wait(10)

    def navigate_to_website(self, url):
        # Navigate to the website
        self.driver.get(url)

    def click_use_email_button(self):
        self.driver.find_element(*HomePageLocators.email_but).click()

    def click_on_email_box(self):
        # locate email input box
        self.driver.find_element(*HomePageLocators.email_input).click()
        self.driver.find_element(*HomePageLocators.email_input).send_keys("shidorifoodsonline@gmail.com")

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

        # Load the Excel file
        workbook = openpyxl.load_workbook(file_name)
        worksheet = workbook['Sheet1']
        print("Opening excel file" + file_name)
        max_row = worksheet.max_row
        self.driver.find_element(*HomePageLocators.question_area).click()

        # Loop through each row in the Excel file
        for row_number in range(1, max_row + 1):
            try:
                self.put_question(row_number, worksheet)

                # Wait for the like button to be visible
                index = self.wait_like_button(row_number)

                # Wait for the hover to ans box
                self.hover_ans_box(index)

                # click to drop down menu
                self.click_drop_down(index)

                # Click on the copy button
                self.click_copy_butt(index)

                # Copy data from clipboard
                clipboard, copied_data = self.copy_clipboard()

                # Write data to Excel file
                self.save_ans(copied_data, row_number, worksheet)
                # Reset clipboard
                clipboard.copy('')

                # Save the Excel file
                workbook.save(file_name)
                print("saving the answer " + str(row_number) + " to the file")
            except Exception as e:
                print("Error occurred for row " + str(row_number) + ": " + str(e))

        print("Saving All the answers of Questions " + str(max_row) + " in the file name " + file_name)

    def save_ans(self, copied_data, row_number, worksheet):
        worksheet.cell(row=row_number, column=2).value = copied_data
        print("Pasting the answer to row " + str(row_number))

    def copy_clipboard(self):
        clipboard = pyperclip
        copied_data = clipboard.paste()
        return clipboard, copied_data

    def click_copy_butt(self, index):
        wait = WebDriverWait(self.driver, timeout=20)
        copy_but = "//*[@id='__next']/div[1]/section/div/div/div[{}]/div[2]/div[3]/div/div/div/div/button[3]".format(
            index)
        copy = wait.until(EC.visibility_of_element_located((By.XPATH, copy_but)))
        time.sleep(2)
        copy.click()
        print("copying text....")

    def click_drop_down(self, index):
        time.sleep(1)
        wait = WebDriverWait(self.driver, timeout=20)
        hover_xpath = "//*[@id='__next']/div[1]/section/div/div/div[{}]/div[2]/div[3]/div/div/button/div".format(index)
        hover = wait.until(EC.visibility_of_element_located((By.XPATH, hover_xpath)))
        time.sleep(1)
        actions = ActionChains(self.driver)
        actions.move_to_element(hover).click().perform()

    def hover_ans_box(self, index):
        wait = WebDriverWait(self.driver, timeout=20)
        hover_box = "//*[@id='__next']/div[1]/section/div/div/div[{}]/div[2]/div[2]/div[1]/div[1]".format(index)
        box = wait.until(EC.visibility_of_element_located((By.XPATH, hover_box)))
        actions = ActionChains(self.driver)
        actions.move_to_element(box).click().perform()

    def wait_like_button(self, row_number):
        wait = WebDriverWait(self.driver, timeout=60)
        index = 0 + row_number
        wait.until(EC.visibility_of_element_located(
            (By.XPATH, "//*[@id='__next']/div[1]/section/div/div/div[{}]/section/button[2]".format(index))))
        return index

    def put_question(self, row_number, worksheet):
        question = worksheet.cell(row=row_number, column=1).value
        print("Asking question no " + str(row_number) + " to Sage")
        # Enter the question in the text area and click on send
        self.driver.find_element(*HomePageLocators.question_area).send_keys(question)
        self.driver.find_element(*HomePageLocators.que_send_butt).click()
