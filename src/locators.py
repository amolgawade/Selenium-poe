
from selenium.webdriver.common.by import By


class HomePageLocators:
    email_but = (By.XPATH, "//*[@id='__next']/main/div/button[2]")
    email_input = (By.XPATH, "//*[@id='__next']/main/div/div[2]/form/input")
    go_butt = (By.XPATH, "//*[@id='__next']/main/div/button[1]")
    login_butt = (By.XPATH, "//*[@id='__next']/main/div/button[2]")
    que_send_butt = (By.XPATH, "//*[@id='__next']/div[1]/section/footer/div/div[3]/button")
    copy_butt = (By.XPATH, "//*[@id='__next']/div[1]/section/div/div/div[4]/div[2]/div[3]/div/div/div/div/button[3]")
    hov_ele = (By.XPATH, "//*[@id='__next']/div[1]/section/div/div/div[4]/div[2]/div[3]/div/div/button/div/svg")
    question_area = (By.XPATH, "//*[@id='__next']/div[1]/section/footer/div/div[2]/textarea")

