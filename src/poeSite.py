import time
from defFunction import PoeSite


# Create an instance of the class
poe = PoeSite()

print("navigate_to_website : https://poe.com/ ")
# Navigate to the website
poe.navigate_to_website("https://poe.com/")
print("------------*****************-------------------- ")
time.sleep(3)

# Click the use email option
poe.click_use_email_button()
print("Using email option ")
time.sleep(1)

print("------------*****************-------------------- ")
# Click the use email option send email
poe.click_on_email_box()
time.sleep(1)

# Click the email verification button
print("sending email")
poe.click_go_button()
time.sleep(25)

print("------------*****************-------------------- ")
# Click to verify email OTP
print("verify email")
poe.click_login_button()
print("------------*****************-------------------- ")
print("Login to account")
time.sleep(3)
print("------------*****************-------------------- ")
poe.ask_questions()
print("Good work all the answer are copied and save to file..... END")

