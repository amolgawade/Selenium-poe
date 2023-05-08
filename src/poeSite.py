import time
from defFunction import PoeSite


# Create an instance of the class
poe = PoeSite()

# Navigate to the website
#poe.navigate_to_website("https://www.google.com/")
poe.navigate_to_website("https://poe.com/")
print("navigate_to_website")
time.sleep(3)


# Click the use email option
poe.click_use_email_button()
print("Using email option ")
time.sleep(2)

# Click the use email option send email
poe.click_on_email_box()
time.sleep(1)

# Click the email verification button
print("sending email")
poe.click_go_button()
time.sleep(25)

# Click to verify email OTP
print("verify email")
poe.click_login_button()
print("Login to account")
time.sleep(3)
print("Looping the Questions ")
poe.ask_questions()
print("Good work all the answer are copied and save to file..... END")

