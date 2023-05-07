import time
from defFunction import PoeSite


# Create an instance of the class
poe = PoeSite()

# Navigate to the website
poe.navigate_to_website("https://poe.com/")
time.sleep(5)

# Click the use email option
poe.click_use_email_button()
time.sleep(3)

# Click the use email option send email
poe.click_on_email_box()
time.sleep(3)

# Click the email verification button
poe.click_go_button()
time.sleep(30)

# Click to verify email OTP
poe.click_login_button()
time.sleep(5)
poe.ask_questions()


