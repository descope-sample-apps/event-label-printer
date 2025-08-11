import os
import time

from dotenv import load_dotenv
from datetime import datetime
from descope import ( AuthException, DescopeClient )

LINE_CLEAR = '\x1b[2K' 
MAX_NAME_LINE = 20
MAX_COMPANY_LINE = 39
MAX_TITLE_LINE = 40
MAX_BALE_HEADER_LINE = 100
FONT_CHAR_SET = "ANSI_CHARSET"
FONT_FACE_NAME = "Consolas"

print("Setting up...")

try:
    import win32printing
    PRINTING_ENV = True
except: 
    print("No win32printing. Running without.")
    PRINTING_ENV = False

load_dotenv()
management_key = os.getenv("MANAGEMENT_KEY")
project_id = os.getenv("PROJECT_ID")

print("Project ID: " + project_id + ". Management Key: " + management_key[:5] + "********" + management_key[-5:])

try:
    descope_client = DescopeClient( project_id = project_id, management_key = management_key )
except Exception as error:
    print("failed to initialize. Error:")
    print(error)
    exit(1)

def get_print_string(arr, key, max):
    if not key in arr:
        return ""
    if (len(arr[key]) > max):
        return arr[key][:max]
    else:
        return arr[key]

def get_name_lines(full_name):
    name_lines = ["", ""]
    words = full_name.split(" ")

    name_lines[0] = words[0]

    if len(words) > 1:
        name_lines[1] = " ".join(words[1:])[:MAX_NAME_LINE]
    else:
        name_lines[1] = " "
    return name_lines


def search_users():
    custom_attributes = {"checkedIn": True, "approved": True, "printed": False}
    try:
        resp = descope_client.mgmt.user.search_all(custom_attributes=custom_attributes)
        users = resp["users"] 
        print(end=LINE_CLEAR) # <-- clear the line where cursor is located
        
        now = datetime.now()
        current_time = now.strftime("%H:%M:%S")

        print("   " + current_time + " Running... ", end='\r')

        if (len(users) > 0):
            print()
            print()
            print("   Found " + str(len(users)) + " users to print.")
        
        return users

    except AuthException as error:
        print("Unable to search users.")
        print("Status Code: " + str(error.status_code))
        print("Error: " + str(error.error_message))

def update_user(user):
    login_id = user["loginIds"][0]
    attribute_key = "printed"
    attribute_val = True

    try:
        resp = descope_client.mgmt.user.update_custom_attribute(login_id=login_id, attribute_key=attribute_key, attribute_val=attribute_val)
        print ("   Successfully updated user. Email: " + login_id)
        print()
    except AuthException as error:
        print ("Unable to update user's custom attribute.")
        print ("Status Code: " + str(error.status_code))
        print ("Error: " + str(error.error_message))
        exit(1)

def print_user(user):
    print("   Printing " + user["name"])

    font_gap = { "height": 12, "weight": 400, "charSet": FONT_CHAR_SET, "faceName": FONT_FACE_NAME }
    font_header = { "height": 12, "weight": 400, "charSet": FONT_CHAR_SET, "faceName": FONT_FACE_NAME }
    font_name = { "height": 32, "weight": 600, "charSet": FONT_CHAR_SET, "faceName": FONT_FACE_NAME }
    font_company = { "height": 20, "weight": 600, "charSet": FONT_CHAR_SET, "faceName": FONT_FACE_NAME }
    font_title = { "height": 16, "weight": 400, "charSet": FONT_CHAR_SET, "faceName": FONT_FACE_NAME }

    label_header = get_print_string(user["customAttributes"],"labelHeader",MAX_BALE_HEADER_LINE)
    name_lines = get_name_lines(user["name"])                
    company_name = get_print_string(user["customAttributes"],"companyName",MAX_COMPANY_LINE)
    title = get_print_string(user["customAttributes"],"title",MAX_TITLE_LINE)

    if (not PRINTING_ENV):
        print("   " + "-" * 35)
        print("   | " + label_header)
        print("   | " + name_lines[0])
        print("   | " + name_lines[1])
        print("   | " + company_name)
        print("   | " + title)
        print("   " + "-" * 35)
        return

    with win32printing.Printer( printer_name="iDPRT SP410", margin=(0, 0, 5, 0) ) as _printer:
        try:
            _printer.start_doc  # start job
            _printer.start_page  # using one label

            _printer.text(" ", align="center", font_config=font_gap)

            _printer.text(label_header, align="center", font_config=font_header)
            _printer.text(" ", align="center", font_config=font_gap)

            _printer.text(name_lines[0], align="center", font_config=font_name)
            _printer.text(" ", align="center", font_config=font_gap)
            _printer.text(name_lines[1], align="center", font_config=font_name)
            _printer.text(" ", align="center", font_config=font_gap)

            _printer.text("\u2500" * 35, align="center")
            _printer.text(" ", align="center", font_config=font_gap)

            _printer.text(company_name, align="center", font_config=font_company)

            _printer.text(" ", align="center", font_config=font_gap)

            _printer.text(title, align="center", font_config=font_title)

            _printer.end_page
        finally:
            _printer.end_doc
    return


def print_loop():
    while True:
        users_list = search_users()  # All users checked in but not printed
        if users_list != None:
            for user in users_list:
                if user != None:
                    print_user(user)
                    update_user(user)

        time.sleep(1)

def main():
    print_loop()
    return

main()
