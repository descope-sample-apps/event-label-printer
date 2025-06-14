import os
import time
from dotenv import load_dotenv
from datetime import datetime

try:
    import win32printing
    PRINTING_ENV = True
except: 
    print("No win32printing. Running test")
    PRINTING_ENV = False    

from descope import (
    AuthException,
    DescopeClient,
)

LINE_CLEAR = '\x1b[2K' 
MAX_NAME_LINE = 20
MAX_COMPANY_LINE = 40
MAX_TITLE_LINE = 40

print("Setting up...")
load_dotenv()
management_key = os.getenv("MANAGEMENT_KEY")
project_id = os.getenv("PROJECT_ID")

print("Project ID: " + project_id + ". Management Key: " + management_key[:5] + "********" + management_key[-5:])

try:
    descope_client = DescopeClient(
        project_id=project_id, management_key=management_key
    )
except Exception as error:
    print("failed to initialize. Error:")
    print(error)

def get_print_string(arr, key, max):
    if not key in arr:
        return ""
    print(arr[key])
    if (len(arr[key]) > max):
        return arr[key][:max]
    else:
        return arr[key]
    
def get_print_name_string(name, max):
    print(name)
    if (len(name) > max):
        return name[:max]
    else:
        return name
    
def get_name_array(full_name):
    name = full_name.strip()
    if (len(name) < 20):
        return [name]
    
    nameArr = name.split(" ")
    
    lines = []
    currentLine = get_print_name_string(nameArr[0], MAX_NAME_LINE)

    for idx in range (1,len(nameArr)):
        namePart =get_print_name_string(nameArr[idx],MAX_NAME_LINE)

        if len(currentLine + " " + namePart) > MAX_NAME_LINE:
            lines.append(currentLine)
            currentLine = namePart
        else:
            currentLine = currentLine + " " + namePart

    lines.append(currentLine)

    return lines[:3] 

def searchUsers():
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

def updateUser(user):
    login_id = user["loginIds"][0]
    print(login_id)
    attribute_key = "printed"
    attribute_val = True

    try:
        resp = descope_client.mgmt.user.update_custom_attribute(login_id=login_id, attribute_key=attribute_key, attribute_val=attribute_val)
        print(resp)
        print ("   Successfully updated user. LoginID: " + login_id + ". userEmail: " + user["email"] + ".")
        print()
    except AuthException as error:
        print ("Unable to update user's custom attribute.")
        print ("Status Code: " + str(error.status_code))
        print ("Error: " + str(error.error_message))


def printThis(user):
    print("   Printing " + user["email"])
    if (not PRINTING_ENV):
        return user

    with win32printing.Printer(
        printer_name="iDPRT SP410", margin=(0, 0, 5, 0)  # up, right, down, left
    ) as _printer:
        try:
            _printer.start_doc  # start job
            try:
                _printer.start_page  # using one label
                _namelist = get_name_array(user["name"])
                carryStr = None
                index = 0
                if len(max(_namelist, key=len)) >= 40:  # Too long...
                    _printer.end_page
                    _printer.end_doc
                    return None
                for name in _namelist:
                    if carryStr != None:
                        _namelist[index] = carryStr + " " + name
                    if len(name) > 20:
                        carryStr = name[19:]
                        _namelist[index] = name[:19] + "-"
                        if index == len(_namelist) - 1:
                            _namelist.append("")
                    else:
                        carryStr = None
                    index += 1
                _printfontAdjust = {
                    "height": 25 - (0.4 * len(max(_namelist, key=len))),
                    "weight": 800,
                    "charSet": "ANSI_CHARSET",
                    "faceName": "Consolas",
                }
                _printfontAdjustAlt = {
                    "height": 25 - 0.9 * len(user["name"]),
                    "weight": 800,
                    "charSet": "ANSI_CHARSET",
                    "faceName": "Consolas",
                }
                _printfontAdjustAltAlt = {
                    "height": 25 - 1.5 * len(user["name"]),
                    "weight": 800,
                    "charSet": "ANSI_CHARSET",
                    "faceName": "Consolas",
                }  # Adjusting the font so it keeps the name(s) on the page
                _printer.linegap = 5  # space between lines
                _printer.text(
                    " ", align="center", font_config=_printfontAdjust
                )  # Formatting
                match len(_namelist):
                    case 1:
                        _printer.text(
                            " ", align="center", font_config=_printfontAdjustAltAlt
                        )
                        _printer.text(
                            user["name"],
                            align="center",
                            font_config=_printfontAdjust,
                        )
                    case 2:
                        _printer.text(
                            " ", align="center", font_config=_printfontAdjustAlt
                        )
                        if len(user["name"]) < 20:
                            _printer.text(
                                _namelist[0] + " " + _namelist[1],
                                align="center",
                                font_config=_printfontAdjust,
                            )
                        else:
                            _printer.text(
                                _namelist[0],
                                align="center",
                                font_config=_printfontAdjust,
                            )
                            _printer.text(
                                _namelist[1],
                                align="center",
                                font_config=_printfontAdjust,
                            )
                    case 3:  # Same as case 2, but with 3 names instead of 2
                        if (len(_namelist[0]) + len(_namelist[1]) > 20) and (
                            len(_namelist[2]) + len(_namelist[1])
                        ) > 20:
                            _printer.text(
                                _namelist[0],
                                align="center",
                                font_config=_printfontAdjust,
                            )
                            _printer.text(
                                _namelist[1],
                                align="center",
                                font_config=_printfontAdjust,
                            )
                            _printer.text(
                                _namelist[2],
                                align="center",
                                font_config=_printfontAdjust,
                            )
                        elif (
                            len(_namelist[0]) + len(_namelist[1]) < 17
                        ):  # First + Middle on one line
                            _printer.text(
                                " ", align="center", font_config=_printfontAdjustAlt
                            )
                            _printer.text(
                                _namelist[0] + " " + _namelist[1],
                                align="center",
                                font_config=_printfontAdjust,
                            )
                            _printer.text(
                                _namelist[2],
                                align="center",
                                font_config=_printfontAdjust,
                            )
                        else:  # Middle + Last on one line
                            _printer.text(
                                " ", align="center", font_config=_printfontAdjustAlt
                            )
                            _printer.text(
                                _namelist[0],
                                align="center",
                                font_config=_printfontAdjust,
                            )
                            _printer.text(
                                _namelist[1] + " " + _namelist[2],
                                align="center",
                                font_config=_printfontAdjust,
                            )
                    case 4:
                        _printer.text(
                            _namelist[0],
                            align="center",
                            font_config=_printfontAdjust,
                        )
                        _printer.text(
                            _namelist[1],
                            align="center",
                            font_config=_printfontAdjust,
                        )
                        _printer.text(
                            _namelist[2],
                            align="center",
                            font_config=_printfontAdjust,
                        )
                        _printer.text(
                            _namelist[3],
                            align="center",
                            font_config=_printfontAdjust,
                        )
                    case _:  # Something went wrong
                        _printer.end_page
                        _printer.end_doc
                        return None
                _printer.linegap = (
                    150  # Bug with win32printing -- need to use large numbers like this
                )
                _printer.text("\u2500" * 10, align="center")  # Formatting line
                companyName = get_print_string(user["customAttributes"],"companyName",MAX_COMPANY_LINE)
                _printfontAdjust = {
                    "height": 22.5 - (0.25 * len(companyName)),
                    "weight": 600,
                    "charSet": "ANSI_CHARSET",
                    "faceName": "Consolas",
                }  # Same thing as before, but now for the company name
                _printer.text(
                    companyName,
                    align="center",
                    font_config=_printfontAdjust,
                )
                
                title = get_print_string(user["customAttributes"],"title",MAX_TITLE_LINE)
                _printfontAdjust = {
                    "height": 20 - (0.25 * len(title)),
                    "weight": 550,
                    "charSet": "ANSI_CHARSET",
                    "faceName": "Consolas",
                }  # Same thing as before, but now for the title
                _printer.text(
                    title,
                    align="center",
                    font_config=_printfontAdjust,
                )
                user["customAttributes"][
                    "print"
                ] = False  # Setting flag to false once done
                # note: this does NOT modify properly -- fix
            finally:
                _printer.end_page
        finally:
            _printer.end_doc
        return user


def printAlgo():
    while True:
        userList = searchUsers()  # All users checked in but not printed
        if userList != None:
            for user in userList:
                if user != None:
                    printThis(user)
                    updateUser(user)

        time.sleep(5)

def main():
    printAlgo()
    return 0

main()
