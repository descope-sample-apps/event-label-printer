import win32printing
import json

from descope import (
    REFRESH_SESSION_TOKEN_NAME,
    SESSION_TOKEN_NAME,
    AuthException,
    DeliveryMethod,
    DescopeClient,
    AssociatedTenant,
    RoleMapping,
    AttributeMapping,
)

management_key = (
    "K2UiaOzJRt8VDlhs1dr8IQW2uT9eCodn3FYblAnCNFvyVZtbx2MAwpzKvjOjFe05o2kEycW"
)

try:
    # You can configure the baseURL by setting the env variable Ex: export DESCOPE_BASE_URI="https://auth.company.com  - this is useful when you utilize CNAME within your Descope project."
    descope_client = DescopeClient(
        project_id="C2J150l8sNop1jhp2AdOy9qmPBqZ", management_key=management_key
    )
except Exception as error:
    # handle the error
    print("failed to initialize. Error:")
    print(error)

printfont = {
    "height": 22.5,
    "weight": 600,
    "charSet": "ANSI_CHARSET",
    "faceName": "Consolas",
}
printfont2 = {
    "height": 20,
    "weight": 500,
    "charSet": "ANSI_CHARSET",
    "faceName": "Consolas",
}
printfont2alt = {
    "height": 18.75,
    "weight": 500,
    "charSet": "ANSI_CHARSET",
    "faceName": "Consolas",
}
printfont3 = {
    "height": 17.5,
    "weight": 400,
    "charSet": "ANSI_CHARSET",
    "faceName": "Consolas",
}


def searchUsers():
    custom_attributes = {"print": False, "checkedIn": True}
    try:
        resp = descope_client.mgmt.user.search_all(custom_attributes=custom_attributes)
        print("Successfully searched users.")
        users = resp["users"]
        for user in users:
            print(print(json.dumps(user, indent=2)))
        return users
    except AuthException as error:
        print("Unable to search users.")
        print("Status Code: " + str(error.status_code))
        print("Error: " + str(error.error_message))


def printThis(_user):
    with win32printing.Printer(
        printer_name="iDPRT SP410", margin=(0, 22.5, 0, 0)
    ) as _printer:
        try:
            _printer.start_doc
            try:
                _printer.start_page
                if len(_user.name) > 16:
                    if len(_user.name) > 70:
                        return None
                    _user.namelist = _user.name.split(" ")
                    _printer.linegap = 5
                    match len(_user.namelist):
                        case 2:
                            _printer.text(
                                _user.namelist[0],
                                align="center",
                                font_config=printfont,
                            )
                            _printer.text(
                                _user.namelist[1],
                                align="center",
                                font_config=printfont,
                            )
                        case 3:
                            _printer.text(
                                _user.namelist[0],
                                align="center",
                                font_config=printfont,
                            )
                            _printer.text(
                                _user.namelist[1] + " " + _user.namelist[2],
                                align="center",
                                font_config=printfont,
                            )
                        case _:
                            _printer.end_page
                            _printer.end_doc
                            return None
                    _printer.linegap = 100
                else:
                    _printer.linegap = 100
                    _printer.text(_user.name, align="center", font_config=printfont)
                _printer.text("\u2500" * 10, align="center")
                if len(_user.org) > 20:
                    _printer.text(_user.org, align="center", font_config=printfont2alt)
                else:
                    _printer.text(_user.org, align="center", font_config=printfont2)
                _printer.text(_user.role, align="center", font_config=printfont3)
                _user.print_status = True
            finally:
                _printer.end_page
        finally:
            _printer.end_doc

def printAlgo():
    userList = searchUsers() #All users checked in but not printed
    for user in userList:
        printThis(user)
    return searchUsers() #Users not printed

def main():
    #printAlgo()
    
    return 0


main()
