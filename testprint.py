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
        printer_name="iDPRT SP410", margin=(0, 0, 0, 0)
    ) as _printer:
        try:
            _printer.start_doc
            try:
                _printer.start_page
                _user.namelist = _user.name.split(" ")
                _maxstringlen = len(max(_user.namelist, key=len))
                printfontAdjust = {
                    "height": 25 - (0.4 * _maxstringlen),
                    "weight": 800,
                    "charSet": "ANSI_CHARSET",
                    "faceName": "Consolas",
                }
                _printer.linegap = 5
                _printer.text(" ", align="center", font_config=printfontAdjust)
                match len(
                    _user.namelist
                ):  # This is just in the case of middle names... might not need...
                    case 2:
                        if len(_user.name) > 17:
                            _printer.text(
                                _user.namelist[0],
                                align="center",
                                font_config=printfontAdjust,
                            )
                            _printer.text(
                                _user.namelist[1],
                                align="center",
                                font_config=printfontAdjust,
                            )
                        else:
                            _printer.text(
                                " ", align="center", font_config=printfontAdjust
                            )
                            _printer.text(
                                _user.namelist[0] + " " + _user.namelist[1],
                                align="center",
                                font_config=printfontAdjust,
                            )
                    case 3:
                        if (len(_user.namelist[0]) + len(_user.namelist[1]) > 17) and (
                            len(_user.namelist[2]) + len(_user.namelist[1])
                        ) > 17:
                            _printer.text(
                                _user.namelist[0],
                                align="center",
                                font_config=printfontAdjust,
                            )
                            _printer.text(
                                _user.namelist[1],
                                align="center",
                                font_config=printfontAdjust,
                            )
                            _printer.text(
                                _user.namelist[2],
                                align="center",
                                font_config=printfontAdjust,
                            )
                        elif len(_user.namelist[0]) + len(_user.namelist[1]) < 17:
                            _printer.text(
                                _user.namelist[0] + " " + _user.namelist[1],
                                align="center",
                                font_config=printfontAdjust,
                            )
                            _printer.text(
                                _user.namelist[2],
                                align="center",
                                font_config=printfontAdjust,
                            )
                        else:
                            _printer.text(
                                _user.namelist[0],
                                align="center",
                                font_config=printfontAdjust,
                            )
                            _printer.text(
                                _user.namelist[1] + " " + _user.namelist[2],
                                align="center",
                                font_config=printfontAdjust,
                            )
                    case _:
                        _printer.end_page
                        _printer.end_doc
                        return None
                _printer.linegap = 150
                _printer.text("\u2500" * 10, align="center")
                printfontAdjust = {
                    "height": 22.5 - (0.25 * len(_user.org)),
                    "weight": 600,
                    "charSet": "ANSI_CHARSET",
                    "faceName": "Consolas",
                }
                _printer.text(_user.org, align="center", font_config=printfontAdjust)
                printfontAdjust = {
                    "height": 20 - (0.25 * len(_user.role)),
                    "weight": 550,
                    "charSet": "ANSI_CHARSET",
                    "faceName": "Consolas",
                }
                _printer.text(_user.role, align="center", font_config=printfontAdjust)
                _user.print_status = True
            finally:
                _printer.end_page
        finally:
            _printer.end_doc


def printAlgo():
    userList = searchUsers()  # All users checked in but not printed
    for user in userList:
        printThis(user)
    return searchUsers()  # Users not printed


def main():
    # printAlgo()
    return 0


main()
