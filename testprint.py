import win32printing

from descope import (
    REFRESH_SESSION_TOKEN_NAME,
    SESSION_TOKEN_NAME,
    AuthException,
    DeliveryMethod,
    DescopeClient,
    AssociatedTenant,
    RoleMapping,
    AttributeMapping
)

management_key = "K2UiaOzJRt8VDlhs1dr8IQW2uT9eCodn3FYblAnCNFvyVZtbx2MAwpzKvjOjFe05o2kEycW"

try:
    # You can configure the baseURL by setting the env variable Ex: export DESCOPE_BASE_URI="https://auth.company.com  - this is useful when you utilize CNAME within your Descope project."
    descope_client = DescopeClient(project_id='C2J150l8sNop1jhp2AdOy9qmPBqZ', management_key=management_key)
except Exception as error:
    # handle the error
    print ("failed to initialize. Error:")
    print (error)

printfont = {"height": 22.5, "weight": 600, "charSet": "ANSI_CHARSET", "faceName": "Consolas"} 
printfontalt = {"height": 20, "weight": 610, "charSet": "ANSI_CHARSET", "faceName": "Consolas"} 
printfont2 = {"height": 20, "weight": 500, "charSet": "ANSI_CHARSET", "faceName": "Consolas"} 
printfont2alt = {"height": 18.75, "weight": 500, "charSet": "ANSI_CHARSET", "faceName": "Consolas"} 
printfont3 = {"height": 17.5, "weight": 400, "charSet": "ANSI_CHARSET", "faceName": "Consolas"} 

class Person:
    def __init__(self, name, org, role):
        self.name = name
        self.namelist = []
        self.org = org
        self.role = role
        self.print_status = False
        
    def printThis(self):
        with win32printing.Printer(printer_name = "iDPRT SP410", margin=(0,22.5,0,0)) as _printer:
            try:
                _printer.start_doc
                try:
                    _printer.start_page
                    if len(self.name) > 16:
                        self.namelist = self.name.split(" ")
                        _printer.linegap = 5
                        match len(self.namelist): #This is just in the case of middle names... might not need...
                            case 2:
                                _printer.text(self.namelist[0], align="center", font_config=printfont)
                                _printer.text(self.namelist[1], align="center",font_config=printfont)
                            case 3:
                                _printer.text(self.namelist[0], align="center", font_config=printfont)
                                _printer.text(self.namelist[1] + " " + self.namelist[2], align="center",font_config=printfont)
                            case 4:
                                _printer.text(self.namelist[0] + " " + self.namelist[1], align="center", font_config=printfont)
                                _printer.text(self.namelist[2] + " " + self.namelist[3], align="center",font_config=printfont)
                            case _:
                                _printer.end_page
                                _printer.end_doc
                                return None
                        _printer.linegap = 100
                    else:
                        _printer.linegap = 100
                        _printer.text(self.name, align="center", font_config=printfont)
                    _printer.text('\u2500' * 10, align="center")
                    if len(self.org) > 20:
                        _printer.text(self.org, align="center",font_config=printfont2alt)
                    else:
                        _printer.text(self.org, align="center",font_config=printfont2)
                    _printer.text(self.role, align="center",font_config=printfont3)
                    self.print_status = True
                finally:
                    _printer.end_page
            finally:
                _printer.end_doc
    
def main(): 
    #Some Sort of Input
    return 0
main()
