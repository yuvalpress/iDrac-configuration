import xlrd

from tkinter import Tk
from tkinter import filedialog

import subprocess

def ping(ip): #Check ping response

    #ping requested server
    try:
        ping = subprocess.Popen(
            ["ping", "-n", "1", '{}'.format(ip)],
            stdout = subprocess.PIPE,
            stderr = subprocess.PIPE)
        out, error = ping.communicate()
    except Exception as err:
        print(err)

    if "TTL" in str(out):
        return True
    else: return False

def powerUpServer(tmp_ip, ip, name):
    if ping(tmp_ip):
        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password --nocertwarn serveraction powerup".format(tmp_ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        isSuccess = sub.stdout.readlines()
        if "successfully" in str(isSuccess).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", ""):
            print(colors.OKGREEN + "{} PowerUp was Successfull".format(name) + colors.ENDC)
            return True
        elif "already" in str(isSuccess).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", ""):
            print(colors.OKPURPLE + "{} already in PowerUp state".format(name) + colors.ENDC)
        else: 
            print(colors.FAIL + "{} PowerUp was not Successfull".format(name) + colors.ENDC)
            return False

    elif ping(ip):
        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password --nocertwarn serveraction powerup".format(ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        isSuccess = sub.stdout.readlines()
        if "successfully" in str(isSuccess).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", ""):
            print(colors.OKGREEN + "{} PowerUp was Successfull".format(name) + colors.ENDC)
            return True
        elif "already" in str(isSuccess).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", ""):
            print(colors.OKPURPLE + "{} already in PowerUp state".format(name) + colors.ENDC)
        else: 
            print(colors.FAIL + "{} PowerUp was not Successfull".format(name) + colors.ENDC)
            return False

    else: print(colors.OKPURPLE + "Both temporary ip address and ip address aren\'t pingable for server {}".format(name) + colors.ENDC)

class colors:
    OKGREEN = '\033[92m'
    FAIL = '\033[91m'
    OKPURPLE = '\033[95m'
    ENDC = '\033[0m'

class readExcel:
    def __init__(self, excelFile):
        self.excelFile = excelFile

    def read(self):
        wb = xlrd.open_workbook(self.excelFile)
        sheet = wb.sheet_by_name("Server_Configuration")
        iDrac_Data = []
        for i in range(6, sheet.nrows, +1):
            if "1" in str(sheet.cell_value(i, 1)):
                iDrac = {}
                iDrac["tmp_ip"] = sheet.cell_value(i, 8)
                iDrac["ip_address"] = sheet.cell_value(i, 9)
                iDrac["name"] = sheet.cell_value(i, 6)

                #Append al idrac data
                iDrac_Data.append(iDrac)
        
        return iDrac_Data

if __name__ == "__main__":
    #Choose excel file 
    file_root = Tk()
    file_root.withdraw()
    excel_file = filedialog.askopenfilename(filetypes =[('Excel Files', '*.xlsx')])
    excelData = readExcel(excel_file).read()

    answer = input("Would you like to continue (y/n/A(Yes to All):")

    if answer == "A":
        for server in excelData:
            powerUpServer(server["tmp_ip"], server["ip_address"], server["name"])

    elif answer == "y":
        print("Put in servers indexes (coma seperated) that you wish to PowerUp (1,4,5..):")
        for i, server in zip(range(0, len(excelData)), excelData):
            print(str(i+1) + ". " + server["name"])

        print("\n")
        serversToClear = input("Servers: ")
        serversToClear = serversToClear.split(",")

        for i, server in zip(range(0, len(excelData)), excelData):
            if str(i+1) in serversToClear:
                powerUpServer(tmp_ip=server["tmp_ip"], ip=server["ip_address"], name=server["name"])

    elif answer == "n":
        print(colors.OKPURPLE + "Exiting script.." + colors.ENDC)

