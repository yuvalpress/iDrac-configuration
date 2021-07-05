import xlrd

import subprocess

import asyncio

from tkinter import Tk
from tkinter import filedialog

from time import sleep

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

async def deleteRaids(iDrac):
    await iDrac.clearAllVD()
    await asyncio.sleep(1)

async def iDracObjectSummon(iDracsList):
    if len(iDracsList) > 0 :
        print("Clearing all raids from selected iDrac servers:")
    await asyncio.gather(*[deleteRaids(x) for x in iDracsList])
   
class colors:
    OKGREEN = '\033[92m'
    FAIL = '\033[91m'
    OKBLUE = '\033[95m'
    ENDC = '\033[0m'

class readExcel:
    def __init__(self, excelFile):
        self.excelFile = excelFile

    def read(self):
        wb = xlrd.open_workbook(self.excelFile)
        sheet = wb.sheet_by_name("Server_Configuration")
        iDrac_Data = []
        for i in range(6, sheet.nrows, +1):
            if "1" in str(sheet.cell_value(i, 23)):
                iDrac = {}
                iDrac["svctag"] = sheet.cell_value(i, 7)
                iDrac["ip_address"] = sheet.cell_value(i, 9)
                iDrac["name"] = sheet.cell_value(i, 6)

                #Append al idrac data
                iDrac_Data.append(iDrac)
        
        return iDrac_Data

class raidData:
    def __init__(self, ip):
        self.ip = ip

    def getControllers(self):
        try:
            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password storage get controllers".format(self.ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output = sub.stdout.readlines()
            for r in output:
                if "RAID" in str(r).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", ""):
                    controller = str(r).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")

            return controller

        except Exception as err:
            print(err)
            return None

class clearConfig():
    def __init__(self, name, serviceTag, ip):
        self.name = name
        self.serviceTag = serviceTag
        self.ip = ip

    def clearName(self):
        try:
            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password set iDRAC.Nic.DNSRacName iDrac-{}".format(self.ip, self.serviceTag)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output = sub.stdout.readlines()
            for line in output:
                if "successfully" in str(line) or "Success" in str(line):
                    print(colors.OKGREEN + "Successfully changed iDrac's name to iDrac-{} for server {}.".format(self.serviceTag, self.name) + colors.ENDC)

        except Exception as err:
            print(err)

    def clearTZ(self):
        try:
            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password set idrac.time.timezone UTC".format(self.ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output = sub.stdout.readlines()
            for line in output:
                if "successfully" in str(line) or "Success" in str(line):
                    print(colors.OKGREEN + "Successfully changed iDrac's Timezone to UTC for server iDrac-{} ({}).".format(self.serviceTag, self.name) + colors.ENDC)

        except Exception as err:
            print(err)

    def clearBootMode(self):
        try:
            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password set BIOS.BiosBootSettings.BootMode Bios".format(self.ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output = sub.stdout.readlines()
            for line in output:
                if "successfully" in str(line) or "Success" in str(line):
                    print(colors.OKGREEN + "Successfully changed iDrac's Boot Mode to Bios for server iDrac-{} ({}).".format(self.serviceTag, self.name) + colors.ENDC)

        except Exception as err:
            print(err)

    def clearVConsole(self):
        try:
            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password set idrac.VirtualConsole.plugintype HTML5".format(self.ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output = sub.stdout.readlines()
            for line in output:
                if "successfully" in str(line) or "Success" in str(line):
                    print(colors.OKGREEN + "Successfully changed iDrac's Virtual Console to HTML5 for server iDrac-{} ({}).".format(self.serviceTag, self.name) + colors.ENDC)

        except Exception as err:
            print(err)

    async def clearAllVD(self):
        try:
            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password --nocertwarn storage get vdisks".format(self.ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output = sub.stdout.readlines()

            if "No virtual disks are displayed" in str(output[0]):
                print(colors.OKBLUE + "NO virtual disks to delete for server {}".format(self.name) + colors.ENDC)

            else:
                #Get servers raid controller
                raids = raidData(self.ip)
                controller = raids.getControllers()
                
                for raid in output:
                    #delete vd command
                    toClear = str(raid).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")
                    print("Deleting raid from VD: {} for server {}".format(toClear, self.name))
                    sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password --nocertwarn raid deletevd:{}".format(self.ip, toClear)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    output = sub.stdout.readlines()

                    #Create JobQueue if successfull    
                    if len(sub.stderr.readlines()) == 0:
                        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password --nocertwarn jobqueue create {} --realtime".format(self.ip, controller)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        output = sub.stdout.readlines()
                        line = str(output[2]).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")

                        JID = line.split("=")[1]
                        print("Job has been initiated succesfully on server {}".format(self.name))
                        print("Process JID:",JID)

                        if len(sub.stderr.readlines()) == 0:
                            trueUntil = True
                            while trueUntil:
                                #Run until job is done
                                sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password --nocertwarn jobqueue view -i {}".format(self.ip, JID)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                                lines = sub.stdout.readlines()
                                line3 = str(lines[3]).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")
                                line7 = str(lines[7]).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")
                                
                                await asyncio.sleep(1)

                                print("{} Raid creation progress: ".format(self.name), line7)
                                if "100" in line7 and "Completed" in line3:
                                    print(colors.OKGREEN + "The deletion of {} succeeded for server: {}".format(toClear, self.name) + colors.ENDC)
                                    # print(colors.OKGREEN + "Job execution finished successfully for server: iDrac-{} ({})".format(self.serviceTag, self.name) + colors.ENDC)
                                    trueUntil = False
                                
                                elif "100" in line7 and "Failed" in line3:
                                    print(colors.FAIL + "The deletion of {} on server {} have failed".format(toClear, self.name) + colors.ENDC)
                                    # print(colors.FAIL + "Job DONE for server: iDrac-{} ({})".format(self.serviceTag, self.name) + colors.ENDC)
                                    trueUntil = False
        except Exception as e:
            print(e)

    def clearIP(self):
        try:
            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password Setniccfg -d".format(self.ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output = sub.stdout.readlines()
            for line in output:
                if "ENABLED" in str(line):
                    print(colors.OKGREEN + "Successfully changed iDrac's NIC to DHCP mode for server iDrac-{} ({}).".format(self.serviceTag, self.name) + colors.ENDC)

        except Exception as err:
            print(err)

if __name__ == "__main__":

    #Get all servers detailes
    file_tk = Tk()
    file_tk.withdraw()
    excel_file = filedialog.askopenfilename(filetypes =[('Excel Files', '*.xlsx')])
    excelData = readExcel(excel_file).read()

    print("Starting clearing process for the following iDrac servers:")
    for server in excelData:
        print(server["name"])

    print("\n")
    answer = input("Would you like to continue (y/n/A(Yes to All):")
    if answer == "A":
        iDracs = []
        for server in excelData:
            if ping(server["ip_address"]):
                iDRAC = clearConfig(server["name"], server["svctag"], server["ip_address"])
                iDracs.append(iDRAC)
            else:
                print("Couldnt connect to iDrac {}".format(server["name"]))

        for server in iDracs:
            server.clearName()
            server.clearTZ()
            server.clearVConsole()
            server.clearBootMode()
            print("\n")

        #Run all raids creation in async mode
        asyncio.run(iDracObjectSummon(iDracs))

        for server in iDracs:
            server.clearIP()
            
    elif answer == "y":
        print("Put in servers indexes (coma seperated) that you wish to clear (1,4,5..):")
        for i, iDrac in zip(range(0, len(excelData)), excelData):
            print(str(i+1) + ".", iDrac["name"])

        print("\n")
        serversToClear = input("Servers: ")
        serversToClear = serversToClear.split(",")

        #Get only approved servers
        iDracs = []
        for i, iDrac in zip(range(0, len(excelData)), excelData):
            if str(i+1) in serversToClear:
                iDRAC = clearConfig(iDrac["name"], iDrac["svctag"], iDrac["ip_address"])
                iDracs.append(iDRAC)

        #Eliminate unreachable ip addresses
        print(len(iDracs))
        toPop = []
        for i, server in zip(range(0, len(iDracs)), iDracs):
            if ping(server.ip) == False:
                print("Couldnt connect to iDrac {}".format(server.name))
                toPop.append(server)

        for p in toPop:
            for i, server in zip(range(0, len(iDracs)), iDracs):
                if p is server:
                    iDracs.pop(i)
                
        for iDrac in iDracs:
            iDrac.clearName()
            iDrac.clearTZ()
            iDrac.clearVConsole()
            iDrac.clearBootMode()

        #Run all raids creation in async mode
        asyncio.run(iDracObjectSummon(iDracs))

        for iDrac in iDracs:
            iDrac.clearIP()

    elif answer == "n":
        print(colors.OKGREEN + "Exiting script" + colors.ENDC)

    
