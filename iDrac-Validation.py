import xlrd

import subprocess

from tkinter import Tk
from tkinter import filedialog

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

class colors:
    OKGREEN = '\033[92m'
    FAIL = '\033[91m'
    OKPURPLE = '\033[96m'
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
                iDrac["svctag"] = sheet.cell_value(i, 7)
                iDrac["tmp_ip"] = sheet.cell_value(i, 8)
                iDrac["ip_address"] = sheet.cell_value(i, 9)
                iDrac["subnet"] = sheet.cell_value(i, 10)
                iDrac["gateway"] = sheet.cell_value(i, 11)
                iDrac["name"] = sheet.cell_value(i, 6)
                iDrac["vconsole"] = sheet.cell_value(i, 13)
                iDrac["timezone"] = sheet.cell_value(i, 14)
                iDrac["boot_mode"] = sheet.cell_value(i, 12)
                iDrac["vdisks"] = [sheet.cell_value(i, 16), sheet.cell_value(i, 18)]
                iDrac["pdisks"] = [sheet.cell_value(i, 17), sheet.cell_value(i, 19)]
                iDrac["IP_Check"] = sheet.cell_value(i, 20)
                iDrac["Raid1_Check"] = sheet.cell_value(i, 21)
                iDrac["Raid2_Check"] = sheet.cell_value(i, 22)
                #Append al idrac data
                iDrac_Data.append(iDrac)
        
        return iDrac_Data

class iDrac_Validation():
    def __init__(self, name, svctag, tmp_ip, ip, subnet, gateway, timezone, vconsole, boot_mode, raids, pdisks, ip_check, raid1_check, raid2_check):
        self.name = name
        self.svctag = svctag
        self.tmp_ip = tmp_ip
        self.ip = ip
        self.subnet = subnet
        self.gateway = gateway
        self.timezone = timezone
        self.vconsole = vconsole
        self.boot_mode = boot_mode
        self.raids = raids
        self.pdisks = pdisks
        self.ip_check = ip_check
        self.raid1_check = raid1_check
        self.raid2_check = raid2_check

        if ping(ip):
            self.pingable = ip
        elif ping(tmp_ip): self.pingable = tmp_ip

    def validateName(self):
        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn get iDRAC.Nic.DNSRacName".format(self.pingable)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output = str(sub.stdout.readlines()[1])
        if "=" in output:
            racadmName = output.replace("\\r", "").replace("\\n", "").strip("b'").split("=")[1]
            space = 48 - len(racadmName)

            if self.name == racadmName:
                print("Server name Status:             " + colors.OKGREEN + racadmName + colors.ENDC + " "*space + self.name)
            else:
                print("Server name Status:             " + colors.FAIL + racadmName + colors.ENDC + " "*space + self.name)

        else:
            print("Server name Status:     " + colors.FAIL + "Error getting data from racadm" + colors.ENDC + "                   " + self.name)

    def validateServiceTag(self):
        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn getsvctag".format(self.pingable)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output = str(sub.stdout.readlines()[0])
        racadmSvctag = output.replace("\\r", "").replace("\\n", "").strip("b'").replace(" ", "")
        space = 48 - len(racadmSvctag)

        if self.svctag == racadmSvctag:
            print("Server Service Tag Status:      " + colors.OKGREEN + racadmSvctag + colors.ENDC + " "*space + self.svctag)
        else:
            print("Server Service Tag Status:      " + colors.FAIL + racadmSvctag + colors.ENDC + " "*space + self.svctag)

    def validateIP(self):
        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn getniccfg".format(self.pingable)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output = sub.stdout.readlines()
        racadmIP = str(output[5]).replace("\\r", "").replace("\\n", "").strip("b'").replace(" ", "").split("=")[1]
        racadmSubnet = str(output[6]).replace("\\r", "").replace("\\n", "").strip("b'").replace(" ", "").split("=")[1]
        racadmGateway = str(output[7]).replace("\\r", "").replace("\\n", "").strip("b'").replace(" ", "").split("=")[1]

        space = 48 - len(racadmIP) - len(racadmSubnet) - len(racadmGateway) -4

        if self.ip == racadmIP and self.subnet == racadmSubnet and self.gateway == racadmGateway:
            print("Server IPv4 Status:             " + colors.OKGREEN + racadmIP + ", " + racadmSubnet + ", " + racadmGateway + colors.ENDC + " "*space + self.ip, self.subnet, self.gateway)
        else:
            print("Server IPv4 Status:             " + colors.FAIL + racadmIP + ", " + racadmSubnet + ", " + racadmGateway + colors.ENDC + " "*space + self.ip, self.subnet, self.gateway)

    def validateBootMode(self):
        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn get BIOS.BiosBootSettings.BootMode".format(self.pingable)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output = str(sub.stdout.readlines()[1])
        racadmBootMode = output.replace("\\r", "").replace("\\n", "").strip("b'").replace(" ", "").split("=")[1]

        if self.boot_mode in racadmBootMode:
            space = 48 - len(self.boot_mode)
            print("Server Boot Mode Status:        " + colors.OKGREEN + self.boot_mode + colors.ENDC + " "*space + self.boot_mode)
        else:
            space = 48 - len(racadmBootMode)
            print("Server Boot Mode Status:        " + colors.FAIL + racadmBootMode + colors.ENDC + " "*space + self.boot_mode)

    def validateVConsole(self):
        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn get idrac.VirtualConsole.plugintype".format(self.pingable)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output = str(sub.stdout.readlines()[1])
        racadmVConsole = output.replace("\\r", "").replace("\\n", "").strip("b'").replace(" ", "").split("=")[1]

        if self.vconsole == "java" and racadmVConsole == "1" or self.vconsole == "html5" and racadmVConsole == "2":
            if racadmVConsole == "2":
                space = 48 - len("html5")
                print("Server Virtual Console Status:  " + colors.OKGREEN + "html5" + colors.ENDC + " "*space + self.vconsole)
            else:
                space = 48 - len("java")
                print("Server Virtual Console Status:  " + colors.OKGREEN + "java" + colors.ENDC + " "*space + self.vconsole)
        else:
            if racadmVConsole == "2":
                space = 48 - len("html5")
                print("Server Virtual Console Status:  " + colors.FAIL + "html5" + colors.ENDC + " "*space + self.vconsole)
            else:
                space = 48 - len("java")
                print("Server Virtual Console Status:  " + colors.FAIL + "java" + colors.ENDC + " "*space + self.vconsole)

    def validateTimezone(self):
        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn get idrac.time.timezone".format(self.pingable)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output = str(sub.stdout.readlines()[1])
        racadmTimezone = output.replace("\\r", "").replace("\\n", "").strip("b'").replace(" ", "").split("=")[1]
        space = 48 - len(racadmTimezone)

        if self.timezone == racadmTimezone :
            print("Server Timezone Status:         " + colors.OKGREEN + racadmTimezone + colors.ENDC + " "*space + self.timezone)
        else:
            print("Server Timezone Status:         " + colors.FAIL + racadmTimezone + colors.ENDC + " "*space + self.timezone)

    def validateRaids(self):
        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn storage get vdisks -o -p Layout,Name".format(self.pingable)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output = sub.stdout.readlines()

        racadmVDisks = []
        racadmLayouts = []
        racadmNames = []

        if "Check" not in str(output): #if virtual disks exist
            for vdisk, layout, name in zip(range(0, len(output), +3), range(1, len(output), +3), range(2, len(output), +3)):
                racadmVDisks.append(str(output[vdisk]).replace("\\r", "").replace("\\n", "").strip("b'").replace(" ", ""))
                racadmLayouts.append(str(output[layout]).replace("\\r", "").replace("\\n", "").strip("b'").replace(" ", "").split("=")[1])
                racadmNames.append(str(output[name]).replace("\\r", "").replace("\\n", "").strip("b'").replace(" ", "").split("=")[1])

            maped = {"0": False, "1": False}
            for raid in self.raids:
                if str(raid) == "N/A":
                    continue
                elif str(int(raid)) == "0":
                    for vdisk, layout, name, (checkedKey, CheckedItem) in zip(racadmVDisks, racadmLayouts, racadmNames, maped.items()):
                        lay = str(layout).split("-")[1]
                        layoutSpace = 48 - len(layout)
                        nameSpace = 48 - len(name)

                        if "0" in lay and "10" not in lay and CheckedItem is False:
                            if "Virtual.0" in str(vdisk):
                                print("Server VDisks0 Layout:          " + colors.OKGREEN + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks0 Name:            " + colors.OKGREEN + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            elif "Virtual.1" in str(vdisk):
                                print("Server VDisks1 Layout:          " + colors.OKGREEN + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks1 Name:            " + colors.OKGREEN + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            maped[checkedKey] = True
                            break
                        elif layout == racadmLayouts[-1]:
                            if "Virtual.0" in str(vdisk):
                                print("Server VDisks0 Layout:         " + colors.FAIL + layout + colors.ENDC + "                                          " + "Raid-{}".format(int(raid)))
                                print("Server VDisks0 Name:           " + colors.FAIL + name + colors.ENDC + "                                           " + "Raid{}".format(int(raid)))

                            elif "Virtual.1" in str(vdisk):
                                print("Server VDisks1 Layout:         " + colors.FAIL + layout + colors.ENDC + "                                          " + "Raid-{}".format(int(raid)))
                                print("Server VDisks1 Name:           " + colors.FAIL + name + colors.ENDC + "                                           " + "Raid{}".format(int(raid)))

                elif str(int(raid)) == "1":
                    for vdisk, layout, name, (checkedKey, CheckedItem) in zip(racadmVDisks, racadmLayouts, racadmNames, maped.items()):
                        lay = str(layout).split("-")[1]
                        layoutSpace = 48 - len(layout)
                        nameSpace = 48 - len(name)

                        if "1" in lay and "10" not in lay and CheckedItem is False:
                            if "Virtual.0" in str(vdisk):
                                print("Server VDisks0 Layout:          " + colors.OKGREEN + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks0 Name:            " + colors.OKGREEN + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            elif "Virtual.1" in str(vdisk):
                                print("Server VDisks1 Layout:          " + colors.OKGREEN + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks1 Name:            " + colors.OKGREEN + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            maped[checkedKey] = True
                            break
                        elif layout == racadmLayouts[-1]:
                            if "Virtual.0" in str(vdisk):
                                print("Server VDisks0 Layout:         " + colors.FAIL + layout + colors.ENDC + "                                          " + "Raid-{}".format(int(raid)))
                                print("Server VDisks0 Name:           " + colors.FAIL + name + colors.ENDC + "                                           " + "Raid{}".format(int(raid)))

                            elif "Virtual.1" in str(vdisk):
                                print("Server VDisks1 Layout:         " + colors.FAIL + layout + colors.ENDC + "                                          " + "Raid-{}".format(int(raid)))
                                print("Server VDisks1 Name:           " + colors.FAIL + name + colors.ENDC + "                                           " + "Raid{}".format(int(raid)))

                elif str(int(raid)) == "5":
                    for vdisk, layout, name, (checkedKey, CheckedItem) in zip(racadmVDisks, racadmLayouts, racadmNames, maped.items()):
                        lay = str(layout).split("-")[1]
                        layoutSpace = 48 - len(layout)
                        nameSpace = 48 - len(name)

                        if "5" in lay and CheckedItem is False:
                            if "Virtual.0" in str(vdisk):
                                print("Server VDisks0 Layout:          " + colors.OKGREEN + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks0 Name:            " + colors.OKGREEN + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            elif "Virtual.1" in str(vdisk):
                                print("Server VDisks1 Layout:          " + colors.OKGREEN + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks1 Name:            " + colors.OKGREEN + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))
                                
                            maped[checkedKey] = True
                            break
                        elif layout == racadmLayouts[-1]:
                            if "Virtual.0" in str(vdisk):
                                print("Server VDisks0 Layout:         " + colors.FAIL + layout + colors.ENDC + "                                          " + "Raid-{}".format(int(raid)))
                                print("Server VDisks0 Name:           " + colors.FAIL + name + colors.ENDC + "                                           " + "Raid{}".format(int(raid)))

                            elif "Virtual.1" in str(vdisk):
                                print("Server VDisks1 Layout:         " + colors.FAIL + layout + colors.ENDC + "                                          " + "Raid-{}".format(int(raid)))
                                print("Server VDisks1 Name:           " + colors.FAIL + name + colors.ENDC + "                                           " + "Raid{}".format(int(raid)))

                elif str(int(raid)) == "6":
                    for vdisk, layout, name, (checkedKey, CheckedItem) in zip(racadmVDisks, racadmLayouts, racadmNames, maped.items()):
                        lay = str(layout).split("-")[1]
                        layoutSpace = 48 - len(layout)
                        nameSpace = 48 - len(name)

                        if "6" in lay and CheckedItem is False:
                            if "Virtual.0" in str(vdisk):
                                print("Server VDisks0 Layout:          " + colors.OKGREEN + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks0 Name:            " + colors.OKGREEN + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            elif "Virtual.1" in str(vdisk):
                                print("Server VDisks1 Layout:          " + colors.OKGREEN + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks1 Name:            " + colors.OKGREEN + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            maped[checkedKey] = True
                            break
                        elif layout == racadmLayouts[-1]:
                            if "Virtual.0" in str(vdisk):
                                print("Server VDisks0 Layout:          " + colors.FAIL + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks0 Name:            " + colors.FAIL + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            elif "Virtual.1" in str(vdisk):
                                print("Server VDisks1 Layout:          " + colors.FAIL + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks1 Name:            " + colors.FAIL + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                elif str(int(raid)) == "10":
                    for vdisk, layout, name, (checkedKey, CheckedItem) in zip(racadmVDisks, racadmLayouts, racadmNames, maped.items()):
                        lay = str(layout).split("-")[1]
                        layoutSpace = 48 - len(layout)
                        nameSpace = 48 - len(name)

                        if "10" in lay and CheckedItem is False:
                            if "Virtual.0" in str(vdisk):
                                print("Server VDisks0 Layout:          " + colors.OKGREEN + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks0 Name:            " + colors.OKGREEN + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            elif "Virtual.1" in str(vdisk):
                                print("Server VDisks1 Layout:          " + colors.OKGREEN + layout + colors.ENDC + " "*layoutSpace+ "Raid-{}".format(int(raid)))
                                print("Server VDisks1 Name:            " + colors.OKGREEN + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            maped[checkedKey] = True
                            break
                        elif layout == racadmLayouts[-1]:
                            if "Virtual.0" in str(vdisk):
                                print("Server VDisks0 Layout:          " + colors.FAIL + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks0 Name:            " + colors.FAIL + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

                            elif "Virtual.1" in str(vdisk):
                                print("Server VDisks1 Layout:          " + colors.FAIL + layout + colors.ENDC + " "*layoutSpace + "Raid-{}".format(int(raid)))
                                print("Server VDisks1 Name:            " + colors.FAIL + name + colors.ENDC + " "*nameSpace + "Raid{}".format(int(raid)))

        else:
            for raid in self.raids:
                if raid != "N/A":
                    print("Server VDisks Status:           " + colors.FAIL + "Not Exist" + colors.ENDC + " "*39 + "Raid-{}".format(int(raid)))

if __name__ == "__main__":
    #Get all servers detailes
    file_tk = Tk()
    file_tk.withdraw()
    excel_file = filedialog.askopenfilename(filetypes =[('Excel Files', '*.xlsx')])
    excelData = readExcel(excel_file).read()

    for server in excelData:
        print("====================Starting iDrac validation process for server {}====================".format(server["name"]))
        iDrac = iDrac_Validation(server["name"], server["svctag"], server["tmp_ip"], server["ip_address"], server["subnet"], server["gateway"], server["timezone"], server["vconsole"], server["boot_mode"], server["vdisks"], server["pdisks"], server["IP_Check"], server["Raid1_Check"], server["Raid2_Check"])
        print(colors.OKPURPLE +  "Category                        " + "Rac Value                                       " + "Excel Value "  + colors.ENDC)
        print("--------                        ---------                                       -----------")
        iDrac.validateName()
        iDrac.validateServiceTag()
        iDrac.validateIP()
        iDrac.validateBootMode()
        iDrac.validateVConsole()
        iDrac.validateTimezone()
        if all(raid == "N/A" for raid in server["vdisks"]): #Check if only "N/A" inside excel raids tabs
            print(colors.OKPURPLE + "No raids to check for." + colors.ENDC)
        else:
            iDrac.validateRaids()
        print("\n")