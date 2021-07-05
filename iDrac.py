import os
import subprocess
import xlrd
from time import sleep
import asyncio

from tkinter import filedialog
from tkinter import Tk

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

async def allRaids(iDrac):
    await iDrac.iDrac_Raids()
    await asyncio.sleep(1)

async def iDracObjectSummon(iDracsList):
    await asyncio.gather(*[allRaids(x) for x in iDracsList])
    
class colors:
    OKGREEN = '\033[92m'
    FAIL = '\033[91m'
    OKPURPLE = '\033[95m'
    ENDC = '\033[0m'

class iDracConf:
    def __init__(self, name, tmp_ip, ip, subnet, gateway, timezone, vconsole, boot_mode, raids, pdisks, ip_check, raid1_check, raid2_check):
        self.name = name
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
    
    def setName(self):
        if self.name != "N/A":
            try:
                sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password set iDRAC.Nic.DNSRacName {}".format(self.ip, self.name)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output = sub.stdout.readlines()
                for line in output:
                    if "successfully" in str(line) or "Success" in str(line):
                        print(colors.OKGREEN + "Successfully changed iDrac's name to {} at ip {}.".format(self.name, self.ip) + colors.ENDC)

            except Exception as err:
                print(err)

    def setTZ(self):
        if self.timezone != "N/A":
            try:
                sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password set idrac.time.timezone {}".format(self.ip, self.timezone)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output = sub.stdout.readlines()
                for line in output:
                    if "successfully" in str(line) or "Success" in str(line):
                        print(colors.OKGREEN + "Successfully changed iDrac's Timezone to {} for server {}.".format(self.timezone, self.name) + colors.ENDC)

            except Exception as err:
                print(err)

    def setBootMode(self):
        if self.boot_mode != "N/A":
            try:
                sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password set BIOS.BiosBootSettings.BootMode {}".format(self.ip, self.boot_mode)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output = sub.stdout.readlines()
                for line in output:
                    if "successfully" in str(line) or "Success" in str(line):
                        print(colors.OKGREEN + "Successfully changed iDrac's Boot Mode to {} for server {}.".format(self.boot_mode, self.name) + colors.ENDC)

            except Exception as err:
                print(err)

    def setVconsole(self):
        if self.vconsole != "N/A":
            try:
                sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password set idrac.VirtualConsole.plugintype {}".format(self.ip, self.vconsole)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output = sub.stdout.readlines()
                for line in output:
                    if "successfully" in str(line) or "Success" in str(line):
                        print(colors.OKGREEN + "Successfully changed iDrac's Virtual Console to {} for server {}.".format(self.vconsole, self.name) + colors.ENDC)

            except Exception as err:
                print(err)

    async def iDrac_Raids(self):
        raid = raidData(self.ip, self.raids, self.pdisks, self.raid1_check, self.raid2_check)
        controller = raid.getControllers()
        raidLevel = raid.getRaidLevel()
        raidNames = raid.getRaidNames()
        disks = raid.getDisksNames()

        # First Raid Configuration Configuration
        if self.raid1_check == 1:
            try:
                for coma in range(0, len(   disks[0])): #Add coma between disks
                    if coma < len(disks[0]) - 1:
                        disks[0][coma] = disks[0][coma] + ","

                diskString = "".join(disks[0])
                bayString = raid.getDisksLocation(diskString)

                print(colors.OKPURPLE + "Configuring VD_0 ({}) for server {}\nThe raid will be created on disks: {}".format(raidNames[0], self.name, bayString) + colors.ENDC)
                sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password raid createvd:{} -rl {} -wp wb -rp ra -pdkey:{} -name {}".format(self.ip, controller, raidLevel[0], diskString, raidNames[0])],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)

                #Create JobQueue if successfull    
                if len(sub.stderr.readlines()) == 0:
                    sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password --nocertwarn jobqueue create {} --realtime".format(self.ip, controller)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    output = str(sub.stdout.readlines()[2]).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")
                    JID = output.split("=")[1]
                    print("Job has been initiated succesfully on server {} with JID {}".format(self.name, JID))

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
                                print(colors.OKGREEN + "The creation of {} succeeded for server: {} ({})".format(raidNames[0], self.name, self.ip) + colors.ENDC)
                                print(colors.OKGREEN + "Job execution finished successfully for server: {} ({})".format(self.name, self.ip) + colors.ENDC)
                                trueUntil = False
                            
                            elif "100" in line7 and "Failed" in line3:
                                print(colors.FAIL + "The creation of {} on server {} ({}) have failed".format(raidNames[0], self.name, self.ip) + colors.ENDC)
                                print(colors.FAIL + "Job DONE for server: {} ({})".format(self.name, self.ip) + colors.ENDC)
                                trueUntil = False

            except Exception as err:
                print(colors.FAIL + "The creation of {} on server {} ({}) have failed with the following error: {}".format(raidNames[0], self.name, self.ip, err) + colors.ENDC)

        # Second Raid Configuration
        if self.raid2_check == 1:
            #if the first raid is approved for configuration
            if self.raid1_check == 1:
                sleep(90)

                for coma in range(0, len(disks[1])): #Add coma between disks
                    if coma < len(disks[1]) - 1:
                        disks[1][coma] = disks[1][coma] + ","

                diskString = "".join(disks[1])
                bayString = raid.getDisksLocation(diskString)

            #if first raid is not approved for configuration
            else:
                for coma in range(0, len(disks[0])): #Add coma between disks
                    if coma < len(disks[0]) - 1:
                        disks[0][coma] = disks[0][coma] + ","

                diskString = "".join(disks[0])
                bayString = raid.getDisksLocation(diskString)

            try:
                print(colors.OKPURPLE + "Configuring VD_1 ({}) for server {}\nThe raid will be created on disks: {}".format(raidNames[0], self.name, bayString) + colors.ENDC)
                sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password raid createvd:{} -rl {} -wp wb -rp ra -pdkey:{} -name {}".format(self.ip, controller, raidLevel[1], diskString, raidNames[1])],
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE)

                #Create JobQueue if successfull    
                if len(sub.stderr.readlines()) == 0:
                    sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password --nocertwarn jobqueue create {} --realtime".format(self.ip, controller)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    output = str(sub.stdout.readlines()[2]).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")

                    JID = output.split("=")[1]
                    print("Job has been initiated succesfully on server {} with JID {}".format(self.name, JID))

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
                                print(colors.OKGREEN + "Succeeded configure {} for server: {}".format(raidNames[1], self.name) + colors.ENDC)
                                print(colors.OKGREEN + "Job execution finished successfully for server: {}".format(self.name) + colors.ENDC)
                                trueUntil = False
                            
                            elif "100" in line7 and "Failed" in line3:
                                print(colors.FAIL + "Failed to configure {} for server: {}".format(raidNames[1], self.name) + colors.ENDC)
                                print(colors.FAIL + "Job DONE for server: {}".format(self.name) + colors.ENDC)
                                trueUntil = False
                
            except Exception as err:
                print(colors.FAIL + "The creation of {} on server {} have failed with the following error: {}".format(raidNames[1], self.name, err) + colors.ENDC)

    def iDrac_IP(self):
        if self.ip_check == 1:
            try:
                sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password Setniccfg -s {} {} {}".format(self.tmp_ip, self.ip, self.subnet, self.gateway)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output = sub.stdout.readlines()
                for line in output:
                    if "successfully" in str(line) or "Success" in str(line):
                        print(colors.OKGREEN + "Successfully changed iDrac's IP Address to {}, {}, {} for server {}.".format(self.ip, self.subnet, self.gateway, self.name) + colors.ENDC)

            except Exception as err:
                print(err)

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

class raidData:
    def __init__(self, ip, raid, disks, raid1_check, raid2_check):
        self.ip = ip
        self.raid = raid
        self.disks = disks
        self.raid1_check = raid1_check
        self.raid2_check = raid2_check

    def getRaidNames(self):
        raidName = []
        if type(self.raid[0]) == float:
            if "0" in str(int(self.raid[0])) and "1" not in str(int(self.raid[0])):
                raidName.append("Raid0")
            elif "1" in str(int(self.raid[0])) and "0" not in str(int(self.raid[0])):
                    raidName.append("Raid1")
            elif "5" in str(int(self.raid[0])):
                raidName.append("Raid5")
            elif "6" in str(int(self.raid[0])):
                raidName.append("Raid6")
            elif "10" in str(int(self.raid[0])):
                raidName.append("Raid10")

        if type(self.raid[1]) == float:
            if "0" in str(int(self.raid[1])) and "1" not in str(int(self.raid[1])):
                raidName.append("Raid0")
            elif "1" in str(int(self.raid[1])) and "0" not in str(int(self.raid[1])):
                raidName.append("Raid1")
            elif "5" in str(int(self.raid[1])):
                raidName.append("Raid5")
            elif "6" in str(int(self.raid[1])):
                raidName.append("Raid6")
            elif "10" in str(int(self.raid[1])):
                raidName.append("Raid10")           

        return raidName

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

    def getDisksNames(self):
        raidDisks = []
        try:
            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password --nocertwarn storage get pdisks -o -p size".format(self.ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output = sub.stdout.readlines()
            pdisks = {}
            for disk, size in zip(range(0, len(output), +2), range(1, len(output), +2)):
                pdisks[str(output[disk]).replace(" ", "").replace("\\r", "").strip("b'").strip("\\n")] = str(output[size]).replace(" ", "").replace("\\r\\n", "").strip("b'")
                if disk + 2 > len(output) or size + 2 > len(output):
                    break
            
            sortedSizes = []
            for disk, size in pdisks.items():
                sortedSizes.append(float(size.split("=")[1].strip("GB")))
                sortedSizes = sorted(sortedSizes)

            #If all disks are the same
            if float(sortedSizes[0]) == float(sortedSizes[len(sortedSizes) - 1]):
                if self.raid1_check == 1:
                    #If number of disks is represented as float\int
                    if type(self.disks[0]) is not str:
                        d = []
                        #for all disks needed for the first raid
                        for dNum in range(0, int(self.disks[0])):
                            d.append(list(pdisks.keys())[dNum])
                        
                        #Delete first disks[0] disks that will be used for the creation of the first raid
                        for dNum in range(0, int(self.disks[0])):
                            pdisks.pop(list(pdisks.keys())[0])
                        raidDisks.append(d)

                    else:
                        #If a string other then "N/A" is written in the specific cell of excel sheet
                        if "N/A" not in self.disks[0]:
                            d = []
                            #for all the disks on server
                            for dNum in range(0, len(pdisks)):
                                d.append(list(pdisks.keys())[dNum])
                            raidDisks.append(d)

                #if raid2 is needed for configuration
                if self.raid2_check == 1:
                    if self.raid1_check == 1:
                        if type(self.disks[1]) is not str:
                            d = []
                            #for all disks needed for the second raid
                            for dNum in range(0, int(self.disks[1])):
                                d.append(list(pdisks.keys())[dNum])
                            
                            #Delete first disks[0] disks that will be used for the creation of the second raid
                            for dNum in range(0, int(self.disks[1])):
                                pdisks.pop(list(pdisks.keys())[0])
                            raidDisks.append(d)
                            
                        else:
                            #if all disks are needed for the second raid
                            if "N/A" not in self.disks[1]:
                                d = []
                                for dNum in range(0, len(pdisks)):
                                    d.append(list(pdisks.keys())[dNum])
                                raidDisks.append(d)
                    else:
                        #Delete first disks[0] disks that were used for the creation of the first raid
                        for dNum in range(0, int(self.disks[0])):
                            pdisks.pop(list(pdisks.keys())[0])

                        if type(self.disks[1]) is not str:
                            d = []
                            #for all disks needed for the second raid
                            for dNum in range(0, int(self.disks[1])):
                                d.append(list(pdisks.keys())[dNum])
                            
                            #Delete first disks[0] disks that will be used for the creation of the second raid
                            for dNum in range(0, int(self.disks[1])):
                                pdisks.pop(list(pdisks.keys())[0])
                            raidDisks.append(d)
                            
                        else:
                            if "N/A" not in self.disks[1]:
                                d = []
                                for dNum in range(0, len(pdisks)):
                                    d.append(list(pdisks.keys())[dNum])
                                raidDisks.append(d)

            #if different size disks exist on the server
            else:
                #if first raid configuration is needed
                if self.raid1_check == 1:
                    #if a number is representing the disks
                    if type(self.disks[0]) is not str:
                        #sum the amount of smaller sized disks on the specified server
                        cnt = 0
                        for disk, sizeCount in pdisks.items():
                            if float(sortedSizes[0]) in sizeCount:
                                cnt += 1

                        d = []
                        #if the smaller sized disks amount is the amount needed for the first raid
                        if cnt == int(self.disks[0]):
                            #add disks to the disk names list
                            for disk, sizeCount in pdisks.items():
                                if float(sortedSizes[0]) in sizeCount:
                                    d.append(disk)

                            #delete disks used for the first raid from the dictionary
                            for disk, sizeCount in pdisks.items():
                                if float(sortedSizes[0]) in sizeCount:
                                    pdisks.pop(disk)
                            raidDisks.append(d)

                        #if the bigger sized disks amount is the amount needed for the first raid
                        elif int(self.disks[0]) == (len(sortedSizes)) - cnt:
                            #add disks to the disk names list
                            for disk, sizeCount in pdisks.items():
                                if float(sortedSizes[-1]) in sizeCount:
                                    d.append(disk)
                            
                            #delete disks used for the first raid from the dictionary
                            for disk, sizeCount in pdisks.items():
                                if float(sortedSizes[-1]) in sizeCount:
                                    pdisks.pop(disk)
                            raidDisks.append(d)

                if self.raid2_check == 1:
                    #if the first raid configuration process was done at this script run
                    if self.raid1_check == 1:
                        #add remaind disks to the disk names list
                        for disk, sizeCount in pdisks.items():
                            if float(sortedSizes[0]) in sizeCount:
                                d.append(disk)

                        raidDisks.append(d)
                    
                    else:
                        #get all disks names and states
                        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p password --nocertwarn storage get pdisks -o -p state".format(self.ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        output = sub.stdout.readlines()

                        #collect all disks in "Ready" state
                        remainedDisks = []
                        for disk, state in zip(range(0, len(output), +2), range(1, len(output), +2)):
                            if "Ready" in output[state]:
                                remainedDisks.append(str(output[disk]).replace(" ", "").replace("\\r\\n", "").strip("b'"))

                            if disk + 2 > len(output) or size + 2 > len(output):
                                break

                        #add remaind disks to the disk names list
                        for disk in remainedDisks:
                                d.append(disk)

                        raidDisks.append(d)

            return raidDisks

        except Exception as err:
            print(err)

    def getRaidLevel(self):
        raidLevel = []
        if type(self.raid[0]) == float:
            if "0" in str(int(self.raid[0])) and "1" not in str(int(self.raid[0])):
                raidLevel.append("r0")
            elif "1" in str(int(self.raid[0])) and "0" not in str(int(self.raid[0])):
                raidLevel.append("r1")
            elif "5" in str(int(self.raid[0])):
                raidLevel.append("r5")
            elif "6" in str(int(self.raid[0])):
                raidLevel.append("r6")
            elif "10" in str(int(self.raid[0])):
                raidLevel.append("r10")

        if type(self.raid[1]) == float:
            if "0" in str(int(self.raid[1])) and "1" not in str(int(self.raid[1])):
                raidLevel.append("r0")
            elif "1" in str(int(self.raid[1])) and "0" not in str(int(self.raid[1])):
                    raidLevel.append("r1")
            elif "5" in str(int(self.raid[1])):
                raidLevel.append("r5")
            elif "6" in str(int(self.raid[1])):
                raidLevel.append("r6")
            elif "10" in str(int(self.raid[1])):
                raidLevel.append("r10")
        
        return raidLevel

    def getDisksLocation(self, comaDisks):
        disksString = ""
        for disk in comaDisks.split(","):
            disksString += str(disk).split(":")[0] + ", "
        disksString = disksString[:-2]

        return disksString

if __name__ == "__main__":
    file_tk = Tk()
    file_tk.withdraw()
    excel_file = filedialog.askopenfilename(filetypes =[('Excel Files', '*.xlsx')])
    excelData = readExcel(excel_file).read()
    #Perform actions

    print("====================Starting iDracs IP Address changes for all servers====================")
    #IP Address changing section
    for server in excelData:
        try:
            #Create iDrac object
            iDrac = iDracConf(server["name"], server["tmp_ip"], server["ip_address"], server["subnet"], server["gateway"], 
                server["timezone"],server["vconsole"], server["boot_mode"], server["vdisks"], server["pdisks"], server["IP_Check"], server["Raid1_Check"], server["Raid2_Check"])
            
            #Change ip address for iDrac server
            iDrac.iDrac_IP()


        except Exception as err:
            print("===================================================================================================")
            print("The iDrac IP address changing process for all servers failed with the following error: ", err)
            continue

    #iDrac base configuration section
    for server in excelData:
        print("====================Starting iDrac configuration process for server {} ({})====================".format(server["name"], server["ip_address"]))
        try:
            #Create iDrac object
            iDrac = iDracConf(server["name"], server["tmp_ip"], server["ip_address"], server["subnet"], server["gateway"], 
                server["timezone"],server["vconsole"], server["boot_mode"], server["vdisks"], server["pdisks"], server["IP_Check"], server["Raid1_Check"], server["Raid2_Check"])

            #Wait for ip address to be pingable
            while ping(server["ip_address"]) == False:
                print("Waiting for ip address {} of server {} to be reachable".format(server["ip_address"], server["name"]))
            
            iDrac.setBootMode()
            iDrac.setVconsole()
            iDrac.setTZ()
            iDrac.setName()
            
            print("===================================================================================================")

        except Exception as err:
            print("===================================================================================================")
            print("iDrac configuration process for server {} ({}) failed with the following error: ".format(server["name"], server["ip_address"]), err)
            continue

    #Raid sections
    iDracs = []
    for server in excelData:
        iDrac = iDracConf(server["name"], server["tmp_ip"], server["ip_address"], server["subnet"], server["gateway"], 
            server["timezone"],server["vconsole"], server["boot_mode"], server["vdisks"], server["pdisks"], server["IP_Check"], server["Raid1_Check"], server["Raid2_Check"])
        iDracs.append(iDrac)

    #Run all raids creation in async mode
    asyncio.run(iDracObjectSummon(iDracs))
