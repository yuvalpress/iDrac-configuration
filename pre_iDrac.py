import xlrd
import subprocess
from tkinter import Tk
from tkinter import filedialog

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
    def __init__(self, ip):
        self.ip = ip

    def getControllers(self):
        try:
            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! storage get controllers".format(self.ip)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output = sub.stdout.readlines()
            if "RAID" in str(output[2]).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", ""):
                controller = str(output[2]).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")

            return controller

        except Exception as err:
            print(err)
            return None

class iDrac_Pre():
    def __init__(self, servers):
        self.servers = servers

    def powerUp(self):
        for server in self.servers:
            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn serveraction powerup".format(server["tmp_ip"])], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            print(sub.stdout.readlines[0].replace(" ", "").replace("\\r\\n", "").strip("b'"))


    def raidReady(self):
        for server in self.servers:
            if server["Raid1_check"] == 1 or server["Raid2_Check"] == 1:
                controller = raidData(server["tmp_ip"]).getControllers()
                #get all disks names and states
                sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn storage get pdisks -o -p state".format(server["tmp_ip"])], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                output = sub.stdout.readlines()

                #collect all disks in "Ready" state
                all_non_ready = []
                for disk, state in zip(range(0, len(output), +2), range(1, len(output), +2)):
                    if "Online" not in output[state] and "Ready" not in output[state]:
                        all_non_ready.append(str(output[disk]).replace(" ", "").replace("\\r\\n", "").strip("b'"))
                        sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn storage converttoraid:{}".format(server["tmp_ip"], str(output[disk]).replace(" ", "").replace("\\r\\n", "").strip("b'"))], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                
                if len(all_non_ready) != 0:
                    sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn jobqueue create {} --realtime".format(server["tmp_ip"], controller)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    output = sub.stdout.readlines()
                    line = str(output[2]).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")

                    JID = line.split("=")[1]
                    print("Job has been initiated succesfully on server {}".format(server["name"]))
                    print("Process JID:",JID)

                    if len(sub.stderr.readlines()) == 0:
                        trueUntil = True
                        while trueUntil:
                            #Run until job is done
                            sub = subprocess.Popen(["powershell", "& racadm -r {} -u root -p Customer1! --nocertwarn jobqueue view -i {}".format(server["ip]"], JID)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                            lines = sub.stdout.readlines()
                            line3 = str(lines[3]).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")
                            line7 = str(lines[7]).strip("b'").replace("\\r\\n", "").replace(" ", "").replace("\\r", "")
                            
                            await asyncio.sleep(1)

                            print("{} Raid creation progress: ".format(self.name), line7)
                            if "100" in line7 and "Completed" in line3:
                                print(colors.OKGREEN + "The deletion of {} succeeded for server: iDrac-{} ({})".format(toClear, self.serviceTag, self.name) + colors.ENDC)
                                # print(colors.OKGREEN + "Job execution finished successfully for server: iDrac-{} ({})".format(self.serviceTag, self.name) + colors.ENDC)
                                trueUntil = False
                            
                            elif "100" in line7 and "Failed" in line3:
                                print(colors.FAIL + "The deletion of {} on server iDrac-{} ({}) have failed".format(toClear, self.serviceTag, self.name) + colors.ENDC)
                                # print(colors.FAIL + "Job DONE for server: iDrac-{} ({})".format(self.serviceTag, self.name) + colors.ENDC)
                                trueUntil = False


if __name__ == "__main__":

    #Get all servers detailes
    file_tk = Tk()
    file_tk.withdraw()
    excel_file = filedialog.askopenfilename(filetypes =[('Excel Files', '*.xlsx')])
    excelData = readExcel(excel_file).read()

    #Power up all servers
    iDrac_Pre