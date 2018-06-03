import wmi
import sys
import ctypes
import ctypes.wintypes
import math
import Tkinter

# Function to get only valid data
def getObjects (objList, values, notEmpty):
    ret = []
    for o in objList:
        cont = True
        for e in notEmpty:
            #print str(getattr(o, e)).strip()
            test = str(getattr(o, e)).strip()
            if test == "None" or len(test) == 0:
                cont = False
                break
        if cont == True:
            item = []
            for v in values:
                item.append(str(getattr(o, v)))
            ret.append(item)
    return ret

def getAcpiTable(tableID):
    try:
        FirmwareTableProviderSignature = ctypes.wintypes.DWORD(1094930505)
        FirmwareTableID = ctypes.wintypes.DWORD(tableID)
        pFirmwareTableBuffer = ctypes.create_string_buffer(0)
        BufferSize = ctypes.wintypes.DWORD(0)
        # Get size of data then allocate buffer
        size = ctypes.windll.kernel32.GetSystemFirmwareTable(FirmwareTableProviderSignature, FirmwareTableID, pFirmwareTableBuffer, BufferSize)
        pFirmwareTableBuffer=ctypes.create_string_buffer(size)
        BufferSize.value = size
        ctypes.windll.kernel32.GetSystemFirmwareTable(FirmwareTableProviderSignature, FirmwareTableID, pFirmwareTableBuffer, BufferSize)
        return pFirmwareTableBuffer.raw
    except:
        return ""

class WaitWindow:
    def __init__(self, windowTitle, msg):
        self.win = Tkinter.Tk()
        self.win.resizable(0, 0)
        self.win.iconbitmap('app.ico')
        self.win.title(windowTitle)
        self.win.protocol("WM_DELETE_WINDOW", sys.exit)
        self.msgLabel = Tkinter.Label(self.win, text=msg)
        self.msgLabel.pack()
        self.win.update()

    def close(self):
        self.win.destroy()

class MainWindow:
    def __init__(self, windowTitle):
        self.rootWin = Tkinter.Tk()
        self.rootWin.resizable(0, 0)
        self.rootWin.iconbitmap('app.ico')
        self.rootWin.title(windowTitle)
        self.rootWin.protocol("WM_DELETE_WINDOW", sys.exit)
        self.root = self.rootWin
        self.elements = {}
        self.gridX = 0
        self.gridY = 0

    def addLabel(self, text):
        self.elements["text_"+text] = Tkinter.Label(self.root, text=text, anchor=Tkinter.CENTER, justify=Tkinter.CENTER)
        self.elements["text_"+text].grid(row=self.gridY, column=self.gridX, sticky=Tkinter.W+Tkinter.E)
        self.gridY += 1

    def startFrame(self, type=0, thickness=1, colour="black"):
        if type == 0:
            self.frame = Tkinter.Frame(self.rootWin, highlightbackground=colour, highlightthickness=thickness)
        elif type == 1:
            self.frame = Tkinter.Frame(self.rootWin, borderwidth=thickness, relief=Tkinter.RAISED)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid(row=self.gridY, column=self.gridX, sticky=Tkinter.W+Tkinter.E)
        self.gridY += 1
        self.root = self.frame

    def endFrame(self):
        self.root = self.rootWin

    def addElement(self, label, data, label2 = "", data2 = ""):
        self.elements["label_"+label] = Tkinter.Label(self.root, text=label, anchor=Tkinter.W, justify=Tkinter.LEFT)
        self.elements["data_"+label] = Tkinter.Label(self.root, text=data, anchor=Tkinter.W, justify=Tkinter.LEFT)
        if label2 != "":
            self.elements["label2_"+label2] = Tkinter.Label(self.root, text=label2, anchor=Tkinter.W, justify=Tkinter.LEFT)
            self.elements["data2_"+label2] = Tkinter.Label(self.root, text=data2, anchor=Tkinter.W, justify=Tkinter.LEFT)
        
        self.elements["label_"+label].grid(row=self.gridY, column=self.gridX, sticky=Tkinter.W+Tkinter.E)
        self.elements["data_"+label].grid(row=self.gridY, column=self.gridX+1, sticky=Tkinter.W+Tkinter.E)
        if label2 != "":
            self.elements["label2_"+label2].grid(row=self.gridY, column=self.gridX+2, sticky=Tkinter.W+Tkinter.E)
            self.elements["data2_"+label2].grid(row=self.gridY, column=self.gridX+3, sticky=Tkinter.W+Tkinter.E)
        
        self.gridY += 1

    def addButton(self, label, action):
        self.elements["button_"+label] = Tkinter.Button(self.root, text=label, command=action)
        self.elements["button_"+label].grid(row=self.gridY, column=self.gridX)

    def focus(self):
        self.rootWin.focus_force()

    def mainLoop(self):
        self.root.mainloop()


# Please wait dialog
waitWindow = WaitWindow("System Detect Utility", "Detecting system specifications. Please wait...")

# Create objects for WMI
wmiObj = wmi.WMI()
wmiObj2 = wmi.WMI(moniker="//./root/wmi")

# Get hardware data
sysManuf = wmiObj.Win32_ComputerSystem()[0].Manufacturer
sysProd = wmiObj.Win32_ComputerSystem()[0].Model
cpuList = wmiObj.Win32_Processor()
ram = str(int(wmiObj.Win32_OperatingSystem()[0].TotalVisibleMemorySize)/1024)+"MB"
diskList = wmiObj.Win32_DiskDrive()
gpuList = wmiObj.Win32_VideoController()

# Try to get monitor sizes
try:
    displayList = wmiObj2.query("SELECT MaxHorizontalImageSize,MaxVerticalImageSize FROM WmiMonitorBasicDisplayParams")
except:
    displayList = False

# Get software data
osName = wmiObj.Win32_OperatingSystem()[0].Name.split("|")[0]
biosLic = getAcpiTable(1296323405)
if len(biosLic) != 0:
    biosLic = biosLic[56:len(biosLic)].decode("utf-8")
else:
    biosLic = "Not Found"

window = MainWindow("System Detect Utility")

window.addLabel("Detected system specifications")

window.startFrame(1, 3)
window.addLabel("=== HARDWARE INFORMATION ===")
window.endFrame()

window.startFrame()

window.addElement("SYSTEM MANUFACTURER:", sysManuf)

window.addElement("SYSTEM PRODUCT:", sysProd)

for i in getObjects(cpuList, ["Name"], ["Name"]):
    window.addElement("CPU:", i[0])

window.addElement("RAM:", ram)

for i in getObjects(diskList, ["Model", "Size"], ["Model", "Size"]):
    window.addElement("DISK:", i[0], "SIZE:", str(int(i[1])/1024/1024/1024)+"GB")

for i in getObjects(gpuList, ["Name", "AdapterRAM"], ["Name", "AdapterRAM"]):
    window.addElement("GPU:", i[0], "VIDEORAM:", str(int(i[1])/1024/1024)+"MB")

if displayList != False:
    for i in getObjects(displayList, ["MaxHorizontalImageSize", "MaxVerticalImageSize"], ["MaxHorizontalImageSize", "MaxVerticalImageSize"]):
        x = pow(float(i[0])/2.54,2)
        y = pow(float(i[1])/2.54,2)
        s = str(round(math.sqrt(x+y),0))
        window.addElement("DISPLAY:", s+" Inches")

window.endFrame()

window.addLabel("")
window.startFrame(1, 3)
window.addLabel("=== SOFTWARE INFORMATION ===")
window.endFrame()

window.startFrame()

window.addElement("OS:", osName)

window.addElement("BIOS LICENSE KEY:", biosLic)

window.endFrame()

waitWindow.close()
window.focus()
window.mainLoop()
