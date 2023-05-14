import os, winshell, win32com.client, pythoncom
import shutil
import sys
import threading
import time
import win32com
import win32con
import win32gui

progressCOLOR = '\033[38;5;33;48;5;236m' #BLUEgreyBG
finalCOLOR = '\033[48;5;33m' #BLUEBG
# check the color codes below and paste above

###### COLORS #######
# WHITEblueBG = '\033[38;5;15;48;5;33m'
# BLUE = '\033[38;5;33m'
# BLUEBG  = '\033[48;5;33m'
# ORANGEBG = '\033[48;5;208m'
# BLUEgreyBG = '\033[38;5;33;48;5;236m'
# ORANGEgreyBG = '\033[38;5;208;48;5;236m' # = '\033[38;5;FOREGROUND;48;5;BACKGROUNDm' # ver 'https://i.stack.imgur.com/KTSQa.png' para 256 color codes
# INVERT = '\033[7m'
###### COLORS #######

BOLD    = '\033[1m'
UNDERLINE = '\033[4m'
CEND    = '\033[0m'

FilesLeft = 0

def FullFolderSize(path):
    TotalSize = 0
    if os.path.exists(path):# to be safely used # if FALSE returns 0
        for root, dirs, files in os.walk(path):
            for file in files:
                TotalSize += os.path.getsize(os.path.join(root, file))
    return TotalSize

def getPERCECENTprogress(source_path, destination_path, bytes_to_copy):
    dstINIsize = FullFolderSize(destination_path)
    time.sleep(.25)
    print( " ")
    print ("FROM:" + "   " + source_path)
    print ("TO:" + "     " + destination_path)
    print (" ")
    if os.path.exists(destination_path):
        while bytes_to_copy != (FullFolderSize(destination_path)-dstINIsize):
            sys.stdout.write('\r')
            percentagem = int((float((FullFolderSize(destination_path)-dstINIsize))/float(bytes_to_copy)) * 100)
            steps = int(percentagem/5)
            copiado = '{:,}'.format(int((FullFolderSize(destination_path)-dstINIsize)/1000000))# Should be 1024000 but this get's closer to the file manager report
            sizzz = '{:,}'.format(int(bytes_to_copy/1000000))
            # sys.stdout.write(("         {:s} / {:s} Mb  ".format(copiado, sizzz)) +  (BOLD + progressCOLOR + "{:20s}".format('|'*steps) + CEND) + ("  {:d}% ".format(percentagem)) + ("  {:d} ToGo ".format(FilesLeft))) #  STYLE 1 progress default #
            # sys.stdout.write(("         {:s} / {:s} Mb  ".format(copiado, sizzz)) +  (BOLD + progressCOLOR + "{:20s}".format('|'*steps) + CEND) + ("  {:d}% ".format(percentagem)) + ("  {:d} ToGo ".format(FilesLeft))) #  STYLE 1 progress default #
            #BOLD# sys.stdout.write(BOLD + ("        {:s} / {:s} Mb  ".format(copiado, sizzz)) +  (progressCOLOR + "{:20s}".format('|'*steps) + CEND) + BOLD + ("  {:d}% ".format(percentagem)) + ("  {:d} ToGo ".format(FilesLeft))+ CEND) # STYLE 2 progress BOLD #
            sys.stdout.write(BOLD + ("        {:s} / {:s} Mb  ".format(copiado, sizzz)) +  ("|{:20s}|".format('|'*steps)) + ("  {:d}% ".format(percentagem)) + ("  {:d} ToGo ".format(FilesLeft))+ CEND) # STYLE 3 progress classic B/W #
            sys.stdout.flush()
            time.sleep(.01)
        sys.stdout.write('\r')
        time.sleep(.05)
        # sys.stdout.write(("         {:s} / {:s} Mb  ".format('{:,}'.format(int((FullFolderSize(destination_path)-dstINIsize)/1000000)), '{:,}'.format(int(bytes_to_copy/1000000)))) +  (BOLD + finalCOLOR + "{:20s}".format(' '*20) + CEND) + ("  {:d}% ".format( 100)) + ("  {:s}      ".format('    ')) + "\n") #  STYLE 1 progress default #
        #BOLD# sys.stdout.write(BOLD + ("        {:s} / {:s} Mb  ".format('{:,}'.format(int((FullFolderSize(destination_path)-dstINIsize)/1000000)), '{:,}'.format(int(bytes_to_copy/1000000)))) +  (finalCOLOR + "{:20s}".format(' '*20) + CEND) + BOLD + ("  {:d}% ".format( 100)) + ("  {:s}      ".format('    ')) + "\n" + CEND ) # STYLE 2 progress BOLD #
        sys.stdout.write(BOLD + ("        {:s} / {:s} Mb  ".format('{:,}'.format(int((FullFolderSize(destination_path)-dstINIsize)/1000000)), '{:,}'.format(int(bytes_to_copy/1000000)))) +  ("|{:20s}|".format('|'*20)) + ("  {:d}% ".format( 100)) + ("  {:s}      ".format('    ')) + "\n" + CEND ) # STYLE 3 progress classic B/W #
        sys.stdout.flush()
        print (" ")
        print (" ")

def CopyProgress(SOURCE, DESTINATION):
    global FilesLeft
    DST = os.path.join(DESTINATION, os.path.basename(SOURCE))
    # <- the previous will copy the Source folder inside of the Destination folder. Result Target: path/to/Destination/SOURCE_NAME
    # -> UNCOMMENT the next (# DST = DESTINATION) to copy the CONTENT of Source to the Destination. Result Target: path/to/Destination
    DST = DESTINATION # UNCOMMENT this to specify the Destination as the target itself and not the root folder of the target
    #
    if DST.startswith(SOURCE):
        print(" ")
        print('Source folder can\'t be changed.')
        print('Please check your target path...')
        print(" ")
        print('        CANCELED')
        print(" ")
        exit()
    #count bytes to copy
    Bytes2copy = 0
    for root, dirs, files in os.walk(SOURCE): # USE for filename in os.listdir(SOURCE): # if you don't want RECURSION #
        dstDIR = root.replace(SOURCE, DST, 1) # USE dstDIR = DST # if you don't want RECURSION #
        for filename in files:                # USE if not os.path.isdir(os.path.join(SOURCE, filename)): # if you don't want RECURSION #
            dstFILE = os.path.join(dstDIR, filename)
            if os.path.exists(dstFILE): continue # must match the main loop (after "threading.Thread")
            #                                      To overwrite delete dstFILE first here so the progress works properly: ex. change continue to os.unlink(dstFILE)
            #                                      To rename new files adding date and time, instead of deleating and overwriting,
            #                                      comment 'if os.path.exists(dstFILE): continue'
            Bytes2copy += os.path.getsize(os.path.join(root, filename)) # USE os.path.getsize(os.path.join(SOURCE, filename)) # if you don't want RECURSION #
            FilesLeft += 1
    # <- count bytes to copy
    #
    # Treading to call the preogress
    threading.Thread(name='progresso', target=getPERCECENTprogress, args=(SOURCE, DST, Bytes2copy)).start()
    # main loop
    for root, dirs, files in os.walk(SOURCE): # USE for filename in os.listdir(SOURCE): # if you don't want RECURSION #
        dstDIR = root.replace(SOURCE, DST, 1) # USE dstDIR = DST # if you don't want RECURSION #
        if not os.path.exists(dstDIR):
            os.makedirs(dstDIR)
        for filename in files:                # USE if not os.path.isdir(os.path.join(SOURCE, filename)): # if you don't want RECURSION #
            srcFILE = os.path.join(root, filename) # USE os.path.join(SOURCE, filename) # if you don't want RECURSION #
            dstFILE = os.path.join(dstDIR, filename)
            if os.path.exists(dstFILE): continue # MUST MATCH THE PREVIOUS count bytes loop
            #   <- <-                              this jumps to the next file without copying this file, if destination file exists.
            #                                      Comment to copy with rename or overwrite dstFILE
            #
            # RENAME part below
            head, tail = os.path.splitext(filename)
            count = -1
            year = int(time.strftime("%Y"))
            month = int(time.strftime("%m"))
            day = int(time.strftime("%d"))
            hour = int(time.strftime("%H"))
            minute = int(time.strftime("%M"))
            while os.path.exists(dstFILE):
                count += 1
                if count == 0:
                    dstFILE = os.path.join(dstDIR, '{:s}[{:d}.{:d}.{:d}]{:d}-{:d}{:s}'.format(head, year, month, day, hour, minute, tail))
                else:
                    dstFILE = os.path.join(dstDIR, '{:s}[{:d}.{:d}.{:d}]{:d}-{:d}[{:d}]{:s}'.format(head, year, month, day, hour, minute, count, tail))
            # END of RENAME part
            shutil.copy2(srcFILE, dstFILE)
            FilesLeft -= 1
            #

chrome_handle = win32gui.FindWindow(None, "Quality Control ToolBox")
if chrome_handle != 0:
    win32gui.PostMessage(chrome_handle, win32con.WM_CLOSE, 0, 0)

desktop = winshell.desktop()

if os.path.isdir('C:/QCCenter'):
    if os.path.isfile(os.path.join(desktop, 'QCToolBox Shortcut.lnk')):
        os.remove(os.path.join(desktop, 'QCToolBox Shortcut.lnk'))
    if os.path.isdir('C:/QCCenter_Old'):
        shutil.rmtree('C:/QCCenter_Old')
        os.rename('C:/QCCenter', 'C:/QCCenter_Old')
        os.makedirs('C:/QCCenter')
    else:
        os.rename('C:/QCCenter', 'C:/QCCenter_Old')
        os.makedirs('C:/QCCenter')
else:
    os.makedirs('C:/QCCenter')

os.makedirs('C:\\QCCenter\\QCHub')
CopyProgress('N:\\Images\\Shahaf\\Projects\\QCHub', 'C:\\QCCenter\\QCHub')

os.makedirs('C:\\QCCenter\\ProjectPython')
CopyProgress("N:\\Images\\Shahaf\\Projects\\ProjectPython", "C:\\QCCenter\\ProjectPython")

os.makedirs('C:\\QCCenter\\QCTools')
CopyProgress("N:\\Images\\Shahaf\\Projects\\QCTools", "C:\\QCCenter\\QCTools")

# os.makedirs('C:\\QCCenter\\TemplateExcelGenerator')
# CopyProgress("N:\\Images\\Shahaf\\Projects\\TemplateExcelGenerator", "C:\\QCCenter\\TemplateExcelGenerator")

os.makedirs('C:\\QCCenter\\BatchFiles')

script = ["QCPrep",
          "Duplicator",
          "DoubleDisplay",
          "ExcelCells2Files",
          "FileList",
          "IDnJsonResults2Excel",
          "JPEGConverter",
          "JsonAge2Excel",
          "JsonFaceLiveness2Excel",
          "Jsons2Excel",
          "MoveFiles",
          "NameScrambler",
          "NameSwitcher",
          "PDF2JPG",
          "POAnJsonResults2Excel",
          "SendMail",
          "BatchCleanser",
          "SetCreator",
          "JsonExtractor",
          "DBRemover"]

for projectName in script:
    if projectName == "DBRemover":
        myBat = open('C:\\QCCenter\\BatchFiles\\{project}.bat'.format(project=projectName), 'w+')
        myBat.write('''@echo off
        powershell -Command "& 'C:\\QCCenter\\ProjectPython\\Scripts\\python.exe' 'C:\\QCCenter\\QCTools\\{project}.py'"
        pause
        '''.format(project=projectName))
        continue

    myBat = open('C:\\QCCenter\\BatchFiles\\{project}.bat'.format(project=projectName), 'w+')
    myBat.write('''@echo off
    powershell -WindowStyle Hidden -Command "& 'C:\\QCCenter\\ProjectPython\\Scripts\\python.exe' 'C:\\QCCenter\\QCTools\\{project}.py'"
    exit
    '''.format(project=projectName))

    myBat.close()

myBat = open('C:\\QCCenter\\BatchFiles\\QCToolBox.bat', 'w+')
myBat.write('''@echo off
powershell -WindowStyle Hidden -Command "& 'C:\\QCCenter\\ProjectPython\\Scripts\\python.exe' 'C:\\QCCenter\\QCHub\\main.py'"
exit
''')
myBat.close()

path = os.path.join(desktop, 'QCToolBox Shortcut.lnk')
target = "C:\\QCCenter\\BatchFiles\\QCToolBox.bat"
icon = r"N:\Images\Shahaf\Projects\Assests\tabicon.ico"
shell = win32com.client.Dispatch("WScript.Shell")
shortcut = shell.CreateShortCut(path)
shortcut.Targetpath = target
shortcut.IconLocation = icon
shortcut.save()

os.startfile(r"C:\QCCenter\BatchFiles\QCToolBox.bat")
