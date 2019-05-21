import os
import getpass
import shutil
import subprocess
import winsound
import ctypes
import time

diskASCII = 0
user_dir = getpass.getuser()
students_list = ['Myung Park', "Tony Blayker"]
student_name = ''
week_no = 1


def detect_usb_storage():
    global diskASCII
    # Start in D drive because most computers have C drives for OS, can change
    label = ['D', 'E', 'F', 'G', 'H']

    for i in label:
        try:
            file = open(i + ':/')
        except Exception as e:
            '''
            error = 2  =>not found
            error = 13 =>permission denied (exist!)
            '''
            if (e.errno == 13):
                print("Disk : " + i + " Found!")
            else:
                if (i != 'A' and i != 'B'):
                    diskASCII = ord(i) - 1
                    break;


def detect_student_name():
    global student_name
    for name in students_list:
        exists = os.path.isdir('C:\\Users\\{}\\Desktop\\{}'.format(user_dir, name))
        if exists:
            student_name = name
            break
    if student_name.isspace():
        winsound.MessageBeep(-1)
        ctypes.windll.user32.MessageBoxW(0, "Folder not found!", "CoderKidsUSB", 0)


def eject_usb():
    global diskASCII
    tmp_file = open('tmp.ps1', 'w')
    tmp_file.write('$driveEject = New-Object -comObject Shell.Application\n')
    tmp_file.write('$driveEject.Namespace(17).ParseName("' + "{}".format(str(chr(diskASCII))) + ':").InvokeVerb("Eject")')
    tmp_file.close()
    process = subprocess.Popen(['powershell.exe', '-ExecutionPolicy', 'Unrestricted', './tmp.ps1'])
    process.communicate()
    time.sleep(2)
    os.remove("tmp.ps1")


detect_usb_storage()
detect_student_name()

#designed to ignore C drive
if str(chr(diskASCII)) != 'C':
    try:
        while os.path.exists('{}:\\Grace Covenant\\Coding Foundations\\Week {}'.format(str(chr(diskASCII)), str(week_no))):
            week_no += 1
        shutil.copytree('C:\\Users\\{}\\Desktop\\{}'.format(user_dir, student_name),
                        '{}:\\Grace Covenant\\Coding Foundations\\Week {}\\{}'.format(str(chr(diskASCII)), str(week_no), student_name),
                        symlinks=False, ignore=None)
        time.sleep(1)
    except Exception as e:
        '''
        error = 183 => File already exists!
        '''
        print(e)
        winsound.MessageBeep(-1)
        ctypes.windll.user32.MessageBoxW(0, "File already exists!", "CoderKidsUSB", 0)
        eject_usb()
        exit()

if os.path.exists('{}:\\Grace Covenant\\Coding Foundations\\Week {}\\{}'.format(str(chr(diskASCII,)), str(week_no), student_name)):
    eject_usb()
    winsound.MessageBeep(-1)
    ctypes.windll.user32.MessageBoxW(0, "Safe to Eject!", "CoderKidsUSB", 0)
else:
    winsound.MessageBeep(-1)
    ctypes.windll.user32.MessageBoxW(0, "No USB Found!", "CoderKidsUSB", 0)
