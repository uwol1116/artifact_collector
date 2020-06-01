import pymysql
import win32evtlog # requires pywin32 pre-installed
import os
import sys
import ctypes
import winreg
import datetime
import string
import binascii
import struct
from datetime import datetime,timedelta
import subprocess
import codecs
import csv
import collections


#레지스트리 관련
def Informationn():
	key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", 0)
	####print ("OS VERSION : " + winreg.QueryValueEx(key,"ProductName")[0])
	####print ("Detailed VERSION : " + winreg.QueryValueEx(key,"BuildLabEx")[0])
	####print ("USER : " + winreg.QueryValueEx(key,"RegisteredOwner")[0])
	a=winreg.QueryValueEx(key,"ProductName")[0]
	b=winreg.QueryValueEx(key,"BuildLabEx")[0]
	c=winreg.QueryValueEx(key,"RegisteredOwner")[0]
	winreg.CloseKey(key)
	
	key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\\ControlSet001\\Control\\Windows", 0)
	####print ("Last_ShutDown_Time : ",end="")
	date = winreg.QueryValueEx(key,"ShutdownTime")[0].hex()
	strdate = date[-2:] + date[-4:-2] + date[-6:-4] + date[-8:-6] + date[-10:-8] + date[-12:-10] + date[-14:-12] + date[-16:-14]
	dateint = int(strdate,16)
	us = dateint / 10.
	####print(datetime(1601,1,1)+ timedelta(microseconds=us))
	i=1
	d=datetime(1601,1,1)+ timedelta(microseconds=us)
	winreg.CloseKey(key)
	sql="""insert into informationn(idx,OS_VERSION,Detailed_VERSION,USER,Last_ShutDown_Time) values (%s,%s,%s,%s,%s)"""
	curs.execute(sql,(i,a,b,c,d,))
	conn.commit()


def UserAssist():
	varReg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
	varKey = winreg.OpenKey(varReg, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\UserAssist\\{CEBFF5CD-ACE2-4F4F-9178-9926F41749EA}\\Count",0, winreg.KEY_ALL_ACCESS)
	#key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\UserAssist\\{CEBFF5CD-ACE2-4F4F-9178-9926F41749EA}\\Count',0, winreg.KEY_ALL_ACCESS)
	
	intab = "ABCDEFGHIJKLMabcdefghijklmNOPQRSTUVWXYZnopqrstuvwxyz"
	outtab = "NOPQRSTUVWXYZnopqrstuvwxyzABCDEFGHIJKLMabcdefghijklm"
	rot13 = str.maketrans(intab,outtab)

	try:
		i = 0
		while True:
			SessionNumber = ""
			RunCount = ""
			LastRunTime = ""
			name, value, type = winreg.EnumValue(varKey,i)
			
			####print (name.translate(rot13))
			text1 = value
			

			for j in range(0,8):
				SessionNumber += binascii.hexlify(text1).decode()[j]
			SessionUnpack = SessionNumber[-2:] + SessionNumber[-4:-2] + SessionNumber[-6:-4]+SessionNumber[-8:-6]
					
			for j in range(8,16):
				RunCount += binascii.hexlify(text1).decode()[j]
			RunUnpack = RunCount[-2:] + RunCount[-4:-2] + RunCount[-6:-4]+RunCount[-8:-6]

			for j in range(120,136):
				LastRunTime += binascii.hexlify(text1).decode()[j]
			LastRunUnpack = LastRunTime[-2:] + LastRunTime[-4:-2] + LastRunTime[-6:-4] + LastRunTime[-8:-6] + LastRunTime[-10:-8] + LastRunTime[-12:-10] + LastRunTime[-14:-12] + LastRunTime[-16:-14]
			lastint = int(LastRunUnpack,16)
			us = lastint / 10.
			#print(datetime(1601,1,1)+ timedelta(microseconds=us))
			
			####print("Session Number : " + str(SessionUnpack) + " Run Count : " + str(RunUnpack) + " Last Run Time : " + str(datetime(1601,1,1)+ timedelta(microseconds=us)))
			####print()


			a=name.translate(rot13)
			b=str(SessionUnpack)
			c=str(RunUnpack)
			d=(datetime(1601,1,1)+ timedelta(microseconds=us))

			#query="INSERT into registry1 VALUES (?,?,?,?,?)"
			sql="""insert into registry1(idx,Name,Session_Number,Run_Count,Last_Run_Time) values (%s,%s,%s,%s,%s)"""
			curs.execute(sql,(i,a,b,c,d,))
			conn.commit()
			i += 1

	except WindowsError:
		pass








def AutoRun(): #시작프로그램
    caption = subprocess.getoutput('wmic startup get Caption').split("\n")
   
    command = subprocess.getoutput('wmic startup get command').split("\n")
    count=1
   
    for i in range(1,len(caption)-2):
        if i % 2 == 0:
            a = caption[i]
            b = command[i]

            sql="""insert into registry2(idx,caption,command) values (%s,%s,%s)"""
            curs.execute(sql,(count-1,a,b,))
            conn.commit()
            count+=1
         #print("caption : " + caption[i] + "command : " + command[i])







def USB(): #usb 연결 관련 정보
	varkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM",0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
	i = 0
	USB_Result = ""
	USB_User = ""

	while True:
		try:
			key = winreg.EnumKey(varkey,i)
			USB_Result += key
			i += 1

		except WindowsError:
			break

	USB_Result.split(',')
	fullname = len(USB_Result)
	index = USB_Result.find("ControlSet")
	USB_User = USB_Result[index+len("ControlSet"):index+len("ControlSet")+3]
	varkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\\ControlSet"+USB_User+"\\Enum\\USB",0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
	####print("USB, USBTOR INFO")

	try:
		i = 0
		while True:
			device = winreg.EnumKey(varkey,i)
			####print(device)
			a=str(device)
			sql="""insert into registry3(idx,USB_Name) values (%s,%s)"""
			curs.execute(sql,(i+1,a,))
			conn.commit()
			i += 1

	except WindowsError:
		pass

	varkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\\ControlSet"+USB_User+"\\Enum\\USBSTOR",0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)	
	try:
		i = 0
		while True:
			device = winreg.EnumKey(varkey,i)
			####print(device)
			i += 1

	except WindowsError:
		pass
	
	####print("="*50)
	####print("DeviceClasses INFO")
	####print("WpdBusEnumRoo GUID")
	varkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet"+USB_User+"\\Control\\DeviceClasses\\{10497B1B-BA51-44E5-8318-A65C837B6661}",0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
	
	try:
		i = 0
		while True:
			device = winreg.EnumKey(varkey,i)
			####print(device)
			a=device
			sql="""insert into registry4_Device_INFO(WpdBusEnumRoo_GUID) values (%s)"""
			curs.execute(sql,(a,))
			conn.commit()
			i += 1
	except WindowsError:
		pass
	####print()
	####print("Disk GUID")
	varkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet"+USB_User+"\\Control\\DeviceClasses\\{53F5630D-B6BF-11D0-94F2-00A0C91EFB8B}",0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
	try : 
		i = 0
		while True:
			device = winreg.EnumKey(varkey,i)
			####print(device)
			a=device
			sql="""insert into registry4_Device_INFO(Disk_GUID) values (%s)"""
			curs.execute(sql,(a,))
			conn.commit()
			i += 1
	except WindowsError:
		pass

	####print()
	####print("Volume GUID")
	varkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet"+USB_User+"\\Control\\DeviceClasses\\{53F5630D-B6BF-11D0-94F2-00A0C91EFB8B}",0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
	try:
		i = 0
		while True:
			device = winreg.EnumKey(varkey,i)
			####print(device)
			a=device
			sql="""insert into registry4_Device_INFO(Volume_GUID) values (%s)"""
			curs.execute(sql,(a,))
			conn.commit()
			i += 1
	except WindowsError:
		pass
	
	####print()
	####print("Portable Device GUID")
	varkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet"+USB_User+"\\Control\\DeviceClasses\\{6AC27878-A6FA-4155-BA85-F98F491D4F33}",0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
	try :
		i = 0
		while True:
			device = winreg.EnumKey(varkey,i)
			####print(device)
			a=device
			sql="""insert into registry4_Device_INFO(Portable_Device_GUID) values (%s)"""
			curs.execute(sql,(a,))
			conn.commit()
			i += 1
	except WindowsError:
		pass

	####print()
	####print("USB GUID")
	varkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SYSTEM\\ControlSet"+USB_User+"\\Control\\DeviceClasses\\{A5DCBF10-6530-11D2-901F-00C04FB951ED}",0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
	try : 
		i = 0
		while True:
			device = winreg.EnumKey(varkey,i)
			####print("device")
			a=device
			sql="""insert into registry4_Device_INFO(USB_GUID) values (%s)"""
			curs.execute(sql,(a,))
			conn.commit()
			i += 1

	except:
		pass
	
	####print()
	####print("List of connected devices")
	varkey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,r"SOFTWARE\\Microsoft\\Windows Portable Devices\\devices",0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
	device_name = ""
	try:
		i = 0
		while True:
			device = winreg.EnumKey(varkey,i)
			device_name += device
			device_name += ":"
			i += 1
	except WindowsError:
		pass

	result = device_name.split(":")
	for i in range(0,len(result)-1):
		location_result = result[i]
		with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Windows Portable Devices\\devices\\"+location_result, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
			####print(winreg.QueryValueEx(key,'FriendlyName')[0])
			a=winreg.QueryValueEx(key,'FriendlyName')[0]
			####print(result[i])
			b=result[i]
			sql="""insert into registry5(idx,Connected_devices_name,specific_name) values (%s,%s,%s)"""
			curs.execute(sql,(i+1,a,b,))
			conn.commit()






def Network(): #무선네트워크 연결정보
				#table = network column = name, password 필요
	Nlist = subprocess.getoutput("netsh wlan show profile").split("모든 사용자 프로필 :")
	b=""

	count = 1
	while count < len(Nlist):
		name = Nlist[count].strip()
		####print("name : " + str(name))
		a=str(name)
	

		profile = subprocess.getoutput("netsh wlan show profile " + str(name) + " key=clear").split()
		
		for i in range(len(profile)):

			if "콘텐츠" in profile[i]:
				####print("password : "+ str(profile[i+2]))
				b=str(profile[i+2])
		sql="""insert into registry6(idx,Name,Password) values (%s,%s,%s)"""
		curs.execute(sql,(count,a,b,))
		conn.commit()

		count += 1





def Information(para):
        date=para
        strdate = date[-2:] + date[-4:-2] + date[-6:-4] + date[-8:-6] + date[-10:-8] + date[-12:-10] + date[-14:-12] + date[-16:-14]
        dateint = int(strdate,16)
        us = dateint / 10.
        ####print(datetime(1601,1,1)+ timedelta(microseconds=us))
        return (datetime(1601,1,1)+ timedelta(microseconds=us))

def search(dirname):
    filenames = os.listdir(dirname)
    idx=1
    for filename in filenames:
        try:
            tmp=""
            flag=1
            full_filename = os.path.join(dirname, filename)
            d= os.path.join(dirname, filename)[56:]
            e=full_filename
            ####print ('파일이름 : ',d)
            ####print ('전체경로 : ',e)
            fd=open(full_filename,'rb')
            for i in range(1,os.path.getsize(full_filename)+1):
                data=fd.read(1)
                if (ord(data)>=16):
                    tmp += str(hex(ord(data))[2:])
                else:
                    tmp += str('0')
                    tmp += str(hex(ord(data))[2:])
                if (i%16==0 and i !=0):
                    tmp += str('\n')
            fd.close()
        
            a=str(tmp[57:65])+str(tmp[66:74]) #Creation_Time
            a=res=Information(a)
        
            b=str(tmp[74:90]) #Access_Time
            b=res=Information(b)
        
            c=str(tmp[90:98])+str(tmp[99:107]) #Write_Time
            c=res=Information(c)
        
            if(flag!=2):
                    sql="""insert into LNK(idx,file_name,full_path,Creation_Time,Access_Time,Write_Time) values (%s,%s,%s,%s,%s,%s)"""
                    curs.execute(sql,(idx,d,e,a,b,c,))
                    conn.commit()
                    idx+=1
            ####print('*'*100)
        except PermissionError :
                flag=2









def find_path(st,end):
    convert_path=""
    ggggg=0
    for q in range(st,end,2):
        if (tmp2[q]== " "):
            ggggg+=1
        else:
            convert_path+=tmp2[q]
    return(convert_path)










def run_as_admin(argv=None, debug=False):  # 관리자권한 실행
	shell32 = ctypes.windll.shell32
	if argv is None and shell32.IsUserAnAdmin():
		return True

	if argv is None:
		argv = sys.argv
	if hasattr(sys, '_MEIPASS'):
		# Support pyinstaller wrapped program.
		arguments = map(str, argv[1:])
	else:
		arguments = map(str, argv)
	argument_line = u' '.join(arguments)
	executable = str(sys.executable)
	if debug:
		print("Command line: ", executable, argument_line)
	ret = shell32.ShellExecuteW(
		None, u"runas", executable, argument_line, None, 1)
	if int(ret) <= 32:
		return False
	return None






def RecentWord():
    SID = []
    i = 0
    varkey = winreg.OpenKey(winreg.HKEY_USERS, r"")
    while True:
        try:
            #HKEY_USERS SID 구하기
            key = winreg.EnumKey(varkey, i)
            SID.append(key)
            i += 1
        except WindowsError:
            break
    winreg.CloseKey(varkey)
    Vers = []
    OfficePath = [SID[i] + "\\Software\\Microsoft\\Office" for i in range(len(SID))]
    #Office 버전 구하기
    varReg = winreg.ConnectRegistry(None, winreg.HKEY_USERS)
    for Opath in OfficePath:
        #.DEFAULT 건너뛰기
        if Opath.find(".DEFAULT") == 0:
            continue
        try:
            varkey = winreg.OpenKey(varReg, Opath, 0, winreg.KEY_ALL_ACCESS)
            i = 0
            while True:
                try:
                    Vers.append(winreg.EnumKey(varkey, i))
                    i += 1
                except:
                    break
            winreg.CloseKey(varkey)
            #최근 Word 파일 출력
            for vers in Vers:
                i = 0
                WordPath = Opath + "\\" + vers + "\\Word\\User MRU"
                #User MRU LiveID 값 가져오기
                varkey = winreg.OpenKey(varReg, WordPath, 0, winreg.KEY_ALL_ACCESS)
                Userkey = winreg.EnumKey(varkey, 0)
                #print("Office",vers,"Word")
                #가져온 LiveID에서 FILE MRU 접근   
                subkey = winreg.OpenKey(varReg, WordPath + "\\" + Userkey + "\\File MRU", 0, winreg.KEY_ALL_ACCESS)
                a="Word"
                b=vers
                while True:
                    try:
                        #File MRU 값들 가져오기
                        name, value, types = winreg.EnumValue(subkey, i)
                        value = value[value.find('*')+1:]
                        #print(name, value, types)
                        c=types
                        d=value
                        sql="""insert into registry7(idx,sort,version,type,path) values (%s,%s,%s,%s,%s)"""
                        curs.execute(sql,(i,a,b,c,d,))
                        conn.commit()
                        i += 1
                    except:
                        break
        except:
            pass


def RecenPowerPoint():
    SID = []
    i = 0
    varkey = winreg.OpenKey(winreg.HKEY_USERS, r"")
    while True:
        try:
            #HKEY_USERS SID 구하기
            key = winreg.EnumKey(varkey, i)
            SID.append(key)
            i += 1
        except WindowsError:
            break
    winreg.CloseKey(varkey)
    Vers = []
    OfficePath = [
    	SID[i] + "\\Software\\Microsoft\\Office" for i in range(len(SID))]
    #Office 버전 구하기
    varReg = winreg.ConnectRegistry(None, winreg.HKEY_USERS)
    for Opath in OfficePath:
        #.DEFAULT 건너뛰기
        if Opath.find(".DEFAULT") == 0:
            continue
        try:
            varkey = winreg.OpenKey(varReg, Opath, 0, winreg.KEY_ALL_ACCESS)
            i = 0
            while True:
                try:
                    Vers.append(winreg.EnumKey(varkey, i))
                    i += 1
                except:
                    break
            winreg.CloseKey(varkey)
            #최근 PowerPoint 파일 출력
            for vers in Vers:
                i = 0
                PowerPointPath = Opath + "\\" + vers + "\\PowerPoint\\User MRU"
                #User MRU LiveID 값 가져오기
                varkey = winreg.OpenKey(
                    varReg, PowerPointPath, 0, winreg.KEY_ALL_ACCESS)
                Userkey = winreg.EnumKey(varkey, 0)
                #print("Office", vers, "PowerPoint")
                a="PowerPoint"
                b=vers
                #가져온 LiveID에서 FILE MRU 접근
                subkey = winreg.OpenKey(
                    varReg, PowerPointPath + "\\" + Userkey + "\\File MRU", 0, winreg.KEY_ALL_ACCESS)
                while True:
                    try:
                        #File MRU 값들 가져오기
                        name, value, types = winreg.EnumValue(subkey, i)
                        value = value[value.find('*')+1:]
                        #print(name, value, types)
                        c=types
                        d=value
                        sql="""insert into registry7(idx,sort,version,type,path) values (%s,%s,%s,%s,%s)"""
                        curs.execute(sql,(i,a,b,c,d,))
                        conn.commit()
                        i += 1
                    except:
                        break
        except:
            pass


def RecentExcel():
    SID = []
    i = 0
    varkey = winreg.OpenKey(winreg.HKEY_USERS, r"")
    while True:
        try:
            #HKEY_USERS SID 구하기
            key = winreg.EnumKey(varkey, i)
            SID.append(key)
            i += 1
        except WindowsError:
            break
    winreg.CloseKey(varkey)
    Vers = []
    OfficePath = [
    	SID[i] + "\\Software\\Microsoft\\Office" for i in range(len(SID))]
    #Office 버전 구하기
    varReg = winreg.ConnectRegistry(None, winreg.HKEY_USERS)
    for Opath in OfficePath:
        #.DEFAULT 건너뛰기
        if Opath.find(".DEFAULT") == 0:
            continue
        try:
            varkey = winreg.OpenKey(varReg, Opath, 0, winreg.KEY_ALL_ACCESS)
            i = 0
            while True:
                try:
                    Vers.append(winreg.EnumKey(varkey, i))
                    i += 1
                except:
                    break
            winreg.CloseKey(varkey)
            #최근 Excel 파일 출력
            for vers in Vers:
                i = 0
                ExcelPath = Opath + "\\" + vers + "\\Excel\\User MRU"
                #User MRU LiveID 값 가져오기
                varkey = winreg.OpenKey(varReg, ExcelPath, 0, winreg.KEY_ALL_ACCESS)
                Userkey = winreg.EnumKey(varkey, 0)
                #print("Office", vers, "Excel")
                a="Excel"
                b=vers
                #가져온 LiveID에서 FILE MRU 접근
                subkey = winreg.OpenKey(
                    varReg, ExcelPath + "\\" + Userkey + "\\File MRU", 0, winreg.KEY_ALL_ACCESS)
                while True:
                    try:
                        #File MRU 값들 가져오기
                        name, value, types = winreg.EnumValue(subkey, i)
                        value = value[value.find('*')+1:]
                        #print(name, value, types)
                        c=types
                        d=value
                        sql="""insert into registry7(idx,sort,version,type,path) values (%s,%s,%s,%s,%s)"""
                        curs.execute(sql,(i,a,b,c,d,))
                        conn.commit()
                        i += 1
                    except:
                        break
        except:
            pass


def RecentHwp():
    SID = []
    i = 0
    varkey = winreg.OpenKey(winreg.HKEY_USERS, r"")
    while True:
        try:
            #HKEY_USERS SID 구하기
            key = winreg.EnumKey(varkey, i)
            SID.append(key)
            i += 1
        except WindowsError:
            break
    winreg.CloseKey(varkey)
    Vers = []
    HncPath = [
    	SID[i] + "\\Software\\HNC\\Common" for i in range(len(SID))]
    #HNC 버전 구하기
    varReg = winreg.ConnectRegistry(None, winreg.HKEY_USERS)
    for Hpath in HncPath:
        #.DEFAULT 건너뛰기
        if Hpath.find(".DEFAULT") == 0:
            continue
        try:
            varkey = winreg.OpenKey(varReg, Hpath, 0, winreg.KEY_ALL_ACCESS)
            i = 0
            while True:
                try:
                    Vers.append(winreg.EnumKey(varkey, i))
                    i += 1
                except:
                    break
            winreg.CloseKey(varkey)
            #최근 Hwp 파일 출력
            for vers in Vers:
                i = 0
                HwpPath = Hpath + "\\" + vers + "\\CommonFrame\\RecentFile"
                #User MRU LiveID 값 가져오기
                varkey = winreg.OpenKey(varReg, HwpPath, 0, winreg.KEY_ALL_ACCESS)
                #print("HNC", vers, "Hwp")
                a="Hwp"
                b=vers
                #RecentFile 값들 출력
                while True:
                    try:
                        name, data, types = winreg.EnumValue(varkey, i)
                        i += 1
                        if name =='Count':
                            continue
                        elif name.find("fix") == 0:
                            continue
                        #print(name, data.decode('utf-16'))
                        d=data.decode('utf-16')
                        sql="""insert into registry7(idx,sort,version,path) values (%s,%s,%s,%s)"""
                        curs.execute(sql,(i/2,a,b,d,))
                        conn.commit()
                    except:
                        break
        except:
            pass










def GetMFT():
    os.system(r"lib\forecopy_handy.exe -m ./")
    #$MFT 파일 forecopy_handy로 추출 
    #cmd = [sys.path[0]+"/lib/forecopy_handy.exe","-m","./"]
    #forecopy_handy 관리자 권한 실행
    #result = subprocess.run(cmd, shell=True)
    #qwer=1
    #if(result.returncode == 0):
    #    print("get mft success")
    #else:
    #    qwer+=1
    
def GetDeletedFile():
    os.system("lib\\analyzeMFT.exe -f mft\\$MFT -o result.csv")
    #cmd = [sys.path[0] + "/lib/analyzeMFT","-f" ,"mft\$MFT","-o", "result.csv"]
    #result = subprocess.run(cmd, shell=False)
    #qwe=1
    #if(result.returncode != 0):
    #        qwe+=1
    with open('result.csv', 'r', encoding="iso-8859-1") as f:
        #csv 파일 읽어오기
        reader = csv.DictReader(f, delimiter=',')
        #csv 파일 컬럼들 가져오기
        headers = reader.fieldnames

        '''
        ['Record Number', 'Good', 'Active', 'Record type', 'Sequence Number', 'Parent File Rec. #', 
        'Parent File Rec. Seq. #', 'Filename #1', 'Std Info Creation date', 'Std Info Modification date', 'Std Info Access date', 'Std InfoEntry date', 
        'FN Info Creation date', 'FN Info Modification date', 'FN Info Access date', 'FN Info Entry date','Object ID', 'Birth Volume ID', 
        'Birth Object ID', 'Birth Domain ID', 'Filename #2', 'FN Info Creation date', 'FN Info Modify date', 'FN Info Access date', 
        'FN Info Entry date', 'Filename #3', 'FN Info Creation date', 'FN Info Modify date', 'FN Info Access date', 'FN Info Entry date', 
        'Filename #4', 'FN Info Creation date', 'FN Info Modify date', 'FN Info Access date', 'FN Info Entry date', 'Standard Information', 
        'Attribute List', 'Filename','Object ID', 'Volume Name', 'Volume Info', 'Data', 
        'Index Root', 'Index Allocation', 'Bitmap', 'Reparse Point','EA Information', 'EA', 
        'Property Set', 'Logged Utility Stream', 'Log/Notes', 'STF FN Shift', 'uSec Zero', 'ADS']
        '''

        ii=0
        for row in reader:
            try:
                #corrupted 파일 제외
                if(row[headers[1]] == 'Zero' or row[headers[7]] == 'NoFNRecord'):
                    continue
                    #삭제된 파일
                if(row[headers[2]]=="Inactive"): #########################################갯수조절#####  ii로   ################3###
                    #print Record type(4), Filename(8), Std Info Creation date(9), Std Info Modification date(10), Std Info Access date(11)
                    #print(row[headers[3]], row[headers[7]], row[headers[8]], row[headers[9]], row[headers[10]])
                    a=row[headers[3]] #Record_type 파일인지 디렉토리인지
                    b=row[headers[7]] #Filename 경로 포함한 파일
                    c=row[headers[8]] #Creation_date 생성날짜
                    d=row[headers[9]] #Modification_date 수정날짜
                    e=row[headers[10]] #Access_date 접근날짜
                    sql="""insert into bin(idx,Record_type,Filename,Creation_date,Modification_date,Access_date) values (%s,%s,%s,%s,%s,%s)"""
                    curs.execute(sql,(ii,a,b,c,d,e,))
                    conn.commit()
                    ii+=1
            except WindowsError:
                break













if __name__ == "__main__":
    run_as_admin()
    USB_User = ""
    conn = pymysql.connect(host='203.229.206.16',port=12349,user='test',password='test',db='artifact')
    #conn = pymysql.connect(host='192.168.43.128',user='root',password='07cc24',db='cysun31')
    curs=conn.cursor()





    
    #drop table
    sql = """DROP TABLE IF EXISTS informationn"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS eventlog"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS registry1"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS registry2"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS registry3"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS registry4_Device_INFO"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS registry5"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS registry6"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS LNK"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS IconCache"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS weblog"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS registry7"""
    curs.execute(sql)
    conn.commit()
    #drop table
    sql = """DROP TABLE IF EXISTS bin"""
    curs.execute(sql)
    conn.commit()














    
    #create table
    sql = """
        create table informationn( idx integer(2) not null, OS_VERSION varchar(200), Detailed_VERSION varchar(90), USER varchar(100), Last_ShutDown_Time varchar(200),primary key (idx) );
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table eventlog( idx integer(8) not null, Event_Category integer(20) not null, Time_Generated varchar(25) not null, Source_Name varchar(100) not null, Event_ID integer(20),Event_Type integer(20) , Event_Data varchar(10000), primary key (idx) );
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table registry1( idx integer(8) not null, Name VARCHAR(200) not null, Session_Number VARCHAR(10), Run_Count VARCHAR(10), Last_Run_Time VARCHAR(30) , primary key(idx));;
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table registry2( idx integer(8) not null, caption varchar(200), command varchar(200), primary key (idx) );
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table registry3( idx integer(8) not null, USB_Name varchar(200), primary key (idx) );
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table registry4_Device_INFO(WpdBusEnumRoo_GUID varchar(200), Disk_GUID varchar(200),Volume_GUID varchar(200),Portable_Device_GUID varchar(200),USB_GUID varchar(200));
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table registry5( idx integer(8) not null, Connected_devices_name varchar(200),specific_name varchar(200),time varchar(20),primary key (idx) );
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table registry6( idx integer(8) not null, Name varchar(30), Password varchar(30), primary key (idx) );
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table LNK( idx integer(8) not null, file_name varchar(1000), full_path varchar(200),Creation_Time varchar(30),Access_Time varchar(30),Write_Time varchar(30), primary key (idx) );
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table IconCache( idx integer(8) not null, Path varchar(200) not null, primary key (idx));
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table weblog( idx integer(8) not null, URL varchar(500), Title varchar(200), Visit_Time varchar(25), Browser varchar(200), primary key (idx) );
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """
        create table registry7( idx varchar(20) not null, sort varchar(30), version varchar(30),type varchar(20),path varchar(250));
        """
    curs.execute(sql)
    conn.commit()
    #create table
    sql = """create table bin( idx varchar(20) not null, Record_type varchar(30), Filename varchar(300),Creation_date varchar(30),Modification_date varchar(30),Access_date varchar(30));"""
    curs.execute(sql)
    conn.commit()

    print("start gathering artifact")
    print("*"*100)






    
    Informationn()

    
    #이벤트로그 관련
    server = 'localhost' # name of the target computer to get event logs
    logtype = 'System' # 'Application' # 'Security'
    hand = win32evtlog.OpenEventLog(server,logtype)
    flags = win32evtlog.EVENTLOG_BACKWARDS_READ|win32evtlog.EVENTLOG_SEQUENTIAL_READ
    total = win32evtlog.GetNumberOfEventLogRecords(hand)

    i=1
    #데이터 읽어오면서 DB에 insert하기!
    while True:
        events = win32evtlog.ReadEventLog(hand, flags,0)
        if events:
            for event in events:
                if (i<7000):  ###################################################갯수조절###################3########     i    로 ##########
                        data = event.StringInserts
                        if data:
                            #print ('Event Data:')
                            f = ""
                            for msg in data:
                                f += str(msg)
                                f += "\n"
                            #print (f)

                        a=event.EventCategory            
                        b=str(event.TimeGenerated)
                        c=event.SourceName
                        d=event.EventID
                        e=event.EventType

                        sql="""insert into eventlog(idx,Event_Category,Time_Generated,Source_Name,Event_ID,Event_Type,Event_Data) values (%s,%s,%s,%s,%s,%s,%s)"""
                        curs.execute(sql,(i,a,b,c,d,e,f,))
                        conn.commit()
                        i+=1

                        #print('*' * 100)
        else:
            break;
        







    
    print("================================end of 02.py  ========================================")
    UserAssist()
    print("================================end of 03.py  =======================================")
    AutoRun()
    print("================================end of 04.py  =======================================")
    USB()
    print("================================end of 05.py  =======================================")
    Network()
    
    print("================================end of 06.py  =======================================")
    USERNAME=os.environ['USERNAME']
    search_path="C:/Users/"+USERNAME+"/AppData/Roaming/Microsoft/Windows/Recent/"
    search(search_path)
    print("================================end of 07.py  =======================================")







    
    tmp=""
    tmp2=""
    list=[]
    path=[]
    list_of_full_path=[]
    list_of_full_path22=[]
    list_of_full_path33=[]
    list_of_full_path44=[]


    #filename='C:/Users/cysun/AppData/Roaming/Microsoft/Windows/Recent/002.gif.lnk'
    #filename='C:/Users/cysun/OneDrive/바탕 화면/asas/IconCache.db'
    USERNAME=os.environ['USERNAME']
    filename="C:/Users/"+USERNAME+"/AppData/Local/IconCache.db"

    fd=open(filename,'rb')
    #print ("size : "+str(os.path.getsize(filename)))
    for i in range(1,os.path.getsize(filename)+1):
        data=fd.read(1)
        if (ord(data)>=16):
            tmp += str(hex(ord(data))[2:])
            tmp2+=chr(ord(data))
        else:
            tmp += str('0')
            tmp += str(hex(ord(data))[2:])
            tmp2+=chr(ord(data))


    lenn=len(tmp)

    for i in range(0,lenn,2):
        temp=""
        temp+=tmp[i]
        temp+=tmp[i+1]
        list.append(temp)

    count=0
    try:
            for j in range(0,os.path.getsize(filename)-10000,1):
                #print (list[j])
                if (list[j]=='63' and list[j+2]=='3a'): # "c:" 문자열 찾는 if
                    #print(j)
                    for k in range(200):
                        #print(list[j+k])
                        if(list[j+k]=='ff' and list[j+k+1]=='ff' and list[j+k+2]=='ff'):
                            path.append((j,j+k-1))
                            count+=1
                            break;
    except IndexError :
            print("err")

    fd.close()

    cnt=1
    for x in range (count-1):
        full_path=""
        full_path=find_path(path[x][0],path[x][1])
        #print(full_path)
        list_of_full_path.append(full_path)
        #sql="""insert into IconCache(idx,Path) values (%s,%s)"""
        #curs.execute(sql,(cnt,full_path,))
        #conn.commit()
        cnt+=1
    #print(list_of_full_path)
    for x in list_of_full_path:
            lengthh=len(x)
            aa=""
            for y in range(lengthh):
                    aa+=x[y]
                    #if(x[y]="." and x[y+1]="d" and x[y+2]="l" and x[y+3]="l"):
                    if(x[y]=="."):
                            if(x[y+1]=="d"):
                                    aa+="dll"
                                    break
                            elif(x[y+1]=="e"):
                                    aa+="exe"
                                    break
                            elif(x[y+1]=="c"):
                                    aa+="cpl"
                                    break
                            elif(x[y+1]=="l"):
                                    aa+="lnk"
                                    break
                            elif(x[y+1]=="i"):
                                    aa+="ico"
                                    break
            #print(aa)
            list_of_full_path22.append(aa)
    
    ccnntt=1
    #print(type(list_of_full_path22))
    list_of_full_path33 = collections.OrderedDict.fromkeys(list_of_full_path22).keys()
    for z in list_of_full_path33:
            sql="""insert into IconCache(idx,Path) values (%s,%s)"""
            curs.execute(sql,(ccnntt,z,))
            conn.commit()
            ccnntt+=1
            
                    
    print("================================end of 08.py  =======================================")











    
    os.system("lib\BHV.exe /HistorySource 1 /LoadIE 1 /LoadFirefox 1 /LoadChrome 1 /LoadSafari 1 /stab lib/weblog.data")

    fff = codecs.open('lib/weblog.data', 'r', encoding='utf16')
    data = fff.read()
    data[1:]

    stream = data[1:].split('\n')
    logg = [i.split('\t') for i in stream]
    roww = len(logg)
    fff.close()
    count=1

    for i in range(2,roww-1):
            a=logg[i][0]
            b=logg[i][1]
            c=logg[i][2]
            d=logg[i][6]
            if (len(a)<500): ########################################갯수조절##########################     i 로     ###################
                    sql="""insert into weblog(idx,URL,Title,Visit_Time,Browser) values (%s,%s,%s,%s,%s)"""
                    curs.execute(sql,(count,a,b,c,d,))
                    conn.commit()
                    count+=1





    print("================================end of 09.py  =======================================")
    RecentWord()
    RecenPowerPoint()
    RecentExcel()
    RecentHwp()
    print("================================end of 10.py  =======================================")
    GetMFT()
    GetDeletedFile()
    print("================================end of 11.py  =======================================")


