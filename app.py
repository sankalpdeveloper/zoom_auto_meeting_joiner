import xlrd
from datetime import datetime
import webbrowser
from time import time, sleep
from win10toast import ToastNotifier 
  
# create an object to ToastNotifier class 
n = ToastNotifier() 

file = ("zoom.xlsx")
wb = xlrd.open_workbook(file)
sheet = wb.sheet_by_index(0)
list1 = []
for i in range(1, 7):
    add = sheet.row_values(i)
    list1.append(add)

print(list1)
datetime_obj = datetime.now().ctime()
day = datetime_obj[0:3].upper()
hrs = datetime_obj[11:13]
min = datetime_obj[14:16]
url= []
teacher = []
print(f"Today is {day} and You have classes of following teacher:")
for i in range(0,len(list1)):
    m = list(list1[i][3].split(","))
    if(day in m):
        n.show_toast(f"Today's Zoom Classes",list1[i][1]+ " @ " + list1[i][0]+":00 AM", duration = 10,icon_path ="zoom.ico")
        print(list1[i][1]+ " @ " + list1[i][0]+":00 AM" )
        teacher.append(list1[i][1])
        url.append(list1[i][2])

while True:
    datetime_obj = datetime.now().ctime()
    hrs = datetime_obj[11:13]
    min = datetime_obj[14:16]
    time = int(hrs)*100+int(min)
    if(time >= int(955) and time <= int(1040)):
        print(f"Opening of class {teacher[0]} ")
        webbrowser.open(url[0])
        sleep(60*61)
    elif(time >= int(1055) and time <= int(1140)):
        print(f"Opening of class {teacher[1]} ")
        webbrowser.open(url[1])
        sleep(60*61)
    elif(time >= int(1155) and time <= int(1240)):
        print(f"Opening of class {teacher[2]}")
        webbrowser.open(url[2])
        sleep(60*61)
    else:
        inclass = True
        print("Trying...")
        sleep(2*60)