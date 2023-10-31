import cv2
import numpy as np
from pyzbar.pyzbar import decode
from playsound import playsound
import openpyxl
import datetime as dt
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

record = cv2.VideoCapture(0)
record.set(3,640)
record.set(4,480)

present_data=[]

while(1):
    s, img = record.read()
    for qr in decode(img):
        data=qr.data.decode('utf-8')
        #print(data)

        #box around
        pts = np.array([qr.polygon],np.int32)
        pts = pts.reshape((-1,1,2))
        
        #display on image
        pts2 = qr.rect
        
        if data not in present_data:
            present_data.append(data)
            #print(10*'\a')
            playsound('beep.mp3')
            color=(0,0,255)
            cv2.polylines(img,[pts],1,color,5)
            #cv2.waitKey(3)
        else:
            color=(0,255,0)    
            cv2.polylines(img,[pts],1,color,5)
            cv2.putText(img,"Thank You! {}".format(data),(pts2[0],pts2[1]),cv2.FONT_HERSHEY_PLAIN, 1,color,2)

    cv2.imshow('Result',img)
    key = cv2.waitKey(1)

    #to release camera frame
    if(key==27): #esc key value is 27
        cv2.destroyWindow("Result")
        record.release()
        break
    
print(present_data)

#offline
# temp = openpyxl.load_workbook("Attendance_Sheet.xlsx")
# sheet = temp.active
# column = sheet.max_column
# sheet.cell(1,column+1).value = dt.date.today()
# for i in range(2,sheet.max_row+1):
#     x=sheet.cell(i,2)
#     if(x.value in present_data):
#        sheet.cell(i,column+1).value="P"
#     else:
#         sheet.cell(i,column+1).value="Ab"

# temp.save("Attendance_Sheet.xlsx")

#online Google Sheets
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('attendance.json', scope)
client = gspread.authorize(creds)
wb = client.open('MERL v2.0 Attendance Sheet')
sheet = wb.get_worksheet(0)

#column = sheet.cell(100,2).value
max_row = len(sheet.get_all_values())
max_col = len(sheet.get_all_values()[0])

#print(column,type(column))
#max_col=int(column)
#print(max_col,type(max_col))

to_day = str(dt.date.today())
#print(to_day,type(to_day))

sheet.update_cell(1,max_col+1,to_day)

#sheet.cell(1,column+1).value = dt.date.today()

for i in range(2,max_row+1):
    print(f"sheet.cell(i,2){sheet.cell(i,2)}")
    x=sheet.cell(i,1)
    print(f"x.value{x.value}")
    if(x.value in present_data):
        sheet.update_cell(i,max_col+1,"P")
    else:
        sheet.update_cell(i,max_col+1,"Ab")
#sheet.update_cell(100,2,max_col+1)
