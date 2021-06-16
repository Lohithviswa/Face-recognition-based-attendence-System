import cv2
import numpy as np
import face_recognition
import os
from datetime import datetime,date
from mail import sendmail
import tkinter as tk
from tkinter import filedialog
import pandas as pd

# from PIL import ImageGrab
roll_numbers=[]
root= tk.Tk()
 
canvas1 = tk.Canvas(root, width = 500, height = 500, bg = 'lightsteelblue')
canvas1.pack()

def getExcel ():
    global df
    
    import_file_path = filedialog.askopenfilename()
    df = pd.read_excel (import_file_path)
    #print(roll_numbers)
    print("file opened")
    
browseButton_Excel = tk.Button(text='Class Roll(Import Excel File)', command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(250, 250, window=browseButton_Excel)
root.mainloop()

roll_numbers=list(df['Roll_numbers']) 
#print(roll_numbers)
path = 'ImagesAttendance'
images = []
classNames = []
myList = os.listdir(path)
print(myList)
for cl in myList:
    curImg = cv2.imread(f'{path}/{cl}')
    images.append(curImg)
    classNames.append(os.path.splitext(cl)[0])
print(classNames)


def markAbsentees():
    with open('Attendance.csv','r+') as f:
        myDataList = f.readlines()
        nameList = []
        for line in myDataList:
            entry = line.split(',')
            nameList.append(entry[0])
    with open('Absentees.csv','r+') as f:
        f.writelines('Roll numbers,Date\n')
        today=date.today()
        
        for roll in roll_numbers:
            if roll not in nameList:
                f.writelines(f'{roll},{today}\n')
                print(roll)
                
#     return capScr
def findEncodings(images):
    encodeList = []
    for img in images:
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        encode = face_recognition.face_encodings(img)[0]
        encodeList.append(encode)
    return encodeList

def markAttendance(name):
    #print(name)
    with open('Attendance.csv','r+') as f:
        myDataList = f.readlines()
        nameList = []
        for line in myDataList:
            entry = line.split(',')
            nameList.append(entry[0])
        if name not in nameList and name!=' ':
            now = datetime.now()
            today=date.today()
            dtString = now.strftime('%H:%M:%S')
            f.writelines(f'\n{name},{dtString},{today}')
            print(name)
            
encodeListKnown = findEncodings(images)
#print(encodeListKnown)
print('Encoding Completed')
 
cap = cv2.VideoCapture(0)
try:
    while True:
        success, img = cap.read()
        #cv2.imshow('Webcam1',img)
        #img = captureScreen()
        imgS = cv2.resize(img,(0,0),None,0.25,0.25)
        imgS = cv2.cvtColor(imgS, cv2.COLOR_BGR2RGB)
     
        facesCurFrame = face_recognition.face_locations(imgS)
        encodesCurFrame = face_recognition.face_encodings(imgS,facesCurFrame)
     
        for encodeFace,faceLoc in zip(encodesCurFrame,facesCurFrame):
            matches = face_recognition.compare_faces(encodeListKnown,encodeFace)
            faceDis = face_recognition.face_distance(encodeListKnown,encodeFace)
            #print(faceDis)
            matchIndex = np.argmin(faceDis)
     
            if matches[matchIndex]:
                name = classNames[matchIndex].upper()
                #print(name)
                #print(faceLoc)
                
                y1,x2,y2,x1 = faceLoc
                #print(y1,x2,y2,x1)
                y1, x2, y2, x1 = y1*4+4,x2*4+4,y2*4+4,x1*4+4
                cv2.rectangle(img,(x1,y1),(x2,y2),(0,255,0),2)
                cv2.rectangle(img,(x1,y2-35),(x2,y2),(255,0,0),cv2.FILLED)
                cv2.putText(img,name,(x1+6,y2-6),cv2.FONT_HERSHEY_COMPLEX,1,(255,255,255),2)
                cv2.imshow('Webcam',img)
                markAttendance(name)
            
        #markAbsentees()       
        if cv2.waitKey(1) & 0xFF == ord('q'):
               #sendmail()
               break
        
        
        #cv2.waitKey(1)
    
    cap.release() 
    cv2.destroyAllWindows() 
        
except Exception as e:
    print("got exception"+str(e))