import cv2
import pickle
import numpy as np
import time
import random
import os
import csv
from cvzone.HandTrackingModule import HandDetector
from sklearn.neighbors import KNeighborsClassifier
from datetime import datetime
from win32com.client import Dispatch
import sys

print('+'+"\033[1;91m-" * 22 + "\033[1;93m-" * 21 + "\033[1;92m-" * 21 + "\033[1;94m-" * 21 + "\033[1;95m-" * 21 + "\033[1;96m-" * 21 + "\033[0m+")

print("|\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t|\n|\t \t \t \t\t\t \033[1;34m|[ ■..\033[0m\033[1;4;38;5;208m.ATTENDANCE SYSTEM.\033[0m\033[1;34m..■ ]|\033[0m\t\t\t\t\t\t|")
print("|\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t|\n|\t\033[1;38;5;207m✿\033[0m  \033[1;4;33mDASHBOARD\033[0m \033[1;38;5;207m✿\033[0m \t\t\t\t\t\t\t\t\t\t\t\t\t\t|")
option=input("|\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t|\n|\t\033[1;96m•\033[0m  Press : 'A' or 'a' for attendance\t\t\t\t\t\t\t\t\t\t\t|\n|\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t|\n|\t\033[1;96m•\033[0m  Press : 'S' or 's' for attendance status\t\t\t\t\t\t\t\t\t\t|\n|\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t|\n|\t\033[1;96m•\033[0m  Press : 'E' or 'e' for exit\t\t\t\t\t\t\t\t\t\t\t\t|\n|\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t|\n|\tYou press : \t\t\t\t\t\t\t\t\t\t\t\t\t\t|\n|\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t|\n\t\033[1;93m>>> \033[0m")

print('+'+"\033[1;91m-" * 22 + "\033[1;93m-" * 21 + "\033[1;92m-" * 21 + "\033[1;94m-" * 21 + "\033[1;95m-" * 21 + "\033[1;96m-" * 21 + "\033[0m+")


if option.lower() == 'a':
    def speak(str1):
        speak = Dispatch(("SAPI.SpVoice"))
        speak.Speak(str1)

    detector = HandDetector(detectionCon=0.9, maxHands=2)

    video = cv2.VideoCapture(0)

    facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

    with open('data/names.pkl', 'rb') as w:
        LABELS = pickle.load(w)
    with open('data/faces_data.pkl', 'rb') as f:
        FACES = pickle.load(f)

    knn = KNeighborsClassifier(n_neighbors=5)
    knn.fit(FACES, LABELS)

    fingercount = []

    while True:
        ret, frame = video.read()

        hands, _ = detector.findHands(frame)

        if hands:
            hands1 = hands[0]
            fingercount = detector.fingersUp(hands1)
            print(fingercount)

            if fingercount == [0, 0, 0, 0, 0]:
                video.release()
                cv2.destroyAllWindows()
                break

        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        faces = facedetect.detectMultiScale(gray, 1.3, 5)

        for (x, y, w, h) in faces:
            crop_img = frame[y:y + h, x:x + w, :]
            resized_img = cv2.resize(crop_img, (50, 50)).flatten().reshape(1, -1)
            output = knn.predict(resized_img)
            ts = time.time()
            date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
            timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
            exist = os.path.isfile("Attendance/Attendance_" + date + ".csv")
            cv2.rectangle(frame, (0, 0), (350, 60), (0, 0, 0), -2, cv2.LINE_AA)
            cv2.putText(frame, 'Name : {}'.format(str(output[0])), (50, 50), cv2.FONT_HERSHEY_COMPLEX, 1,
                        (255, 255, 255), 1)

            attendance = {'Name': str(output[0]), 'Time': str(timestamp)}

        cv2.imshow("Frame", frame)
        q = cv2.waitKey(1)
        if q == ord('p') or fingercount == [1, 1, 1, 1, 1]:
            speak("Attendance Taken..")
            print("\n>>> Attendance Taken !!! <<<\n")
            time.sleep(2)
            if exist:
                with open("Attendance/Attendance_" + date + ".csv", "+a", newline='') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=['Name',': Roll_NO.', 'Time'])
                    if os.stat("Attendance/Attendance_" + date + ".csv").st_size == 0:
                        writer.writeheader()
                    writer.writerow(attendance)
            else:
                with open("Attendance/Attendance_" + date + ".csv", "+a", newline='') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=['Name',': Roll_NO.', 'Time'])
                    writer.writeheader()
                    writer.writerow(attendance)
            break

    video.release()
    cv2.destroyAllWindows()



elif option.lower() == 's':
    n = input("Enter your name : ")
    ts = time.time()
    date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
    attendance_file_path = "Attendance/Attendance_" + date + ".csv"
    if os.path.isfile(attendance_file_path):
        with open(attendance_file_path, 'r') as csvfile:
            reader = csv.DictReader(csvfile)
            print(f"\nAttendance details of {n}.\n")
            print("{:<20} {:<20}".format('Name', 'Time'))
            for row in reader:
                if n in row['Name']:
                    print("{:<20} {:<20}".format(row['Name'], row['Time']))
                    print(f"\n--> Name   : {n}\n--> Status : Present\n")
                    print("\n# NOTE --> Please check your 'NAME' and 'ROLL_NO.' carefully..\n")

    else:
        print(f"No attendance data found for {date}")


elif option.lower() == 'e':
    sys.exit()


else:
    print("\n\033[1;91m# Warning\033[0m : Invalid option!!!\n\nPlease press 'A' or 'S' or 'E'\n")