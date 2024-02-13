import cv2 
import mediapipe as mp 
import time 
import tkinter 
from tkinter import messagebox
import numpy as np 
#import inquirer
import math 
import pandas as pd 
from xlutils.copy import copy
import matplotlib.pyplot as plt 
from matplotlib import style
import datetime as dt


def mouseclick(event, x, y, flags, param): 
  
    global Dificulty,InitialMenu, Xpoints, Ypoints
     
    if event == cv2.EVENT_LBUTTONDOWN: 

        if (x>40 and x<200) and (y>375 and y<450):
            Xpoints =  np.arange(0,440)
            Ypoints =  100*np.sin(Xpoints/28)+240
            Ypoints = Ypoints.astype(int)
            Xpoints = Xpoints + 100
            InitialMenu=1
            Dificulty = 'Easy'

        if (x>240 and x<400) and (y>375 and y<450):
            Xpoints =  np.arange(0,440)
            Ypoints =  100*np.sin(1.4*Xpoints/28)+240
            Ypoints = Ypoints.astype(int)
            Xpoints = Xpoints + 100
            InitialMenu=1
            Dificulty = 'Medium'

        if (x>440 and x<600) and (y>375 and y<450):
            Xpoints =  np.arange(0,440)
            Ypoints =  100*np.sin(2*Xpoints/28)+240
            Ypoints = Ypoints.astype(int)
            Xpoints = Xpoints + 100
            InitialMenu=1
            Dificulty = 'Hard'

    return Dificulty, InitialMenu,Xpoints,Ypoints

if __name__=="__main__": 

    DATE = dt.datetime.today().strftime("%d/%m/%Y")
    Attempt = 0

    while True:
        ID = input("Insert your ID: ") 
        
        if ID.isdigit() and int(ID)>0:
            break
        else:
            print("Invalid syntax. You need a digit greater than 0!")


    ExcelFile_Sheet = pd.read_excel("Results.xlsx") 

    if len(ExcelFile_Sheet.index)>0:
        if int(ID) in ExcelFile_Sheet['ID'].values:
            for row in range(len(ExcelFile_Sheet.index)):
                if int(ID) == ExcelFile_Sheet['ID'][row]:
                    New_Patient = 0
                    Name_Patient = ExcelFile_Sheet['NAME'][row]
                    age = ExcelFile_Sheet['AGE'][row]
                    Attempt +=1
        else:
            New_Patient=1
    else:
        New_Patient=1
    
    if New_Patient==1:
        Name_Patient = input("Hi there! Insert your name: ") 
        Attempt = 0

        while True:
            age = input("Insert your age: ") 
            
            if age.isdigit() and int(age)>0:
                break
            else:
                print("Invalid number!")
    else:
        print(f"\nHi again, {Name_Patient}, great to have you back!")


    tkinter.Tk().geometry("200x200")
    advance = messagebox.askokcancel(title="Important Information", message="This project is being developed by Biomedical Engineering MsC Students from the Instituto Superior de Engenharia do Porto. The computer camera will be turned on but your image will NOT be stored.")

    if advance == True:

        mpHand = mp.solutions.hands 
        hands = mpHand.Hands(max_num_hands=1) 

        mpDrawing = mp.solutions.drawing_utils

        video = cv2.VideoCapture(0) 


       
        InitialTime = 0
        
        
        global InitialMenu 
        InitialMenu = 0
        k=0
        VideoCapture1=0
        VideoCapture2=0
        InitialFingerCoordinates=[0,0]
        mask1 = 0
        mask3 = 0
        mask5 = 0
        mask6 = 0
        draw = 0
        b=0
        a=0
        MSE = 0. 
        cycle = 0
        TotalMSE = 0.
        start_counting=0
        HitPercentage=0
        Seconds=0

        global Dificulty, Xpoints, Ypoints
        Dificulty = ''
        Xpoints=[]
        Ypoints=[]



        while (cv2.waitKey(1) != 27): 
            
            open, VideoCapture = video.read() 
            VideoCapture = cv2.flip(VideoCapture, 1) 
            videoWidth = video.get(cv2.CAP_PROP_FRAME_WIDTH) 
            videoHeight = video.get(cv2.CAP_PROP_FRAME_HEIGHT) 

            mask = np.zeros_like(VideoCapture) 
            mask2 = np.zeros_like(VideoCapture)
            mask4 = np.zeros_like(VideoCapture)

            if not open: 
                break


            TimeNow=time.time() 
            FPS = int(1/(TimeNow-InitialTime)) 
            InitialTime = TimeNow 
            

            result = hands.process(cv2.cvtColor(VideoCapture, cv2.COLOR_BGR2RGB)) 

            ButtonInitiate_Coordinates = [70,220] 
            ButtonStop_Coordinates = [580,220]
            ButtonExit_Coordinates = [590,40]
            ButtonReset_Coordinates = [520,40]
            ButtonDificulty_Coordinates = [450,40]

            VideoCapture1 = VideoCapture 
            
            if InitialMenu==0:
                cv2.rectangle(VideoCapture1,(40,375),(200,450),(0,255,0),-1)
                cv2.rectangle(VideoCapture1,(240,375),(400,450),(0,255,255),-1)
                cv2.rectangle(VideoCapture1,(440,375),(600,450),(0,0,255),-1)
                cv2.putText(VideoCapture1, "EASY",[85,420], cv2.FONT_HERSHEY_TRIPLEX, 0.7, (255,255,255), 1)
                cv2.putText(VideoCapture1, "MEDIUM",[285,420], cv2.FONT_HERSHEY_TRIPLEX, 0.7, (0,0,0), 1)
                cv2.putText(VideoCapture1, "HARD",[485,420], cv2.FONT_HERSHEY_TRIPLEX, 0.7, (255,255,255), 1)
                cv2.putText(VideoCapture1, "Hello! Please select the dificulty level!",[100,60], cv2.FONT_HERSHEY_TRIPLEX, 0.5, (255,255,255), 1)
            
                cv2.namedWindow("Project")
                cv2.setMouseCallback("Project", mouseclick)
            
            else:

                if result.multi_hand_landmarks: 

                    Finger = result.multi_hand_landmarks[0].landmark[mpHand.HandLandmark.INDEX_FINGER_TIP]  
                    FinalFingerCoordinates = mpDrawing._normalized_to_pixel_coordinates(Finger.x, Finger.y, videoWidth, videoHeight)  
                        
                    if (FinalFingerCoordinates!=None):
                    
                        cv2.circle(VideoCapture1,ButtonInitiate_Coordinates,30,(0,255,0),-1)
                        cv2.putText(VideoCapture1, "Go",[ButtonInitiate_Coordinates[0]-22,ButtonInitiate_Coordinates[1]+2], cv2.FONT_HERSHEY_TRIPLEX, 0.4, (0,0,0), 1)
                        cv2.circle(VideoCapture1,ButtonStop_Coordinates,30,(0,0,255),-1)
                        cv2.putText(VideoCapture1, "Stop",[ButtonStop_Coordinates[0]-20,ButtonStop_Coordinates[1]+5], cv2.FONT_HERSHEY_TRIPLEX, 0.4, (0,0,0), 1)
                        cv2.circle(VideoCapture1,ButtonExit_Coordinates,30,(0,0,0),-1)
                        cv2.putText(VideoCapture1, "X",[ButtonExit_Coordinates[0]-3,ButtonExit_Coordinates[1]+5], cv2.FONT_HERSHEY_TRIPLEX, 0.5, (255,255,255), 1)
                        cv2.circle(VideoCapture1, ButtonReset_Coordinates, 30, (255,0,0), -1) 
                        cv2.putText(VideoCapture1, "Clean",[ButtonReset_Coordinates[0]-22,ButtonReset_Coordinates[1]+6], cv2.FONT_HERSHEY_TRIPLEX, 0.4, (255,255,255), 1)
                        cv2.circle(VideoCapture1, FinalFingerCoordinates, 2, (76,171,206), 3) 
                        cv2.circle(VideoCapture1, ButtonDificulty_Coordinates, 30, (255,255,255), -1) 
                        cv2.putText(VideoCapture1, "Change",[ButtonDificulty_Coordinates[0]-22,ButtonDificulty_Coordinates[1]+6], cv2.FONT_HERSHEY_TRIPLEX, 0.4, (0,0,0), 1)
                        cv2.putText(VideoCapture1, f"Dificulty: {Dificulty}",[10,15], cv2.FONT_HERSHEY_TRIPLEX, 0.4, (0,0,0), 1)
        
                        GoEquation = (ButtonInitiate_Coordinates[0]-FinalFingerCoordinates[0])**2+(ButtonInitiate_Coordinates[1]-FinalFingerCoordinates[1])**2
                        if (GoEquation)<800:
                            draw = 1
                            Begin_time = time.time()
                            start_counting = 1
                            

                        StopEquation = (ButtonStop_Coordinates[0]-FinalFingerCoordinates[0])**2+(ButtonStop_Coordinates[1]-FinalFingerCoordinates[1])**2
                        if (StopEquation)<800:
                            draw = 0
                            k = 0
                            if start_counting == 1:
                                start_counting = 0
                                Stop_time = time.time()
                                Seconds = Stop_time - Begin_time
                                Seconds = round(Seconds,2)

                        LeaveEquation = (ButtonExit_Coordinates[0]-FinalFingerCoordinates[0])**2+(ButtonExit_Coordinates[1]-FinalFingerCoordinates[1])**2
                        if LeaveEquation<800:
                            if start_counting == 0:
                                Stop_time = time.time()
                                Seconds = Stop_time - Begin_time
                                Seconds = round(Seconds,2)
                            Exit =1
                            break

                        CleanEquation = (ButtonReset_Coordinates[0]-FinalFingerCoordinates[0])**2+(ButtonReset_Coordinates[1]-FinalFingerCoordinates[1])**2
                        if CleanEquation<800:
                            mask1 = np.zeros_like(VideoCapture)
                            mask4 = np.zeros_like(VideoCapture)
                            mask2 = np.zeros_like(VideoCapture)
                            mask6 = 0

                            mask3 = 0
                            mask5 = 0
                            draw = 0
                            k = 0 
                            MSE = 0.
                            TotalMSE = 0.
                            cycle = 0
                            start_counting = 0
                            a=0

                        DificultyEquation = (ButtonDificulty_Coordinates[0]-FinalFingerCoordinates[0])**2+(ButtonDificulty_Coordinates[1]-FinalFingerCoordinates[1])**2
                        if (DificultyEquation)<800:
                            InitialMenu = 0
                            test = np.zeros_like(VideoCapture)
                            if ((mask5|test).any()!=0) and cycle>0 and Seconds!=0:
                                Attempt+=1
                                TotalMSE = TotalMSE/cycle
                                TotalMSE = round(TotalMSE,2)
                                HitPercentage = (np.sum(mask6&mask3)/np.sum(mask3))*100                            
                                HitPercentage = round(HitPercentage,2)
                                FILENAME = ID + '_' + Name_Patient + '_' + Dificulty + '_'+ str(Attempt) + 'Attempt.jpg' 
                                cv2.putText(mask5, f"MSE: {str(TotalMSE)} ", (25,400), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                                cv2.putText(mask5, f"Age: {str(age)} ", (25,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                                cv2.putText(mask5, f"Name: {str(Name_Patient)} ", (220,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                                cv2.putText(mask5, f"Dificulty: {str(Dificulty)} ", (400,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                                cv2.putText(mask5, f"Drawing: ", (25,120), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                                cv2.putText(mask5, f"Time: {str(Seconds)} s", (350,400), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                                cv2.putText(mask5, f"Hit Percentage: {str(HitPercentage)} %", (350,430), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                                cv2.putText(mask5, f"Date: {str(DATE)}", (350,460), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                                cv2.circle(mask5,(50,425),5,(0,255,0),-1)
                                cv2.putText(mask5, "MSE < 50", (57,428), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  
                                cv2.circle(mask5,(50,445),5,(0,255,255),-1)
                                cv2.putText(mask5, "50 =< MSE < 70", (57,448), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  
                                cv2.circle(mask5,(50,465),5,(0,0,255),-1)
                                cv2.putText(mask5, "MSE >= 70", (57,468), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  

                                newline = pd.DataFrame([[str(ID),Name_Patient,str(age), Dificulty,str(TotalMSE), str(HitPercentage), str(Seconds),str(DATE),str(Attempt)]], columns=['ID','NAME','AGE','DIFICULTY','MSE','HIT %','TIME (s)','DATE','ATTEMPT'])
                                newtable = pd.concat([ExcelFile_Sheet,newline])

                                newtable.to_excel('Results.xlsx', index=False)

                                cv2.imwrite(FILENAME, mask5) 
                            
                            mask1 = np.zeros_like(VideoCapture)
                            mask4 = np.zeros_like(VideoCapture)
                            mask2 = np.zeros_like(VideoCapture)

                            mask3 = 0
                            mask5 = 0
                            mask6 = 0
                            draw = 0
                            k = 0 
                            MSE = 0.
                            TotalMSE = 0.
                            cycle = 0
                            start_counting = 0
                            a=0


                        if draw == 1:  
                            if k<1:
                                InitialFingerCoordinates = FinalFingerCoordinates

                                while a<439:  
                                    PointsCoordinates = [Xpoints[a],Ypoints[a]]
                                    mask2 = cv2.line(mask2,PointsCoordinates,[Xpoints[a+1],Ypoints[a+1]],(255,255,255),1)
                                    mask3 = mask3 + mask2
                                    a=a+1

                                k+=1
                                
                            elif k==1:
                                i=0
                                MSE=0
                                points =  np.arange(40,610)
                                if FinalFingerCoordinates[0] in points:
                                    mask = cv2.line(mask, FinalFingerCoordinates, InitialFingerCoordinates, (255,255,255), 5)  
                                    
                                    if FinalFingerCoordinates[0] in Xpoints:

                                        i = list(Xpoints).index(FinalFingerCoordinates[0])
                                        MSE = math.fabs(FinalFingerCoordinates[1]-Ypoints[i])
                                        MSE = float(MSE**2)

                                        if MSE <50:
                                            mask4 = cv2.line(mask4, FinalFingerCoordinates, InitialFingerCoordinates, (0,255,0), 3)  
                                        elif (MSE >=50 and MSE <70):
                                            mask4 = cv2.line(mask4, FinalFingerCoordinates, InitialFingerCoordinates, (0,255,255), 3)  
                                        else:
                                            mask4 = cv2.line(mask4, FinalFingerCoordinates, InitialFingerCoordinates, (0,0,255), 3)  

                                        mask6 = mask6 + mask
                                        mask5 = mask5 + mask4 + mask3

                                    mask1 = mask1 + mask + mask3
                                    VideoCapture1 = cv2.add(VideoCapture1,mask1)
                                    InitialFingerCoordinates = FinalFingerCoordinates

                                cycle += 1
                                TotalMSE = float(TotalMSE + MSE)
                                
                    else:
                        print("Your finger is not visible!")

                else:
                    cv2.putText(VideoCapture1, "Hand not fully visible", (100,50), cv2.FONT_HERSHEY_TRIPLEX, 0.8, (0,0,0), 1)  


            cv2.putText(VideoCapture1, f"FPS: {str(FPS)}", (10,475), cv2.FONT_HERSHEY_TRIPLEX, 0.5, (76,171,206), 1)  
            cv2.imshow("Project", VideoCapture1) 
        
        else:
            Exit = 0

        if Exit == 1:
            if cycle>0:
                Attempt+=1
                TotalMSE = TotalMSE/cycle
                TotalMSE = round(TotalMSE,2)
                HitPercentage = (np.sum(mask6&mask3)/np.sum(mask3))*100                            
                HitPercentage = round(HitPercentage,2)
                FILENAME = ID + '_' + Name_Patient + '_' + Dificulty + '_'+ str(Attempt) + 'Attempt.jpg' 
                cv2.putText(mask5, f"MSE: {str(TotalMSE)} ", (25,400), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                cv2.putText(mask5, f"Age: {str(age)} ", (25,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)   
                cv2.putText(mask5, f"Name: {str(Name_Patient)} ", (220,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)   
                cv2.putText(mask5, f"Dificulty: {str(Dificulty)} ", (400,50), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                cv2.putText(mask5, f"Drawing: ", (25,120), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                cv2.putText(mask5, f"Time: {str(Seconds)} s", (350,400), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                cv2.putText(mask5, f"Hit Percentage: {str(HitPercentage)} %", (350,430), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                cv2.putText(mask5, f"Date: {str(DATE)}", (350,460), cv2.FONT_HERSHEY_TRIPLEX, 0.6, (255,255,255), 1)  
                cv2.circle(mask5,(50,425),5,(0,255,0),-1)
                cv2.putText(mask5, "MSE < 50", (57,428), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  
                cv2.circle(mask5,(50,445),5,(0,255,255),-1)
                cv2.putText(mask5, "50 =< MSE < 70", (57,448), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  
                cv2.circle(mask5,(50,465),5,(0,0,255),-1)
                cv2.putText(mask5, "MSE >= 70", (57,468), cv2.FONT_HERSHEY_TRIPLEX, 0.3, (255,255,255), 1)  

                newline = pd.DataFrame([[str(ID),Name_Patient,str(age), Dificulty,str(TotalMSE), str(HitPercentage), str(Seconds),str(DATE),str(Attempt)]], columns=['ID','NAME','AGE','DIFICULTY','MSE','HIT %','TIME (s)','DATE','ATTEMPT'])
                newtable = pd.concat([ExcelFile_Sheet,newline])

                newtable.to_excel('Results.xlsx', index=False)

                cv2.imwrite(FILENAME, mask5) 
            video.release() 
            cv2.destroyAllWindows() 

        elif Exit == 0:
            messagebox.showinfo("Attention", "You have pressed the ESC button and, for doing it, your progress and data will not be stored!")

        else:
            messagebox.showinfo("Attention", "MSEr occured while exiting. Your progress and data will not be stored.")

    if New_Patient==0 and Exit==1:
        tkinter.Tk().geometry("200x200")
        CreateGraphy = messagebox.askyesno(title="Important Information", message="Would you like to create a follow-up graph with previous evaluations?") 

        if CreateGraphy == True:
            file = pd.read_excel("Results.xlsx") 

            x=[]
            y=[]
            colors = []
            size = []
            dificulties = []

            for line in range(len(file.index)):
                if int(ID) in file['ID'].values:
                    if int(ID) == file['ID'][line]:
                        x.append(str(file['DATE'][line]))
                        y.append(file['HIT %'][line])
                        if int(file['MSE'][line])<250:
                            size = 1.3 * int(file['MSE'][line])
                            size.append(size)
                        elif (int(file['MSE'][line])>=250) and (int(file['MSE'][line])<500):
                            size = 1.3 * int(file['MSE'][line])
                            size.append(size)
                        else:
                            size = 1.2 * int(file['MSE'][line])
                            size.append(size)
                        
                        if file['DIFICULTY'][line] == 'Easy':
                            colors.append('green')
                            dificulties.append('Easy')
                        elif file['DIFICULTY'][line] == 'Medium':
                            colors.append('yellow')
                            dificulties.append('Medium')
                        else:
                            colors.append('red')
                            dificulties.append('Hard')


            plt.style.use('seaborn-v0_8-dark-palette')
            
            fig, ax = plt.subplots()
            scatter = ax.scatter(x,y,marker="o",s=size,edgecolors="black",c=colors)
            ax.grid

            plt.xlabel("DATE")
            plt.ylabel("HIT %")
            plt.ylim(0,100)
            plt.title(f"Follow-up graph for patient {str(ID)} - Mr(s) {str(Name_Patient)} ({str(age)} years old)")
            plt.show()