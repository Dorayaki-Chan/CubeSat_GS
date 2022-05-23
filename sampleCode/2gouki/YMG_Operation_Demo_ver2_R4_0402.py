import serial
import time
import openpyxl
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText
from tkinter import filedialog
import datetime
from PIL import Image, ImageTk
import threading
from sgp4.api import Satrec, jday
import pyorbital
from pyorbital.orbital import Orbital
import julian
import numpy as np
from tkinter import RAISED, ttk

import YMG_Operation_HKdataLabel


#--------------------------------------------- 千葉工大2号館　北緯　35.68826575760112, 東経　140.02027051128135 ---------------------------------------------------
cit_lat = 35.68826575760112
cit_lon = 140.02027051128135
cit_alt = 0
#-------------------------------------------------------------------- cmd ---------------------------------------------------------------------------------

main_cmd =  [0x59, 0x00, 0x11, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
freqUplink = [0xfe, 0xfe, 0x00, 0x9d, 0x00, 0x00, 0x30, 0x31, 0x35, 0x04, 0xfd ]
freqDownlink = [0xfe, 0xfe, 0x00, 0x9d, 0x00, 0x00,  0x50, 0x37, 0x37, 0x04, 0xfd ]
radioMode_CW =[0xFE, 0xFE, 0x00, 0x7C, 0x01, 0x03, 0x01, 0xFD, 0xFE, 0xFE, 0x00, 0x7C, 0x1A, 0x06, 0x00, 0x00, 0xFD ]
radioMode_FMD = [0xFE, 0xFE, 0x00, 0x7C, 0x01, 0x05, 0x01, 0xFD, 0xFE, 0xFE, 0x00, 0x7C, 0x1A, 0x06, 0x01, 0x01, 0xFD ]

add_freq = [0xfe, 0xfe, 0x00, 0x9d, 0x00, 0x00,  0x50, 0x37, 0x37, 0x04, 0xfd]

#---------------------------------------------------------------- Application --------------------------------------------------------------------

class Application(tk.Frame):
    def __init__(self,master = None):
        super().__init__(master)

        self.master.geometry('1350x750')
        self.master.title('YOMOGI Operation')
        self.ent = []
        self.Addjust_value = 0
        self.s = 0
        self.t = 0

        self.port_read()
        self.create_frame()
        self.create_label()
        self.create_button()
        self.create_combobox()
        self.create_check()
        self.create_textbox()
        self.create_listbox()
        self.create_entry()
        self.create_image()
        self.ChangeDownlinkFrequency()
        self.thread_start()
#------------------------------------------------ read serial port ----------------------------------------------------
    def port_read(self):
        ports = ['COM%s' % (i + 1) for i in range(256)]
        self.result = []
        for port in ports:
            try:
                ser_port = serial.Serial(port)
                ser_port.close()
                if len(self.result)>=0:
                    self.result.append(port)
                else:
                    pass
            except (OSError, serial.SerialException):
                pass
#------------------------------------------------------------------ frame --------------------------------------------------------------------
    def create_frame(self):
        self.sub_frame = tk.Frame(self.master,height=725, width=1350,relief="groove")
        self.sub_frame.place(x=0,y=25)
        self.main_frame = tk.Frame(self.master,height=725, width=1350,relief="groove")
        self.main_frame.place(x=0,y=25)
        self.port_frame = tk.Frame(self.sub_frame,height=120, width=350)
        self.port_frame.place(x=20,y=20)
        self.main_button_frame = tk.Frame(self.main_frame,bg="light gray",height=200, width=400,relief="ridge")
        self.main_button_frame.place(x=375, y=510)
        self.time_frame = tk.Frame(self.main_frame,height=100, width=200,relief="groove")
        self.time_frame.place(x=600, y=30)
        self.freqency_frame = tk.Frame(self.main_frame,height=100, width=200,relief="groove")
        self.freqency_frame.place(x=400, y=30)
        self.freqency_button_frame = tk.Frame(self.main_frame,height=80, width=180)
        self.freqency_button_frame.place(x=200, y=50)
        self.file_frame = tk.Frame(self.sub_frame,height=400, width=800)
        self.file_frame.place(x=0,y=150)
        self.HK_frame = tk.Frame(self.main_frame,bg="gainsboro",height=650, width=500)
        self.HK_frame.place(x=830,y=40)
#----------------------------------------------------------------- Label ---------------------------------------------------------------
    def label_create(self,frame,str,px,py):
        self.label = tk.Label(frame,text=str)
        self.label.place(x=px, y=py)

    def create_label(self):
        self.label = tk.Label(self.main_frame,text=u'Command List', font=("",18))
        self.label.place(x=10, y=470)
        self.label = tk.Label(self.HK_frame,text=u'HK Data', font=("",14))
        self.label.place(x=15, y=10)

        self.label_create(self.port_frame,u'ポート選択',50,5)
        self.label_create(self.port_frame,u'Radio',20,29)
        self.label_create(self.port_frame,u'TNC',20,56)
        self.label_create(self.port_frame,u'Rotator',20,83)
        self.label_create(self.main_button_frame,u'Mode',290,25)
        self.label_create(self.main_frame,u'Received Data',30,160)
        self.label_create(self.time_frame,u'JST',5,0)
        self.label_create(self.file_frame,u'TLE',50,20)

        YMG_Operation_HKdataLabel.hklabel(self.HK_frame)
#-------------------------------------------------------------------- Bottun -------------------------------------------------------------------
    def create_button(self):
        self.Button = tk.Button(self.master,text=u'main', width=4, command=lambda:self.change_app(self.main_frame))
        self.Button.place(x=0, y=0)
        self.Button = tk.Button(self.master,text=u'sub', width=4, command=lambda:self.change_app(self.sub_frame))
        self.Button.place(x=40, y=0)
        self.Button = tk.Button(self.port_frame,text=u'connect', width=10, command=lambda:self.rd_connected())
        self.Button.place(x=140, y=25)
        self.Button = tk.Button(self.port_frame,text=u'connect', width=10, command=lambda:self.tnc_connected())
        self.Button.place(x=140, y=52)
        self.Button = tk.Button(self.port_frame,text=u'connect', width=10, command=lambda:self.rotater_connected())
        self.Button.place(x=140, y=79)
        self.Button = tk.Button(self.main_button_frame,text=u'Transmit', width=15,height=3, command=lambda:self.checkMode())
        self.Button.place(x=250, y=110)
        self.Button = tk.Button(self.main_frame,text=u'Save & Clear', width=15, command=lambda:self.save_ReceiveData())
        self.Button.place(x=650, y=150)
        self.Button = tk.Button(self.main_button_frame,text=u'CW', width=3, command=lambda:self.CWcmd() )
        self.Button.place(x=280, y=50)
        self.Button = tk.Button(self.main_button_frame,text=u'FM', width=3, command=lambda:self.FMcmd())
        self.Button.place(x=310, y=50)
        self.Button = tk.Button(self.main_button_frame,text=u'Setting', width=8,command=lambda:self.command_create())
        self.Button.place(x=20, y=20)
        self.Button = tk.Button(self.main_button_frame,text=u'set', width=8,command=lambda:self.cmd_main_set() )
        self.Button.place(x=20, y=50)
        self.Button = tk.Button(self.main_button_frame,text=u'clear', width=8,command=lambda:self.list_del() )
        self.Button.place(x=20, y=80)
        self.Button = tk.Button(self.main_button_frame,text=u'Save', width=8,command=lambda:self.saveCMD() )
        self.Button.place(x=20, y=110)
        self.Button = tk.Button(self.main_button_frame,text=u'Log', width=8,command=lambda:self.read_CommandLog() )
        self.Button.place(x=20, y=140)
        self.Button = tk.Button(self.file_frame,text=u'...', width=3,command=lambda:self.read_tle() )
        self.Button.place(x=730, y=15)
#freqency button
        self.Button = tk.Button(self.freqency_button_frame,text=u'+500Hz',font=("",8), width=7, command=lambda:self.addjustvalue(500) )
        self.Button.place(x=10, y=40)
        self.Button = tk.Button(self.freqency_button_frame,text=u'-500Hz', width=7,font=("",8), command=lambda:self.addjustvalue(-500) )
        self.Button.place(x=10, y=60)
        self.Button = tk.Button(self.freqency_button_frame,text=u'+1kHz', width=7,font=("",8), command=lambda:self.addjustvalue(1000) )
        self.Button.place(x=60, y=40)
        self.Button = tk.Button(self.freqency_button_frame,text=u'-1Hz', width=7,font=("",8), command=lambda:self.addjustvalue(-1000) )
        self.Button.place(x=60, y=60)
        self.Button = tk.Button(self.freqency_button_frame,text=u'+5kHz', width=7,font=("",8), command=lambda:self.addjustvalue(5000) )
        self.Button.place(x=110, y=40)
        self.Button = tk.Button(self.freqency_button_frame,text=u'-5kHz', width=7,font=("",8), command=lambda:self.addjustvalue(-5000) )
        self.Button.place(x=110, y=60)

#------------------------------------------------------------------ Combobox --------------------------------------------------------------------
    def create_combobox(self):
        self.combobox1 = tk.ttk.Combobox(self.port_frame, height= len(self.result), width=7, state="readonly", values=self.result)
        self.combobox1.place(x=70, y=28)
        self.combobox2 = tk.ttk.Combobox(self.port_frame, height= len(self.result), width=7, state="readonly", values=self.result)
        self.combobox2.place(x=70, y=53)
        self.combobox3 = tk.ttk.Combobox(self.port_frame, height= len(self.result), width=7, state="readonly", values=self.result)
        self.combobox3.place(x=70, y=82)
#-------------------------------------------------------------------- checkbutton --------------------------------------------------------------------
    def create_check(self):
        self.chk_value = tk.BooleanVar(value = False)
        self.chk = tk.Checkbutton(self.main_button_frame, text='Continue Mode',variable = self.chk_value)
        self.chk.place(x=250, y=80)
        self.chk_add_value = tk.BooleanVar(value = False)
        self.chk_add = tk.Checkbutton(self.freqency_button_frame, text='Tracking',variable = self.chk_add_value)
        self.chk_add.place(x=90, y=15)

#-------------------------------------------------------------------- textbox --------------------------------------------------------------------
    def create_textbox(self):
        self.text_main = ScrolledText(self.main_frame, font=("", 15), height=14, width=75, bg="azure")
        self.text_main.place(x=20, y=180)

#-------------------------------------------------------------------- listbox --------------------------------------------------------------------
    def create_listbox(self):
        self.Commandbox = tk.Listbox(self.main_frame,width=30, height=9 ,font=("",16), selectmode="single")
        self.Commandbox.place(x=20, y=510)
#-------------------------------------------------------------------- entry --------------------------------------------------------------------
    def hk_entry(self,frame,word_count,px,py):
        he = tk.Entry(frame,width=word_count)
        he.place(x=px, y=py)
        self.ent.append(he)

    def create_entry(self):
        for i in range(47):
            if i < 22:
                self.hk_entry(self.HK_frame,8,150,50+(i*23))
            elif i >= 22 and i < 36:
                self.hk_entry(self.HK_frame,8,350,50+((i-22)*23))
            else :
                self.hk_entry(self.HK_frame,3,365,50+((i-22)*23))

        self.Add_ent = tk.Entry(self.freqency_button_frame,width=8,font=("",10))
        self.Add_ent.place(x=20, y=10)
        self.tle_entry = tk.Entry(self.file_frame,font=6 ,width=70)
        self.tle_entry.place(x=80,y=20)

#---------------------------------------------------------- image ----------------------------------------------------------
    def create_image(self):
        self.YMG_image = Image.open('YOMOGI_settingfile/ロゴ候補1文字あり2021_11_02.png')
        self.YMG_image = self.YMG_image.resize((120, 120))
        self.YMG_image = ImageTk.PhotoImage(self.YMG_image)
        self.canvas = tk.Canvas(self.main_frame,bg="white", width=120, height=120)
        self.canvas.place(x=30, y=20)
        self.canvas.create_image(0, 0, image=self.YMG_image, anchor=tk.NW)

    def change_app(self,frame):
        frame.tkraise()
#---------------------------------------------------------- time ---------------------------------------------------------
    def get_time(self):
        self.timevas = tk.Canvas(self.time_frame,bg="gainsboro", width=200, height=100)
        self.timevas.place(y=0)
        self.label = tk.Label(self.timevas,bg="gainsboro",text=u'JST:', font=("",18))
        self.label.place(x=10, y=15)
        self.label = tk.Label(self.timevas,bg="gainsboro",text=u'UTC:', font=("",18))
        self.label.place(x=10, y=55)
        while True:
            self.jst_now = datetime.datetime.now()
            self.utc_now = datetime.datetime.utcnow()
            self.jst = "{:02}:{:02}:{:02}".format(self.jst_now.hour, self.jst_now.minute, self.jst_now.second)
            self.utc = "{:02}:{:02}:{:02}".format(self.utc_now.hour, self.utc_now.minute, self.utc_now.second)
            self.timevas.delete("all")
            self.timevas.create_text(120, 30, text=self.jst, font=(None,20))
            self.timevas.create_text(120, 70, text=self.utc, font=(None,20))
            time.sleep(1)
#---------------------------------------------------------- Port ---------------------------------------------------------

    def rd_connected(self):                      #radio port setting
        self.RadioPort = self.combobox1.get()
        self.rd = serial.Serial(self.RadioPort,'9600',timeout=0.5)
        self.label_create(self.port_frame,u'Radio Ready!',230,29)
        self.thread_addCW.start()

    def tnc_connected(self):                      #tnc port setting
        self.TNCPort = self.combobox2.get()
        self.tnc = serial.Serial(self.TNCPort,'9600',timeout=0.5)
        self.label_create(self.port_frame,u'TNC Ready!',230,56)
        self.thread_received.start()

    def rotater_connected(self):                      #rotater port setting
        self.RtPort = self.combobox3.get()
        self.rt = serial.Serial(self.RtPort,'9600',timeout=0.5)
        self.label_create(self.port_frame,u'rotater Ready!',230,83)

#---------------------------------------------------------- frequency ---------------------------------------------------------
    def addjustvalue(self,dx):
        self.Addjust_value = self.Addjust_value + dx
        self.Add_ent.delete(0,tk.END)
        self.Add_ent.insert(tk.END,self.Addjust_value)

    #ユリウス日付変換
    def julian_cal(self,time):
        jd2 = julian.to_jd(time, fmt='jd')
        self.jd = int(jd2)
        self.fr = jd2 - self.jd
        return self.jd,self.fr

    def AdjustFrequency(self):
        self.freqvas = tk.Canvas(self.freqency_frame,bg="gainsboro", width=200, height=100)
        self.freqvas.place(y=0)
        self.deltaT = 0.1
        while True:
            self.correctFrequency = 437375000 + self.Addjust_value
            if self.chk_add_value.get() == False or self.s == 0 or self.t == 0:
                self.freqvas.delete("all")
                self.ef = '{:.6f}'.format(self.correctFrequency*10**(-6))
                self.freqvas.create_text(100, 50, text=self.ef, font=(None,20))
                time.sleep(0.1)
            else:
                self.nowtime = datetime.datetime.now()
                self.next = self.nowtime + datetime.timedelta(milliseconds=self.deltaT*1000)
                self.satellite = Satrec.twoline2rv(self.s, self.t)
#cit
                self.cit_observer = pyorbital.astronomy.observer_position(self.nowtime, cit_lon, cit_lat, cit_alt)
                self.cit_observer_next = pyorbital.astronomy.observer_position(self.next, cit_lon, cit_lat, cit_alt)
                self.Rgs_list = np.array(list(self.cit_observer))
                self.Rgs_list_next = np.array(list(self.cit_observer_next))

#now　ユリウス日付から衛星の現在位置を算出
                self.julian_cal(self.nowtime)
                self.e, self.r, self.v = self.satellite.sgp4(self.jd, self.fr)
                self.rSATg = np.array([self.r,self.v])

#next
                self.julian_cal(self.next)
                self.e, self.r, self.v = self.satellite.sgp4(self.jd, self.fr)
                self.rSATg_next = np.array([self.r,self.v])

# Calculate the distance between the satellite and the ground station
                self.r1 = self.rSATg - self.Rgs_list
                self.r2 = self.rSATg_next - self.Rgs_list_next

#Calculate delta r
                self.rX = self.r1[:,0][0]
                self.rY = self.r1[:,1][0]
                self.rZ = self.r1[:,2][0]
                self.rXnext = self.r2[:,0][0]
                self.rYnext = self.r2[:,1][0]
                self.rZnext = self.r2[:,2][0]

                self.r_now = (self.rX**2 + self.rY**2 + self.rZ**2)**(0.5)
                self.r_next = (self.rXnext ** 2 + self.rYnext ** 2 + self.rZnext ** 2) ** (0.5)
                self.dr = (self.r_next - self.r_now) * 1000 / self.deltaT

#Calculate current frequency
                if self.uplinkFlag == False:
                    self.df = 1 - (self.dr/299792458)
                    self.correctFrequency = self.correctFrequency*self.df
                else :
                    self.df = 1 + (self.dr/299792458)
                    self.correctFrequency = self.correctFrequency*self.df
                self.ef = '{:.6f}'.format(self.correctFrequency*10**(-6))
                self.freqvas.delete("all")
                self.freqvas.create_text(100, 50, text=self.ef, font=(None,20))
                self.STRcorrectFrequency = str(int(self.correctFrequency))
                time.sleep(0.1)

    def add_CW(self):
        while True:
            if self.chk_add_value.get() == False:
                self.rd.write(freqDownlink)
                time.sleep(1)
            else:
                time.sleep(1)
                #print(self.add_freq)
                del add_freq[5:10]
                add_freq.insert(5,int(self.STRcorrectFrequency[:1],16))
                add_freq.insert(5,int(self.STRcorrectFrequency[1:3],16))
                add_freq.insert(5,int(self.STRcorrectFrequency[3:5],16))
                add_freq.insert(5,int(self.STRcorrectFrequency[5:7],16))
                add_freq.insert(5,int(self.STRcorrectFrequency[7:9],16))
                self.rd.write(add_freq)

#---------------------------------------------------------- command ---------------------------------------------------------

    def formCMD(self):
        self.Mis_detail()
        self.make_cmd.delete(0,tk.END)
        for i in range(len(self.main_cmd)):
            self.form_command = hex(main_cmd[i]).lstrip("0x")
            self.make_cmd.insert(tk.END,' ')
            self.make_cmd.insert(tk.END,self.form_command.zfill(2))

    def set_cmd(self):
        if len(self.make_cmd.get()) > 0:
            self.Commandbox.insert(tk.END,self.make_cmd.get())
            self.make_cmd.delete(0,tk.END)
        else:
            pass

    def cmd_main_set(self):
        self.set_main_cmd = []
        self.CMD_index = self.Commandbox.curselection()
        if len(self.CMD_index) > 0:
            self.c = self.Commandbox.get(self.CMD_index).split(' ')
            for i in range(11):
                self.set_main_cmd.append(int("0x" + self.c[i+1],16))
        else:
            tk.messagebox.showerror('Error', 'No command')

    def list_del(self):
        self.CMD_index = self.Commandbox.curselection()
        self.Commandbox.delete(self.CMD_index)

#------------------------------------------------------ send command ---------------------------------------------------
    def ChangeUplinkFrequency(self):
        self.uplinkFlag = True

    def ChangeDownlinkFrequency(self):
        self.uplinkFlag = False

    def checkMode(self):
        if self.chk_value.get() == False:
            self.sendCMD()
        else:
            self.continue_cmd()

    def sendCMD(self):      #send cmd
        self.send_main_cmd = []
        for i in range(len(self.set_main_cmd)):
            if self.set_main_cmd[i] == 0xc0:
                self.send_main_cmd.append(0xDB)
                self.send_main_cmd.append(0xDC)
            elif self.set_main_cmd[i] == 0xDB:
                self.send_main_cmd.append(0xDB)
                self.send_main_cmd.append(0xDD)
            else:
                self.send_main_cmd.append(self.set_main_cmd[i])
        self.send_main_cmd.insert(0,0x42)
        self.send_main_cmd.insert(0,0x00)
        self.send_main_cmd.insert(0,0xc0)
        self.send_main_cmd.append(0xc0)
        self.rd.write(radioMode_FMD)
        time.sleep(0.1)
        self.ChangeUplinkFrequency()
        self.rd.write(freqUplink)
        time.sleep(0.2)
        self.tnc.write(self.send_main_cmd)
        time.sleep(1.5)
        self.ChangeDownlinkFrequency()
        self.rd.write(freqDownlink)
        self.tx = ""
        for i in range(len(self.set_main_cmd)):
            self.tx = self.tx + hex(self.set_main_cmd[i]).lstrip("0x").zfill(2) + " "
        self.text_main.insert(tk.END,"#" + str(datetime.datetime.now()) + " " + "send command: ")
        self.text_main.insert(tk.END,self.tx)
        self.text_main.insert(tk.END,"\n")
        self.text_main.see("end")

    def continue_cmd(self):
        self.Commandbox.selection_set(first=0,last=None)
        self.cmd_main_set()
        self.sendCMD()
        self.Commandbox.delete(self.CMD_index)

    def CWcmd(self):        #CW mode on
        self.rd.write(radioMode_CW)

    def FMcmd(self):        #FM mode on & 437.375.00 setting
        self.rd.write(radioMode_FMD)
        self.rd.write(freqDownlink)

    def save_ReceiveData(self):
        now = datetime.datetime.now()
        filename = "YOMOGI_ReceiveData/YMG data" + now.strftime('%Y%m%d_%H%M%S') + '.txt'
        Pidgeot = open(filename,'w')
        input = self.text_main.get("1.0", "end-1c")
        Pidgeot.write(str(input))
        Pidgeot.close()
        self.text_main.delete("1.0","end")

    def saveCMD(self):
        self.cmd_log = "YOMOGI_command_Log/" + "YOMOGI_commando_Log_" + datetime.datetime.now().strftime('%Y%m%d_%H%M%S') + '.txt'
        self.save = self.Commandbox.get(0, self.Commandbox.size())
        with open(self.cmd_log,'w') as Rattata:
            for i in range(len(self.save)):
                Rattata.write(self.save[i] + "\n")
        messagebox.showinfo('complete', 'save complete')

    def read_CommandLog(self):
        typ = [('テキストファイル','*.txt')]
        dir = "YOMOGI_command_Log/"
        log_fle = filedialog.askopenfilename(filetypes = typ, initialdir = dir)
        with open(log_fle,'r') as self.lf:
            self.cmd_log = self.lf.read().split("\n")
        for i in range(len(self.cmd_log)):
            self.Commandbox.insert(tk.END,self.cmd_log[i])

    def read_tle(self):
        typ = [('テキストファイル',".txt")]
        dir = "YOMOGI_settingfile/"
        tle_fle = filedialog.askopenfilename(filetypes = typ, initialdir = dir)
        self.tle_entry.insert(tk.END,tle_fle)
        with open(tle_fle,'r') as self.lf:
            self.tle_data = self.lf.read().split("\n")
        self.s = self.tle_data[1]
        self.t = self.tle_data[2]

#------------------------------------------------------ read cmd xl file ---------------------------------------------------
    def read_cmd_xl(self,sheet,hex_cel,detail_cel,hex_list,detail_list):
        self.active_sheet = sheet
        for i in range(300):
            hex_value = hex_cel + str(i + 2)
            detail_value = detail_cel + str(i + 2)
            self.hex = self.active_sheet[hex_value].value
            self.detail = self.active_sheet[detail_value].value
            if self.detail == None:
                pass
            else:
                detail_list.append(self.detail)
                hex_list.append(self.hex)

#---------------------------------------------------------- mission command ---------------------------------------------------------

    def win_close(self):
        self.mis_cmd_win.destroy()

    def cmd_entry(self,frame,word_count,px,py,value):
        command_ent = tk.Entry(frame,width=word_count,justify=tk.CENTER,font=20)
        command_ent.place(x=px, y=py)
        command_ent.insert(0, value)
        self.cmd_ent.append(command_ent)

    def detail_cmd(self,cmdlist,cmd,cmd_detail,cmd_commbo):              #obc command setting
        self.cmd_ent[cmdlist].delete(0, tk.END)
        self.cmd_ent[cmdlist].insert(tk.END,cmd[cmd_detail.index(cmd_commbo.get())].replace("0x",""))
        if cmd == self.obc_cmd_list:
            self.frag = 1
        elif cmd == self.mis_cmd_list:
            self.frag = 2
        else:
            pass

    def set_command(self):
        txt = ""
        main_cmd.clear
        for i in range(11):
            main_cmd.append(int("0x" + self.cmd_ent[i].get(),16))
            txt = txt + " " + str(self.cmd_ent[i].get())
        self.Commandbox.insert(tk.END,txt)

    def command_create(self):
        self.All_cmd_xlsx = openpyxl.load_workbook("YOMOGI_settingfile/006_YMGテレコマリスト.xlsx",data_only=True)
        self.cmd_ent = []

        self.mis_cmd_win = tk.Toplevel()
        self.mis_cmd_win.geometry('800x450')
        self.mis_cmd_win.title('Setting Command')

        self.mis_button_frame = tk.Frame(self.mis_cmd_win,  height=90, width=750)
        self.mis_button_frame.place(x=25, y=40)

        for i in range(11):
            if i == 0:
                self.cmd_entry(self.mis_cmd_win, 4, 50+60*i, 340, 59)
            else:
                self.cmd_entry(self.mis_cmd_win, 4, 50+60*i, 340, '{0:02d}'.format(00))

        self.Button = tk.Button(self.mis_cmd_win,text=u' set ', width=6, relief=RAISED,command=lambda:self.set_command())
        self.Button.place(x=580, y=400)
        self.Button = tk.Button(self.mis_cmd_win,text=u' OK ', width=6, relief=RAISED,command=lambda:self.win_close())
        self.Button.place(x=640, y=400)

#------------------------------------------- one~two ------------------------------------------------------

        self.one = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.one.place(x=50, y=150)
        self.label_create(self.one,u' 工事中... ',330,60)
        self.two = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.two.place(x=50, y=150)
        self.label_create(self.two,u' 工事中... ',330,60)

#------------------------------------------- three ------------------------------------------------------

        self.three = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.three.place(x=50, y=150)
        self.label_create(self.three,u' COM command ',20,20)
        self.com_cmd_list = []
        self.com_cmd_detail = []
        self.com_cmd_xlsx = self.All_cmd_xlsx["COM"]
        self.read_cmd_xl(self.com_cmd_xlsx,"c","d",self.com_cmd_list,self.com_cmd_detail)
        self.combobox4 = tk.ttk.Combobox(self.three, height= 70, width=50, state="readonly", values=self.com_cmd_detail)
        self.combobox4.bind('<<ComboboxSelected>>',lambda event:self.detail_cmd(2,self.com_cmd_list,self.com_cmd_detail,self.combobox4))
        self.combobox4.place(x=150, y=80)

#------------------------------------------- four ------------------------------------------------------
        self.four = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.four.place(x=50, y=150)
        self.label_create(self.four,u' Main or Mis command ',20,20)

        self.label_create(self.four,u' OBC ',100,60)
        self.obc_cmd_list = []
        self.obc_cmd_detail = []
        self.obc_cmd_xlsx = self.All_cmd_xlsx["OBC"]
        self.read_cmd_xl(self.obc_cmd_xlsx,"c","d",self.obc_cmd_list,self.obc_cmd_detail)
        self.combobox51 = tk.ttk.Combobox(self.four, height= 70, width=50, state="readonly", values=self.obc_cmd_detail)
        self.combobox51.bind('<<ComboboxSelected>>',lambda event:self.detail_cmd(3,self.obc_cmd_list,self.obc_cmd_detail,self.combobox51))
        self.combobox51.place(x=150, y=60)

        self.label_create(self.four,u' MIS ',100,100)
        self.mis_cmd_list = []
        self.mis_cmd_detail = []
        self.mis_cmd_xlsx = self.All_cmd_xlsx["MIS"]
        self.read_cmd_xl(self.mis_cmd_xlsx,"c","d",self.mis_cmd_list,self.mis_cmd_detail)
        self.combobox52 = tk.ttk.Combobox(self.four, height= 70, width=50, state="readonly", values=self.mis_cmd_detail)
        self.combobox52.bind('<<ComboboxSelected>>',lambda event:self.detail_cmd(3,self.mis_cmd_list,self.mis_cmd_detail,self.combobox52))
        self.combobox52.place(x=150, y=100)

#------------------------------------------- five~ ------------------------------------------------------
        self.five = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.five.place(x=50, y=150)
        self.label_create(self.five,u' 工事中... ',330,60)
        self.six = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.six.place(x=50, y=150)
        self.label_create(self.six,u' 工事中... ',330,60)
        self.seven = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.seven.place(x=50, y=150)
        self.label_create(self.seven,u' 工事中... ',330,60)
        self.eight = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.eight.place(x=50, y=150)
        self.label_create(self.eight,u' 工事中... ',330,60)
        self.nine = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.nine.place(x=50, y=150)
        self.label_create(self.nine,u' 工事中... ',330,60)
        self.ten = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.ten.place(x=50, y=150)
        self.label_create(self.ten,u' 工事中... ',330,60)
        self.eleven = tk.Frame(self.mis_cmd_win, bg="lavender",  height=150, width=700)
        self.eleven.place(x=50, y=150)
        self.label_create(self.eleven,u' 工事中... ',330,60)

        self.Button = tk.Button(self.mis_button_frame,text=u'1', font=18, relief=RAISED, padx=20,pady=15,command=lambda:self.change_app(self.one))
        self.Button.place(x=50, y=20)
        self.Button = tk.Button(self.mis_button_frame,text=u'2', font=18, relief=RAISED, padx=20,pady=15,command=lambda:self.change_app(self.two))
        self.Button.place(x=110, y=20)
        self.Button = tk.Button(self.mis_button_frame,text=u'3', font=18, relief=RAISED, padx=20,pady=15,command=lambda:self.change_app(self.three))
        self.Button.place(x=170, y=20)
        self.Button = tk.Button(self.mis_button_frame,text=u'4', font=18, relief=RAISED, padx=20,pady=15,command=lambda:self.change_app(self.four))
        self.Button.place(x=230, y=20)
        self.Button = tk.Button(self.mis_button_frame,text=u'5', font=18, relief=RAISED, padx=20,pady=15,command=lambda:self.change_app(self.five))
        self.Button.place(x=290, y=20)
        self.Button = tk.Button(self.mis_button_frame,text=u'6', font=18, relief=RAISED, padx=20,pady=15,command=lambda:self.change_app(self.six))
        self.Button.place(x=350, y=20)
        self.Button = tk.Button(self.mis_button_frame,text=u'7', font=18, relief=RAISED, padx=20,pady=15,command=lambda:self.change_app(self.seven))
        self.Button.place(x=410, y=20)
        self.Button = tk.Button(self.mis_button_frame,text=u'8', font=18, relief=RAISED, padx=20,pady=15,command=lambda:self.change_app(self.eight))
        self.Button.place(x=470, y=20)
        self.Button = tk.Button(self.mis_button_frame,text=u'9', font=18, relief=RAISED, padx=20,pady=15,command=lambda:self.change_app(self.nine))
        self.Button.place(x=530, y=20)
        self.Button = tk.Button(self.mis_button_frame,text=u'10', font=18, relief=RAISED, padx=16,pady=15,command=lambda:self.change_app(self.ten))
        self.Button.place(x=590, y=20)
        self.Button = tk.Button(self.mis_button_frame,text=u'11', font=18, relief=RAISED, padx=16,pady=15,command=lambda:self.change_app(self.eleven))
        self.Button.place(x=650, y=20)


#---------------------------------------------------------- receive ---------------------------------------------------------

    def received(self):
        while True:
            self.byte_data_full = []
            count = 0
            while True:
                self.byte_data = self.tnc.readline()

                if 0 < len(self.byte_data):
                    self.byte_data_full += self.byte_data.hex()
                    count = 1

                else:
                    if count == 1:
                        count = 3

                if count == 3:
                    break

            self.data = ""
            for i in range(len(self.byte_data_full)):
                self.data = self.data + self.byte_data_full[i]

            self.text_main.insert(tk.END,self.data)
            self.text_main.insert(tk.END,"\n")
            self.text_main.see("end")

#---------------------------------------------------------- thread start ---------------------------------------------------------
    def thread_start(self):
        self.thread_freqency = threading.Thread(target=self.AdjustFrequency, daemon=True)
        self.thread_addCW = threading.Thread(target=self.add_CW, daemon=True)
        self.timer = threading.Thread(target = self.get_time, daemon=True)
        self.thread_received = threading.Thread(target=self.received, daemon=True)

        self.thread_freqency.start()
        self.timer.start()

def main():
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()

if __name__ == "__main__":
    main()

