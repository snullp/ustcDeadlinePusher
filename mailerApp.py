#!/usr/local/bin/python3
#coding:utf-8

#you can set your personal information here to avoid input them manually
username = ""
password = ""

smtpserver = "202.38.64.8"

ntTimeString = "%a, %d %b %Y %H:%M:%S %z"
unixTimeString = "%a, %d %b %Y %H:%M:%S %z (%Z)"

from tkinter import *
from tkinter.scrolledtext import *
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
from datetime import datetime
import os
import os.path
from smtplib import *
import mimetypes
from email import encoders
from email.utils import localtime
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class Application(Frame):
    attachList = []
    def __init__(self, master=None):
        Frame.__init__(self, master)
        
        self.vStatus = StringVar()
        self.vDatetime = StringVar()
        self.vFrom = StringVar()
        self.vTo = StringVar()
        self.vPasswd = StringVar()
        self.vSubject = StringVar()
        self.vAttachInfo = StringVar()
        
        self.setgeometry()
        self.createWidgets()

        self.updateDatetime()
        self.vFrom.set(username)
        self.vPasswd.set(password)

        self.setStatus("Ready.")

    def genMessage(self):
        msg = MIMEMultipart()
        msg['Date'] = self.vDatetime.get()
        msg['From'] = self.vFrom.get()
        msg['To'] = self.vTo.get()
        msg['Subject'] = self.vSubject.get()
        msg.attach(MIMEText(self.wtText.get(1.0,END)))
        for file in self.attachList:
            ctype, encoding = mimetypes.guess_type(file)
            if ctype is None or encoding is not None:
                # No guess could be made, or the file is encoded (compressed), so
                # use a generic bag-of-bits type.
                ctype = 'application/octet-stream'
            maintype, subtype = ctype.split('/', 1)
            if maintype == 'text':
                fp = open(file,'rb')
                # skip annoying charset issue
                #attach = MIMEText(fp.read(), _subtype=subtype)
                attach = MIMEBase(maintype, subtype)
                attach.set_payload(fp.read())
                fp.close()
                encoders.encode_base64(attach)
            elif maintype == 'image':
                fp = open(file, 'rb')
                attach = MIMEImage(fp.read(), _subtype=subtype)
                fp.close()
            elif maintype == 'audio':
                fp = open(file, 'rb')
                attach = MIMEAudio(fp.read(), _subtype=subtype)
                fp.close()
            elif maintype == 'application':
                fp = open(file, 'rb')
                attach = MIMEApplication(fp.read(), _subtype=subtype)
                fp.close()
            else:
                fp = open(file, 'rb')
                attach = MIMEBase(maintype, subtype)
                attach.set_payload(fp.read())
                fp.close()
                # Encode the payload using Base64
                encoders.encode_base64(attach)
            # Set the filename parameter
            attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
            msg.attach(attach)
        return msg

    def send(self):
        if not self.checkDatetime(None): return
        msg = self.genMessage()
        try:
            sender = SMTP(smtpserver)
            sender.login(self.vFrom.get(),self.vPasswd.get())
            sender.send_message(msg,self.vFrom.get(),self.vTo.get().split(','))
            sender.quit()
            self.setStatus("Done.")
        except Exception as e:
            self.setStatus("Error: " + str(e))

    def addAttach(self):
        file = filedialog.askopenfilename()
        if file != "":
            self.attachList.append(file)
            nameList = ""
            for path in self.attachList:
                nameList += os.path.basename(path) + ", "
            self.vAttachInfo.set(nameList)

    def checkDatetime(self,event):
        if os.name == 'nt': timestring = ntTimeString
        else: timestring = unixTimeString
        try:
            timeStr = self.vDatetime.get()
            dt = datetime.strptime(timeStr,timestring)
            if dt.strftime(timestring) != timeStr: raise ValueError
            return True
        except ValueError:
            if messagebox.askyesno("这个时间存在吗？","我感觉这个时间不一定靠谱，你确定它大丈夫？\n点No再来一次吧。",icon=messagebox.ERROR,default=messagebox.NO) == False:
                self.updateDatetime()
                return False
            else: return True
            
    def updateDatetime(self):
        if os.name == 'nt': timestring = ntTimeString
        else: timestring = unixTimeString
        timeStr = localtime().strftime(timestring)
        self.vDatetime.set(timeStr)

    def setgeometry(self):
        top=self.winfo_toplevel()
        top.columnconfigure(0,weight=1)
        top.rowconfigure(0,weight=1)
        self.grid(sticky=N+S+E+W,padx=5,pady=5)
        self.columnconfigure(0,minsize=80)
        self.columnconfigure(1,weight=1)
        self.columnconfigure(2,minsize=80)
        self.columnconfigure(3,weight=1)
        self.rowconfigure(4,weight=1)

    def setStatus(self,s):
        self.vStatus.set(s)

    def createWidgets(self):
        Label(self,text="您的邮箱: ").grid(row=0,column=0)
        Label(self,text="密码: ").grid(row=0,column=2)
        Label(self,text="发送至: ").grid(row=1,column=0)
        Label(self,text="发送时间: ").grid(row=1,column=2)
        Label(self,text="主题: ").grid(row=2,column=0)
        Label(self,textvariable=self.vAttachInfo,anchor=W).grid(row=3,column=1,columnspan=3,sticky=E+W)
        Label(self,textvariable=self.vStatus,anchor=W).grid(row=5,column=1,columnspan=3,sticky=E+W)

        Button(self,text="添加附件",command=self.addAttach).grid(row=3,column=0,sticky=E+W)
        Button(self,text="发送",command=self.send).grid(row=5,column=0,sticky=E+W)
        
        Entry(self,textvariable=self.vFrom).grid(row=0,column=1,sticky=E+W)
        Entry(self,textvariable=self.vPasswd,show="*").grid(row=0,column=3,sticky=E+W)
        Entry(self,textvariable=self.vTo).grid(row=1,column=1,sticky=E+W)
        weDatetime = Entry(self,textvariable=self.vDatetime)
        weDatetime.bind("<FocusOut>",self.checkDatetime)
        weDatetime.grid(row=1,column=3,sticky=E+W)
        Entry(self,textvariable=self.vSubject).grid(row=2,column=1,columnspan=3,sticky=E+W)

        self.wtText = ScrolledText(self)
        self.wtText.grid(row=4,column=0,columnspan=4,sticky=N+S+E+W,pady=5)

app = Application()
app.master.title("USTC Deadline Pusher")
app.mainloop()
