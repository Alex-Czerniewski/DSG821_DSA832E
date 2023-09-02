#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Aug 10 17:04:17 2023

@author: pi
"""

import math
import socket
import sys
import time
import threading
from ui import *
from PyQt5 import QtCore as qtc
from PyQt5 import QtWidgets  as qtw
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, 
QLineEdit, QInputDialog)
import xlsxwriter
import datetime
import os
from xlsxwriter.utility import xl_rowcol_to_cell
#'SOUR:POW:PEAK:AMP?'
from ExcelUtil import ExcelUtil

class Main(qtw.QMainWindow):    
    def __init__(self):        
        super(Main,self).__init__()
        self.ui=Ui_MainWindow()
        self.ui.setupUi(self)    
        print ("show GUI")
        self.ui.btn_connect.clicked.connect(self.threaded_connec)
        self.ui.btn_connectSA.clicked.connect(self.connectSA)
        self.ui.btn_sendfrq.clicked.connect(self.threaded_freq)
        self.ui.btn_sendfrq.clicked.connect(self.syncSAToSGfixed)
        self.ui.btn_sweep.clicked.connect(self.syncSAToSGsweep)
        self.ui.btn_sendpwr.clicked.connect(self.threaded_pwr)
        self.ui.btn_sweep.clicked.connect(self.threaded_sweep)
        self.s2 = socket
        self.f="100"
        self.ui.ledit_minf.setText("1000")
        self.ui.ledit_maxf.setText("2000")
        self.ui.ledit_stepS.setText("100")
        self.ui.ledit_dwellT.setText("3")        
        self.ui.ledit_StartPower.setText("1")
        # init excel utility
        self.excelUtil=ExcelUtil()
        # open file for writting
        directory=self.ui.ledit_Directory.text()
        filename=self.ui.ledit_FileName.text()
        
        self.fileName=directory+filename+".xlsx"
        print( self.fileName)
        self.excelUtil.createFile(self.fileName)
        self.writeLabels();
        
     
    def writeLabels(self):
          self.excelUtil.writeData(1,0,"Freq MHz"  )
          self.excelUtil.writeData(2,0,"Power dBm"  )
        
            
    def connectSA(self):
        try:
            self.s2 = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.s2.settimeout(10)
            print ("Socket successfully created")
        except socket.error as err:
            print ("socket creation failed with error %s" %(err)) 
        
        HOST = self.ui.ledit_IPSA.text()
        PORT = self.ui.ledit_portSA.text()
        
        PORT = int(PORT)
        HOST = f'{HOST}'
        print(PORT,HOST)
        self.s2.connect((HOST, PORT))
        
        self.s2.sendall('*IDN?\r\n'.encode())
        print("idn? ",self.s2.recv(4096))
        
    # This Function Will Allow The User To Enter The Parameters For The Sweep
    def Sweep(self):  
        
        #preset Spec Ananlyzer
        self.s.sendall(':SOUR:SWE:MODE AUTO\n'.encode())
        self.s.sendall(':OUTP ON\n'.encode())
        self.s.sendall(':SOUR:SWE:EXEC\n'.encode())  

        rowF=1;RowP=2;col=1               
        minf = float(self.ui.ledit_minf.text())
        maxf = float(self.ui.ledit_maxf.text())
        stepsize = float(self.ui.ledit_stepS.text())
        SETPOWER = self.ui.ledit_StartPower.text()
        t = self.ui.ledit_dwellT.text()
        
        SETPOWER = int(SETPOWER)
        cmd=":SENSe:FREQuency:SPAN 100000000" +"\r\n"
        self.s2.send(cmd.encode())                           
        time.sleep(1)


        
        if SETPOWER <= 5 and SETPOWER >= -100:
            setpower1=":LEV "+str(SETPOWER)+"dBm"+"\r\n"
            self.s.send(setpower1.encode())
        else:
            print('Please enter power less than 10dBm and greater than -100 dBm')
            
        t = int(t)
        
        if minf >= 0.009 and maxf <= 2100:
            self.f=minf
            while self.f<=maxf:
             

                # set Sig Gen
                cmd = ":FREQ "+str(self.f)+"MHz\r\n"
                #print(cmd)
                self.s.sendall(cmd.encode())
                time.sleep(2)
                        
                # set Spec An
                self.centerFrequency = ':SENSe:FREQuency:CENTer '+str(self.f) +"MHz\r\n"
                self.s2.send(self.centerFrequency.encode())
                time.sleep(3)
                
                #set Spec An peak search
                cmd="CALC:MARK1:Y?"+"\r\n"
                self.s2.sendall(cmd.encode())
                time.sleep(2)        
        
                rv=self.s2.recv(4096)
                rv=rv.decode()
                try:
                    self.Rv=float(rv)
                    self.Rv=round(self.Rv,1)
                    
                except valueError:
                    self.Rv=99.9
                    
                time.sleep(5)
           
                if rv==(""):
                    rv="99";
                print("Freq Pwr =",self.Rv,'\t',self.f)
                 
                # write to excel file
                rowF=1                 
                self.excelUtil.writeData(rowF,col,self.f)
                rowP=2                
                self.excelUtil.writeData(rowP,col,self.Rv)
                col=col+1
                           
                self.f += stepsize
                
            print('Sweep Has Been Completed')
                    

            self.excelUtil.createExcelChart()
            self.excelUtil.closeFile()
            
        else:
            print('Please enter a frequency less than 2.1 GHz and greater than 0.009 MHz')
   
    def syncSAToSGfixed(self):
        self.pP1=":CALCulate:MARKer1:CPEak:STATe ON"
        self.pP1=self.pP1+"\r\n"
        self.freq = str(self.freq)
        self.centerFrequencyf = ':SENSe:FREQuency:CENTer '+self.freq +"MHz\r\n"
        self.s2.send(self.pP1.encode())
        self.s2.send(self.centerFrequencyf.encode())
        
    def syncSAToSGsweep(self):     
        self.pP2=":CALCulate:MARKer1:CPEak:STATe ON"
        self.pP2=self.pP2+"\r\n"
        self.centerFrequency = ':SENSe:FREQuency:CENTer '+str(self.f) +"MHz\r\n"
        self.s2.send(self.pP2.encode())
        self.s2.send(self.centerFrequency.encode())
        
        print("sent ctr freq ")
    
    # This Function Will Make a Socket Server and, Connect to the Instrument Using The IP and The Port
    def makeConnec(self):
        try:
            self.s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.s.settimeout(10)
            print ("Socket successfully created")
        except socket.error as err:
            print ("socket creation failed with error %s" %(err))
        
        host = self.ui.ledit_IP.text()
        port = self.ui.ledit_port.text()
        
        Port = int(port)
        Host = f'{host}'
        
        self.s.connect((Host, Port))
    
    # This Function Will Allow The User To Set The Frequency Of The Signal Unless, The Frequency Is Greater Than 2.1 GHz Or Less Than 9 kHz
    def sigFreq(self):
        self.freq = self.ui.ledit_freq.text()
        self.freq=int(float(self.freq))
        if self.freq <= 2100 and self.freq >= 0.009:
            setFreq=":FREQ "+str(self.freq)+"MHz\r\n"
            self.s.send(setFreq.encode())
            time.sleep(2)
        else:
            print('Please enter a frequency less than 2100 MHz and greater than 0.009 MHz')
    
    # This Function Will Allow The User To Set The Power Of The Signal Unless, The Power Is Greater Than 5 dBm Or Less Than 100 dBm
    def sigPwr(self):
        power = self.ui.ledit_pwr.text()
        power=int(float(power))
        if power <= 5 and power >= -100:
            setPower=":LEV "+str(power)+"dBm"+"\r\n"
            self.s.send(setPower.encode())
            time.sleep(1)
        else:         
            print('Please enter power less than 10dBm and greater than -100 dBm')

        
    # Functions That Start With 'threaded' Is Just To Add Threading To The GUI, This Way It Won't Freeze Will Doing An Action
    def threaded_connec(self):
        t = threading.Thread(target=self.makeConnec)
        t.start()
    
    # def threaded_connecSA(self):
    #     t = threading.Thread(target=self.connectSA)
    #     t.start()    
    
    def threaded_freq(self):
        t = threading.Thread(target=self.sigFreq)
        t.start()
        
    def threaded_pwr(self):
        t = threading.Thread(target=self.sigPwr)
        t.start()
    
    def threaded_sweep(self):
        t = threading.Thread(target=self.Sweep)
        t.start()

if __name__=='__main__':
    
    app=qtw.QApplication([])
    
    widget=Main()
    widget.show()
    
    
    app.exec_()
