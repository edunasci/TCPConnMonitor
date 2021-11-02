#!/usr/bin/env python3
"""
pyEventLog.py - Log event test
Usage:
  pyEventLog.py
Options:
  no options

before using create the eventid log
eventcreate /id 101 /t WARNING /l application /SO RemoteAcessLog /d "Remote Access Log"

"""
__author__      = 'Eduardo Marsola do Nascimento'
__copyright__   = 'Copyright 2021-10-11'
__credits__     = ''
__license__     = 'MIT'
__version__     = '0.01'
__maintainer__  = ''
__email__       = ''
__status__      = 'Production'

import time

import win32serviceutil
import win32service
import servicemanager

import win32api
import win32con
import win32evtlog
import win32security
import win32evtlogutil
 


def init():
    #ph = win32api.GetCurrentProcess()
    #th = win32security.OpenProcessToken(ph, win32con.TOKEN_READ)
    #my_sid = win32security.GetTokenInformation(th, win32security.TokenUser)[0]
    my_sid = None
    applicationName = 'RemoteAcessLog'
    eventID = 101
    category = 0
    myType = win32evtlog.EVENTLOG_WARNING_TYPE
    descr = ['User XXX Has connected to IP Y.Y.Y.Y on port ZZZ']
    data = None
    
    win32evtlog.RegisterEventSource( None, applicationName)
    running = True
    while running: 
        #print("Hello World")
        win32evtlogutil.ReportEvent(applicationName, eventID, eventCategory=category, 
            eventType=myType, strings=descr, data=data, sid=my_sid)
        #time.sleep(15)
        running = False
  
if __name__ == '__main__':
    init()
