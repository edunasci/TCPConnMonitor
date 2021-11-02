#!/usr/bin/env python3
"""
TCPConnMonitor.py - MonitorTCPConnection

# Prerequisites:
pip3 install pywin32 pyinstaller

# Build:
pyinstaller.exe TCPConnMonitor.py

## Copy the distribution folder to the other servers, install the service, start the service

# Install with autostart:
TCPConnMonitor.exe --startup delayed install
# Install:
TCPConnMonitor.exe install
# Start:
TCPConnMonitor.exe start
# Debug:
TCPConnMonitor.exe debug
# Stop:
TCPConnMonitor.exe stop
# Uninstall:
TCPConnMonitor.exe remove

before using create the eventid log
eventcreate /id 101 /t WARNING /l application /SO TCPConnMonitor /d "TCP Connection Monitor"

references:
https://metallapan.se/post/windows-service-pywin32-pyinstaller/
https://rosettacode.org/wiki/Write_to_Windows_event_log

"""
__author__      = 'Eduardo Marsola do Nascimento'
__copyright__   = 'Copyright 2021-11-02'
__credits__     = ''
__license__     = 'MIT'
__version__     = '0.01'
__maintainer__  = ''
__email__       = ''
__status__      = 'Production'

import psutil
import time

import win32serviceutil
import win32service
import servicemanager

import win32evtlog
import win32evtlogutil

class TCPConnMonitorService:
    def stop(self):
        self.running = False

    def run(self):
        MonitorPorts = [22,23,3389,5900]
        appName   = 'TCPConnMonitor'
        eventID   = 101
        activeTCPConn=[]
        activeTCPConnHistory=[]
        self.running = True
        while self.running:
            # avoid exception if the process had terminated before collection username 
            try: 
                TCPConns = psutil.net_connections(kind="tcp")
                for TCPConn in TCPConns:
                    if TCPConn[5] == 'ESTABLISHED':
                        if TCPConn[4][1] in MonitorPorts:
                            line = [TCPConn[4][0], TCPConn[4][1], psutil.Process(TCPConn[6]).username()]
                            if line not in activeTCPConn:
                                activeTCPConn.append(line)                        
                activeTCPConn.sort()
                # include new connection in the history 
                for TCPConnHist in activeTCPConnHistory.copy():
                    if TCPConnHist[0:3] not in activeTCPConn:
                        activeTCPConnHistory.remove(TCPConnHist)
                        line = TCPConnHist
                        descr = [f'RemoteAddress={line[0]}, RemotePort={line[1]}, UserName={line[2]}, ConnectTime={line[3]}, DisconnectTime={time.strftime("%Y-%m-%dT%H:%M",time.localtime())}']
                        win32evtlogutil.ReportEvent(appName, eventID, eventCategory=0, eventType=win32evtlog.EVENTLOG_WARNING_TYPE, strings=descr, data=None, sid=None)

                    else:
                        activeTCPConn.remove(TCPConnHist[0:3])
                # activeTCPConn has only new connections
                for TCPConn in activeTCPConn:
                    line = [TCPConn[0],TCPConn[1],TCPConn[2],time.strftime("%Y-%m-%dT%H:%M",time.localtime())]
                    descr=[f'RemoteAddress={line[0]}, RemotePort={line[1]}, UserName={line[2]}, ConnectTime={line[3]}']
                    win32evtlogutil.ReportEvent(appName, eventID, eventCategory=0, eventType=win32evtlog.EVENTLOG_WARNING_TYPE, strings=descr, data=None, sid=None)
                    activeTCPConnHistory.append( line )
            except Exception as e: 
                descr=[f'Exception: {e}']
                win32evtlogutil.ReportEvent(appName, eventID, eventCategory=0, eventType=win32evtlog.EVENTLOG_WARNING_TYPE, strings=descr, data=None, sid=None)
            time.sleep(15)

class TCPConnMonitorServiceFramework(win32serviceutil.ServiceFramework):

    _svc_name_ = 'TCPConnMonitorService'
    _svc_display_name_ = 'TCP Connection Monitor Service'
    _svc_description_ = 'Monitor and log TCP connections to specific ports'

    def SvcStop(self):
        """Stop the service"""
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.service_impl.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def SvcDoRun(self):
        """Start the service; does not return until stopped"""
        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
        self.service_impl = TCPConnMonitorService()
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        # Run the service
        self.service_impl.run()

def init():
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(TCPConnMonitorServiceFramework)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(TCPConnMonitorServiceFramework)

if __name__ == '__main__':
    init()
