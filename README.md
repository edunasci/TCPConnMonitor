# TCPConnMonitor
a simple windows service to monitor outbound TCP connections on ports 22, 23, 3389 and 5900

# Usage:

download TCPConnMonitor.zip\n
unzip to a new directory\n
configure new source on eventviewer: eventcreate /id 101 /t WARNING /l application /SO TCPConnMonitor /d "TCP Connection Monitor"\n
install the new service with the command: TCPConnMonitor.exe --startup delayed install\n
start the service: TCPConnMonitor.exe start\n
