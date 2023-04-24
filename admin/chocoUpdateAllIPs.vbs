' Porpose:
' Remotely upgrade all chocolatey pachages on multiple hosts. A new command prompt will spawn
' for each ip address, run the choco upgrade command, and then exit when done.
' 
' Usage:
' If you want to automate the process, set a value for PASSWD. Otherwise, you will have to type in the password for
' each window that is spawned.

' Requiremnts:
' Chocolatey - https://chocolatey.org/install
' Sysinternals (psexec) - After Chocolatey is installed run: choco install sysinternals

Option Explicit

Dim passwd
Dim ipAddr
Dim IPs
Dim cmd

passwd = "Password"

Set CHIPs = CreateObject("System.Collections.ArrayList")
IPs.Add "192.168.1.2"
IPs.Add "192.168.1.3"
IPs.Add "192.168.1.4"
IPs.Add "192.168.1.5"
IPs.Add "192.168.1.6"


For Each ipAddr in CHIPs
	Call Upgrade("domain\administrator",ipAddr,passwd)
Next

Sub Upgrade(username,ipaddr,passwd)
	dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")
	cmd = "runas /user:" & username & " ""psexec \\"  & ipAddr & " choco upgrade all -y""" 
	'wscript.echo cmd
	'wscript.echo passwd
	objshell.SendKeys cmd
	objshell.SendKeys("{ENTER}") 
	objShell.SendKeys passwd
	objshell.SendKeys("{ENTER}") 
End Sub