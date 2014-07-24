'  ----------------------------------------------------------------------------------------------------
'  PrgName:		System Performance.vbs
'  Section:		System Performance
'  Purpose:		Detect System's Performance.
'  Versions:	V1
'  Last Maint:	7/23/2014	By:	Leo Yan
'  Prereq: 		User must have administrator permission for the servers
'  ----------------------------------------------------------------------------------------------------
'  Notes:
'  ----------------------------------------------------------------------------------------------------
'declare the variables 
On Error Resume Next 
Dim fso,Server,hh,objfile
Dim Strcomputer,connObj
Dim objnetwork,objping,objstatus,strping
Dim objproc,strcpu,objwmi
Dim colos,objos,strram
Dim objwmiservice,colitems,objitem,strhd
Dim colservices,objservice
Dim NameSpace,emailContent,Email
Dim ClickCancel
'Get Client Info
Set objNetwork=CreateObject("Wscript.NetWork") 
Setlocale "en-us"
'ClickCancel=0
'Write email content
emailtitle = "<h1 style=""font: bold 16px Verdana, Arial, Helvetica, sans-serif;"">GRT - Server Status Report</h1>"_
				& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "Time: " & now & "</h3>"_ 	
				& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "** Some servers need extra attention, please help check! </h3>"_
				& "</TR>"				

emaillabel_1 = "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "(Uptime > 4 Hours；Service Stopped) </h3>"_
				& "<table width=85% cellspacing=0 cellpadding=0 border=0>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Server</th>"_				
				& "<th style = ""font: 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;""><B>UPTIME</B> [Day:Hour:Min]</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Services</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Status</th>" _
				& "</TR>"

emaillabel_2 =  "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "(CPU > 85%；Memory > 85%；Session = 0) </h3>"_	
				& "<table width=85% cellspacing=0 cellpadding=0 border=0>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Server</th>"_
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">CPU</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Memory</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Sessions</th>" _
				& "</TR>"				
emaillabel_3 =  "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "(CPU > 85%；Memory > 85%；Disk > 85%；Service Stopped) </h3>"_	
				& "<table width=85% cellspacing=0 cellpadding=0 border=0>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Server</th>"_
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">CPU</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Memory</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Disk</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Services</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Status</th>" _
				& "</TR>"
emailtail = "</table><h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">IT Global Response Team</h3>"_
			& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">Jabil - Confidential</h3>"_	
			
set ws = createobject("wscript.shell")
ws.popup "Please Wait,The Script is Running....." & vbcrlf & "This Window Will Be Closed In 10 Seconds",10,"Notice",64


Set fso = CreateObject("Scripting.FileSystemObject")
Set Server= fso.OpenTextFile(".\servers.txt", 1 , TRUE)

'EmailContent1
EmailContent2
'EmailContent3

emailContent = emailContent & emailtail		  				

' ----------------------------------- 
' Get CPU Usage
' -----------------------------------
Function Check_CPU(Strcomputer)
	Set objProc = GetObject("winmgmts:\\" & strcomputer & "\root\cimv2:win32_processor='cpu0'")  
	strcpu = objProc.LoadPercentage
	If strcpu >= 85 then
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>" & strcpu & "%"
	else
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strcpu & "%"
	end if
End Function						

' ----------------------------------- 
' Get Memory Usage
' -----------------------------------
Function Check_Memory(Strcomputer) 				
	set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
	set colOS = objWMI.InstancesOf("Win32_OperatingSystem") 
	for each objOS in colOS 
		strram = Round(((objOS.TotalVisibleMemorySize-objOS.FreePhysicalMemory)/objOS.TotalVisibleMemorySize)*100) 
		if strram >= 85 then
			emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>" & strram & "%"
		else 
			emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strram & "%"
		end If
	next	
End Function

' ----------------------------------- 
' Get Disk Usage
' ----------------------------------- 
Function Check_Disk(Strcomputer) 						
	emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">"
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery( "SELECT * FROM Win32_LogicalDisk",,48) 
		For Each objItem in colItems 
		Err.Clear
		If objItem.size Then
			strHD =round((100*((objItem.size-objItem.FreeSpace)/objItem.size))) 
			if strHD >= 90 then
				emailContent = emailContent & "<font color=red>" & left(objItem.DeviceID,2) & strHD &"%</font>" & "<br>"
			else 
				emailContent = emailContent & left(objItem.DeviceID,2) & strHD &"%" & "<br>"
			end if
		Else
		End If
	Next	
End Function

' ----------------------------------- 
' Get The Server Uptime
' -----------------------------------
Function Check_Uptime(Strcomputer)  			
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colOses =  objWMIService.ExecQuery("SELECT LastBootUpTime From Win32_OperatingSystem")
	Struptime=""
		For Each objOs In colOses
			diffMin = DateDiff("n", wmiDateStringToDate(objOs.LastBootUpTime), Now)
			diffDays = Fix(diffMin / (60 * 24))
			diffMin = diffMin - diffDays * 24 * 60
			If diffDays >= 1 Then
				Struptime = Struptime & CStr(diffDays) & "D "
			End If
			diffHours = Fix(diffMin / 60)
			diffMin = diffMin - diffHours * 60
			If diffHours >= 1 Then
				Struptime = Struptime & CStr(diffHours) & "H "
			End If
			If diffMin >= 1 Then
				Struptime = Struptime & CStr(diffMin) & "M "
			End If
			'WScript.Echo Struptime
			
			if diffDays >= 1 then
				emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>" & struptime & "<br>"
			elseif diffHours >= 4 then
				emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>" & struptime & "<br>"
			else 
				emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & struptime  & "<br>"
			end If
			
		Next
End Function

' ----------------------------------- 
' Calculate Uptime
' -----------------------------------
Function wmiDateStringToDate(dtmDate)
    wmiDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) & " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate, 13, 2))
End Function							

' ----------------------------------- 
' Get Service Status
' ----------------------------------- 
Function Check_Srv(Strcomputer) 	
	set objwmiservice=getobject("winmgmts:" & "\\" & strcomputer & "\root\cimv2")
	if instr(strcomputer,"PAR")>1 then 
		Check_Service("PAR")
	elseif instr(strcomputer,"CMP")>1 then 
		Check_Service("CMP")
	elseif instr(strcomputer,"SQL")>1 then 
		Check_Service("SQL")
	elseif instr(strcomputer,"LOFT")>1 then 
		Check_Service("LOFT")
	elseif instr(strcomputer,"PLVC")>1 then
		set colservices=objwmiservice.execquery("select * from win32_service where displayname='SCADA for SMT' or displayname='PLVC Data Service' or displayname='MES Data SERVICE for PLVC'")
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" 
		for each objservice in colservices
			emailContent = emailContent & objservice.displayname & "<br>"
		next
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">"
		for each objservice in colservices
			if objservice.state="Running" then							
				emailContent = emailContent & objservice.state & "<br>"
			else 
				emailContent = emailContent & "<font color=red>" & objservice.state & "<br>"
			end if
		next
	else
		Check_Service("Terminal")
	end If
End Function					

Function Check_Service(objname)
	If objname="CMP" then 
		strservice = "DataImporter"
	elseif objname="PAR" then
		strservice = "MesParserNet"
	elseif objname="SQL" then
		strservice = "MSSQLSERVER"
	elseif objname="LOFT" then
		strservice = "WatchDogNT"
	elseif objname="Terminal" then
		strservice = "Termservice"
	end if
    set colservices=objwmiservice.execquery("select * from win32_service where name='"& strservice &"'")
	i = 0
	for each objservice in colservices
		i = i + 1
		if objservice.state="Running" then
			emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & objservice.displayname & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & objservice.state
		else 
			emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & objservice.displayname & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>" & objservice.state
		end if
	next
' ----------------------------------- 
' Can't find the Service
' ----------------------------------- 
	if i=0 then
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & "N/A" & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & "N/A"				
	end if 
End Function

' -----------------------------------
' Get Terminal Service Connections
' -----------------------------------
Function Check_Session(Strcomputer) 											
	' -----------------------------------
	' 		Check The OS Version
	' -----------------------------------
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery( _
				"SELECT * FROM Win32_OperatingSystem",,48)
	For Each objItem in colItems 
		objos = objItem.Caption
	Next
' -----------------------------------
' 		Windows 2003 OS
' -----------------------------------
	if InStr(objos,"2003")<>0 Then
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
		Set colItems = objWMIService.ExecQuery( _
					"SELECT * FROM Win32_PerfFormattedData_TermService_TerminalServices",,48)
' -----------------------------------
' 		Windows 2008 OS
' -----------------------------------										
	elseif InStr(objos,"2008")<>0 Then
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
		Set colItems = objWMIService.ExecQuery( _
					"SELECT * FROM Win32_PerfFormattedData_LocalSessionManager_TerminalServices",,48)
	else
	End if	
	For Each objItem in colItems
		strconnection = objItem.ActiveSessions						
	Next
	If strconnection = 0 then
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>" & strconnection
	else
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strconnection
	end if
End Function

Function EmailContent1
	emailContent = emailtitle & emaillabel_1 & "<TR>"
	Do While Server.AtEndOfLine <> True 
		strcomputer= UCase(Server.ReadLine)
		
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strcomputer
	' ----------------------------------- 
	' Confirm Server Is Available
	' -----------------------------------		
		Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & strcomputer & "'")
		For Each objStatus in objPing
			strping = objStatus.protocoladdress

	' ----------------------------------- 
	' The server can't be access
	' -----------------------------------
			if strping = "" then
				emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"" colspan=3>" & "Still Booting.Offline or Wrong Server Name!"
				emailContent = emailContent & "<TR>"
				
	' ----------------------------------- 
	' The server can be access
	' -----------------------------------				
			else
				emailContent = emailContent & Check_Uptime(Strcomputer)
				emailContent = emailContent & Check_Srv(Strcomputer)
				emailContent = emailContent & "<TR>"
			end If
		NEXT
	Loop
End Function	

Function EmailContent2
	emailContent = emailtitle & emaillabel_2 & "<TR>"
	Do While Server.AtEndOfLine <> True 
		strcomputer= UCase(Server.ReadLine)
		
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strcomputer
	' ----------------------------------- 
	' Confirm Server Is Available
	' -----------------------------------	
			Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & strcomputer & "'")
		For Each objStatus in objPing
			strping = objStatus.protocoladdress

	' ----------------------------------- 
	' The server can't be access
	' -----------------------------------
			if strping = "" then
				emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"" colspan=3>" & "Still Booting.Offline or Wrong Server Name!"
	' ----------------------------------- 
	' The server can be access
	' -----------------------------------				
			else
				emailContent = emailContent & Check_CPU(Strcomputer) 
				emailContent = emailContent & Check_Memory(Strcomputer)
				emailContent = emailContent & Check_Session(Strcomputer)
				emailContent = emailContent & "<TR>"
			end If
		NEXT
	Loop
End Function

Function EmailContent3
	emailContent = emailtitle & emaillabel_3 & "<TR>"
	Do While Server.AtEndOfLine <> True 
		strcomputer= UCase(Server.ReadLine)
		
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strcomputer
	' ----------------------------------- 
	' Confirm Server Is Available
	' -----------------------------------	
			Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & strcomputer & "'")
		For Each objStatus in objPing
			strping = objStatus.protocoladdress

	' ----------------------------------- 
	' The server can't be access
	' -----------------------------------
			if strping = "" then
				emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"" colspan=5>" & "Still Booting.Offline or Wrong Server Name!"
	' ----------------------------------- 
	' The server can be access
	' -----------------------------------				
			else
				emailContent = emailContent & Check_CPU(Strcomputer) 
				emailContent = emailContent & Check_Memory(Strcomputer)
				emailContent = emailContent & Check_Disk(Strcomputer)
				emailContent = emailContent & Check_Srv(Strcomputer)
				emailContent = emailContent & "<TR>"
			end If
		NEXT
	Loop
End Function

Server.Close

' ----------------------------------- 
' Define email parameters
' ----------------------------------- 		
NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
Set Email = CreateObject("CDO.Message")
Email.From = "ITGlobalResponseTeam@jabil.com"    
Email.To = "Leo_Yan@jabil.com"
Email.Subject = "GRT - Server Status after Reboot"
Email.Htmlbody =emailContent
With Email.Configuration.Fields
.Item(NameSpace&"sendusing") = 2
.Item(NameSpace&"smtpserver") = "CORIMC04" 
.Item(NameSpace&"smtpserverport") = 25
.Item(NameSpace&"smtpauthenticate") = 1
.update
End With

' ----------------------------------- 
' Send the email report
' ----------------------------------- 		'
Email.Send		

WS.popup "Completed! Please Check E-Mail!",10,"Notice",64

