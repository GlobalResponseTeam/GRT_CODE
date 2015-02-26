'  ----------------------------------------------------------------------------------------------------
'  PrgName:		System Performance.vbs
'  Section:		System Performance
'  Purpose:		Detect System's Performance.
'  Versions:	V1
'  Last Maint:	7/23/2014	By:	Leo Yan
'  Prereq: 		User must have administrator permission for the servers
'  ----------------------------------------------------------------------------------------------------
'  Notes:
'  Add CPU, Memory, Disk info by Jerry
'  ----------------------------------------------------------------------------------------------------
' ----------------------------------- 
' Declare the variables 
' -----------------------------------
On Error Resume Next 
Dim fso,Server,hh,objfile
Dim Strcomputer,connObj
Dim objnetwork,objping,objstatus,strping
Dim objproc,strcpu,objwmi,totalcpu,colproc ' jerry add
Dim colos,objos,strram,totalram 'jerry add
Dim objwmiservice,colitems,objitem,strhd, totalhd 'jerry add
Dim colservices,objservice
Dim NameSpace,emailContent,Email
Dim ClickCancel

Set objNetwork=CreateObject("Wscript.NetWork") 
Setlocale "en-us"
ClickCancel=0
' ----------------------------------- 
'  Define Email Content
' -----------------------------------
emailtitle = "<h1 style=""font: bold 16px Verdana, Arial, Helvetica, sans-serif;"">GRT - Server Status Report</h1>"_
				& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "Time: " & now & "</h3>"_ 	
				& "</TR>"				

emaillabel_1 = "<table width=85% cellspacing=0 cellpadding=0 border=0>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Server</th>"_				
				& "<th style = ""font: 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;""><B>UPTIME</B> [Day:Hour:Min]</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Services</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Status</th>" _
				& "</TR>"

emaillabel_2 =  "<table width=85% cellspacing=0 cellpadding=0 border=0>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Server</th>"_
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">CPU(Load)</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Memory</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Sessions</th>" _
				& "</TR>"				
emaillabel_3 =  "<table width=85% cellspacing=0 cellpadding=0 border=0>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Server</th>"_
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">CPU(Load)</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Memory</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">HDD(Usage)</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Services</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Status</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Sessions</th>" _
				& "</TR>"
emailtail = "</table><h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">IT Global Response Team</h3>"_
			& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">Jabil - Confidential</h3>"_	
			

Set fso = CreateObject("Scripting.FileSystemObject")
' ----------------------------------- 
' Input Site Server List
' -----------------------------------
msg_serverlist = "GRT - System Performance Report" & chr(10) & chr(10) & "Please input the server list file name:" & chr(10) & "Example: [ servers.txt ]"
objfile = inputbox(msg_serverlist,"GRT - Input Server List File Name")

selectResult_serverlist = ClickCancel_F(objfile)

' ------------------------------------------ 
' Judge the Server List Input Box Value 
' ------------------------------------------
Do While selectResult_serverlist = "1"
	
	' --------------------------------------------------------------------------- 
	'  Judge the file name have the suffix (.txt),if not,we add the suffix to it
	'  UCase can convert lower-case to Capital
	' ---------------------------------------------------------------------------	
	
	If InStr(UCase(objfile),"TXT")=0 Then  
		objfile = objfile & ".TXT"
	End If
	
	' ---------------------------------------------------------- 
	'  Judge the file whether exist in current directory or not
	' ----------------------------------------------------------	
	If fso.FileExists(objfile) =0 Then
		msgbox "The file isn't exist! Please double check!",48,"Warning"
		objfile = inputbox(msg_serverlist,"Input Server List File Name")	
		selectResult_serverlist = ClickCancel_F(objfile)
	Else
		selectResult_serverlist = "2"
	End If 
Loop

If selectResult_serverlist = "2" Then
	' ----------------------------------- 
	' Select Report Format
	' -----------------------------------
	msg_Report = "Please select report type:" & chr(10) & chr(10) & "1. Uptime + Service Status" & chr(10) & "2. CPU + Memory + Terminal Sessions" & chr(10) & "3. CPU + Memory + Disk + Service Status" & chr(10) & chr(10) & "Please Input:"

	objselection = inputbox(msg_Report,"GRT - Select Report Type")

	selectResult_Report = ClickCancel_F(objselection)

	Do While selectResult_Report = "1"
			' ------------------------------------- 
			'  Judge the Report Number is correct
			' -------------------------------------
		If objselection = "1" Then
			selectResult_Report = "2"
		ElseIf objselection = "2" Then
			selectResult_Report = "2"
		ElseIf objselection = "3" Then
			selectResult_Report = "2"
		Else 
			msgbox "You Select Wrong No.Please Input Your Report No.(1-3)",48,"Warning"
			objselection = inputbox(msg_Report,"System Performance Report")	
			selectResult_Report = ClickCancel_F(objselection)
		End If 
	Loop
End If

' ------------------------------------------ 
' Judge the Report Format Input Box Value 
' ------------------------------------------
If selectResult_Report = "2" Then
	ExeuteJob_F(objselection)
End If

' ------------------------------------------ 
' Judge the Input BOX Click Cancel Or Blank
' ------------------------------------------
Function ClickCancel_F(objselection)
	' ------------------------------------- 
	' If user click cancel,exit the script
	' -------------------------------------
	if objselection = False Then 
		ClickCancel_F = "0"
		GotoEnd
		Exit Function
	End If 
	
	' --------------------------------------------------- 
	' If the selection is blank,define the value equal 1
	' ---------------------------------------------------	
	ClickCancel_F = "1"

End Function

' ----------------------------------- 
' Get System Status Report
' -----------------------------------
Function ExeuteJob_F(objselection)
	set ws = createobject("wscript.shell")
	ws.popup "Please Wait,The Script is Running....." & vbcrlf & "This Window Will Be Closed In 10 Seconds",10,"Notice",64
	Set Server= fso.OpenTextFile(objfile, 1 , TRUE)
	'Set fso = CreateObject("Scripting.FileSystemObject")
	'Set Server= fso.OpenTextFile(".\servers.txt", 1 , TRUE)
	if objselection = 1 Then
		EmailContent1
	Elseif objselection = 2 Then
		EmailContent2
	Elseif objselection = 3 Then
		EmailContent3
	End if
	emailContent = emailContent & emailtail	
	Server.Close

	' ----------------------------------- 
	' Define email parameters
	' ----------------------------------- 		
	NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
	Set Email = CreateObject("CDO.Message")
	Email.From = "ITGlobalResponseTeam@jabil.com"  
	Email.To = "jerry_hou@jabil.com"	
	'Email.To = "_f7736@jabil.com"
	Email.Subject = "GRT - Server Status Report"
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
End Function


' ----------------------------------- 
' Get CPU Usage
' -----------------------------------
Function Check_CPU(Strcomputer)
	Set objProc = GetObject("winmgmts:\\" & strcomputer & "\root\cimv2:win32_processor='cpu0'")  
	strcpu = objProc.LoadPercentage
	totalcpu = Round(objProc.CurrentClockSpeed/1024) 'jerry add
	If strcpu >= 85 then
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>" & totalcpu & "G (" & strcpu & "%)"
	else
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & totalcpu & "G (" & strcpu & "%)"
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
		totalram = Round(objOS.TotalVisibleMemorySize/1024/1024) 
		if strram >= 85 then
			emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>" & totalram & "G (" & strram & "%)"
		else 
			emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & totalram & "G (" & strram & "%)"
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
			totalhd = round(objItem.size/1073741824) ' jerry
			if strHD >= 90 then
				emailContent = emailContent & "<font color=red>" & left(objItem.DeviceID,2) & " " & totalhd & "G (" & strHD &"%)</font>" & "<br>"
			else 
				emailContent = emailContent & left(objItem.DeviceID,2) & " " & totalhd & "G (" & strHD &"%)" & "<br>"
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
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #F1F1F1;"">N/A"
	else
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strconnection
	end if
End Function

' -----------------------------------
' Define Report No.1
' -----------------------------------
Function EmailContent1
	emailContent = emailtitle & emaillabel_1 & "<TR>"	
	Do While Server.AtEndOfLine <> True 
		strcomputer= UCase(Server.ReadLine)
		emailContent = emailContent & "<TR>"
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">"  & "<a href=""http://alfrdcniosmst01.corp.jabil.org/ninja_upgrade/index.php/search/lookup?query=" & strcomputer & """>" & strcomputer
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
				emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"" colspan=3>" & "<center>Please double check Server Name, Server or Network connections!</center>"
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

' -----------------------------------
' Define Report No.2
' -----------------------------------
Function EmailContent2
	emailContent = emailtitle & emaillabel_2 & "<TR>"
	Do While Server.AtEndOfLine <> True 
		strcomputer= UCase(Server.ReadLine)
		emailContent = emailContent & "<TR>"
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">"  & "<a href=""http://alfrdcniosmst01.corp.jabil.org/ninja_upgrade/index.php/search/lookup?query=" & strcomputer & """>" & strcomputer
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
				emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"" colspan=3>" & "<center>Please double check Server Name, Server or Network connections!</center>"
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

' -----------------------------------
' Define Report No.3
' -----------------------------------
Function EmailContent3
	emailContent = emailtitle & emaillabel_3 & "<TR>"
	Do While Server.AtEndOfLine <> True 
		strcomputer= UCase(Server.ReadLine)
		emailContent = emailContent & "<TR>"
		emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">"  & "<a href=""http://alfrdcniosmst01.corp.jabil.org/ninja_upgrade/index.php/search/lookup?query="  & strcomputer & """>" & strcomputer
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
				emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"" colspan=6>" & "<center>Please double check Server Name, Server or Network connections!</center>"
	' ----------------------------------- 
	' The server can be access
	' -----------------------------------				
			else
				emailContent = emailContent & Check_CPU(Strcomputer) 
				emailContent = emailContent & Check_Memory(Strcomputer)
				emailContent = emailContent & Check_Disk(Strcomputer)
				emailContent = emailContent & Check_Srv(Strcomputer)
				emailContent = emailContent & Check_Session(Strcomputer)
				emailContent = emailContent & "<TR>"
			end If
		NEXT
	Loop
End Function

' -----------------------------------
' Close & Exit The VBS
' -----------------------------------
Function GotoEnd()
	Exit Function
End Function