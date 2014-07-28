'  ----------------------------------------------------------------------------------------------------
'  PrgName:		Server Performance.vbs
'  Section:		Server Performance
'  Purpose:		Detect Server's Performance.
'  Versions:	V1
'  Last Maint:	12/04/2013	By:	Leo Yan
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
emailContent = "<h1 style=""font: bold 16px Verdana, Arial, Helvetica, sans-serif;"">GRT - System Performance Report</h1>"_
				& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "Time: " & now & "</h3>"_
				& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "** Some servers need extra attention, please help check! </h3>"_
				& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "(CPU > 85%；Memory > 85%；Disk > 85%；Service Stopped) </h3>"_		 				
				& "<table width=85% cellspacing=0 cellpadding=0 border=0>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Server</th>"_
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">CPU</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Memory</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Disk</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Services</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Status</th>" _
				& "</TR>"
'if there are records then loop through the fields 
'Set fso = CreateObject("Scripting.FileSystemObject")


'Input Site Server List
'objfile = inputbox("Please input the server list file name")

'selectResult = ClickCancel_F(objfile)
'Do While selectResult = "1"
	
	'Judge the file name have the suffix (.txt),if not,we add the suffix to it
	'UCase can convert lower-case to Capital
	'If InStr(UCase(objfile),"TXT")=0 Then  
		'objfile = objfile & ".TXT"
	'End If
	
	'Judge the file whether exist in current directory or not
	'If fso.FileExists(objfile) =0 Then
		'msgbox "The file isn't exist! Please input the correct file name!",48,"Warning"
		'objfile = inputbox("Please input the server list file")	
		'selectResult = ClickCancel_F(objfile)
	'Else
		'selectResult = "2"
	'End If 
'Loop



'If selectResult = "2" Then
	'ExeuteJob_F(objfile)
'End If

'Function GotoEnd()
	'Exit Function
'End Function

'Define function to judge the file name what user input is valid or not.If user click cancel or input blank file name,it's invalid.  
'Function ClickCancel_F(objFileName)
    'If user click cancel,exit the script
	'If objFileName = False Then 
		'ClickCancel_F = "0"
		'GotoEnd
		'Exit Function
	'End If 
	'If the file name is blank,define the value equal 1
	'ClickCancel_F = "1"
'End Function


'Function ExeuteJob_F(objfile)
		'set ws = createobject("wscript.shell")
		'ws.popup "Please Wait,The Script is Running....." & vbcrlf & "This Window Will Be Closed In 10 Seconds",10,"Notice",64	
		
		'Set Server= fso.OpenTextFile(objfile, 1 , TRUE)
		
		set ws = createobject("wscript.shell")
		ws.popup "Please Wait,The Script is Running....." & vbcrlf & "This Window Will Be Closed In 10 Seconds",10,"Notice",64
		
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set Server= fso.OpenTextFile(".\servers.txt", 1 , TRUE)
		'Set Status= fso.OpenTextFile(".\status.txt", 2 , TRUE) 
		Do While Server.AtEndOfLine <> True 
			strcomputer= UCase(Server.ReadLine)
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
					emailContent = emailContent & "<TR>"
					emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strcomputer
					emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"" colspan=5>" & "Still Booting.Offline or Wrong Server Name!"
' ----------------------------------- 
' The server can be access
' -----------------------------------		
				else  				
' ----------------------------------- 
' Get CPU Usage
' -----------------------------------	
					Set objProc = GetObject("winmgmts:\\" & strcomputer & "\root\cimv2:win32_processor='cpu0'")  
					strcpu = objProc.LoadPercentage
					If strcpu >= 85 then
						emailContent = emailContent & "<TR>"
						emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strcomputer
						emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>" & strcpu & "%"
					else
						emailContent = emailContent & "<TR>"
						emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strcomputer
						emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strcpu & "%"
					end if
' ----------------------------------- 
' Get Memory Usage
' -----------------------------------		
					set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
					set colOS = objWMI.InstancesOf("Win32_OperatingSystem") 
					for each objOS in colOS 
						strram = Round(((objOS.TotalVisibleMemorySize-objOS.FreePhysicalMemory)/objOS.TotalVisibleMemorySize)*100) 
						if strram >= 85 then
							emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>" & strram & "%" & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">"
						else 
							emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">" & strram & "%" & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">"
						end If
					next	
' ----------------------------------- 
' Get Disk Usage
' ----------------------------------- 		
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
' ----------------------------------- 
' Get Service Status
' ----------------------------------- 	
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
							'Status.write vbcrlf
						next
					else
						Check_Service("Terminal")
					end If	
				end If
				'Status.write vbcrlf
			NEXT
		Loop

		emailContent = emailContent & "</table><h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">IT Global Response Team</h3>"_
									& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">Jabil - Confidential</h3>"_

' "----------------------------------------------------------------------------------------------------------------------------------------"   
		Server.Close

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
		'define email parameters
		NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
		Set Email = CreateObject("CDO.Message")
		Email.From = "ITGlobalResponseTeam@jabil.com"    
		Email.To = "_f7736@jabil.com"
		Email.Subject = "GRT - System Performance Report"
		Email.Htmlbody =emailContent
		With Email.Configuration.Fields
		.Item(NameSpace&"sendusing") = 2
		.Item(NameSpace&"smtpserver") = "CORIMC04" 
		.Item(NameSpace&"smtpserverport") = 25
		.Item(NameSpace&"smtpauthenticate") = 1
		.update
		End With

		'Send the email report
		Email.Send
		ws.popup "Completed! Please Check E-Mail!",10,"Notice",64

'End Function 
