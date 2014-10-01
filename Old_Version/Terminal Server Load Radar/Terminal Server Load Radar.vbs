'  ----------------------------------------------------------------------------------------------------
'  PrgName:		Terminal Server Load Radar.vbs
'  Section:		Terminal Server Load Radar
'  Purpose:		Detect Terminal Server Connection Status.
'  Versions:	V2
'  Last Maint:	05/06/2014	By:	Leo Yan
'  Prereq: 		User must have administrator permission for the servers
'  ----------------------------------------------------------------------------------------------------
'  Notes:  Monitor the Terminal Server Connections and CPU Usage and Memory Usage
'  ----------------------------------------------------------------------------------------------------
'declare the variables 
On Error Resume Next 
Dim fso,Server,hh,objfile
Dim Strcomputer,connObj
Dim objnetwork,objping,objstatus,strping
Dim objproc,strconnection,objwmi
Dim colos,objos,strram
Dim objwmiservice,colitems,objitem,strhd
Dim colservices,objservice
Dim NameSpace,emailContent,Email
Dim ClickCancel
'Get Client Info
Set objNetwork=CreateObject("Wscript.NetWork") 
Setlocale "en-us"

' -----------------------------------
' Define Email Title
' -----------------------------------
emailtitle = "<h1 style=""font: bold 16px Verdana, Arial, Helvetica, sans-serif;"">GRT - Terminal Server Load Radar</h1>"_
				& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "Time: " & now & "</h3>"_
				& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "** Below servers need extra attention, please help check! </h3>"_
				& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">" & "<font color=grey>" & "(CPU > 85%；Memory > 85%；Session = 0) </h3>"_
				& "<table width=85% cellspacing=0 cellpadding=0 border=0>"_
				
' -----------------------------------
' Define Email Warning Part
' -----------------------------------
emailwarning= "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Server</th>"_
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">CPU</th>"_
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Memory</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Session</th>" _
				& "<TR>" _
				
' -----------------------------------
' Define Email Content
' -----------------------------------
emailContent= "<TR>"_
				& "</table><h3 style=""font: bold 11px Verdana, Arial, Helvetica, sans-serif;"">" & "**Summary</h3>"_
				& "<table width=85% cellspacing=0 cellpadding=0 border=0>"_

set ws = createobject("wscript.shell")
ws.popup "Please Wait,The Script is Running....." & vbcrlf & "This Window Will Be Closed In 10 Seconds",10,"Notice",64
				
' -----------------------------------
' Get ALL of Site Server File
' -----------------------------------
Set cmd=CreateObject("Wscript.Shell")
cmd.run "cmd /c dir/b .\server\*.txt >.\log\serverlist.log",0
' -----------------------------------------------
' Hold VBS 1s to waiting for the cmd running
' -----------------------------------------------
set WshShell = WScript.CreateObject("WScript.Shell")   
WScript.Sleep 3000
' -----------------------------------
' Get All of Site Server List
' -----------------------------------
Set fso = CreateObject("Scripting.FileSystemObject")
Set Serverlist= fso.OpenTextFile(".\log\serverlist.log", 1, TRUE)

Do While Serverlist.AtEndOfLine <> True 
	serverfile= fso.getbasename(Serverlist.ReadLine)	
	emailContent = emailContent & "<TR>"
	emailContent = emailContent & "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">"& serverfile &"</th>"_
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">CPU</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Memory</th>" _
				& "<th style = ""font: bold 11px Verdana, Arial, Helvetica, sans-serif; color: #FFFFFF; border: 1px solid #C1DAD7; letter-spacing: 2px; text-transform: uppercase; text-align: left; padding: 6px 6px 6px 12px; background: #4F81BD;rowspan: 2; align: center;"">Session</th>" _

' ------------------------------------------
' Initialize Connection and Average Number
' ------------------------------------------
	intTotalLoad = 0
	intSvrCount = 0
	fltAvgLoad = 1.1
	
	Set Server= fso.OpenTextFile(".\Server\"& Serverfile &".txt", 1, True)
' -----------------------------------
' Check The Server Online OR NOT
' -----------------------------------	
	Do While Server.AtEndOfLine <> True 
		strcomputer= UCase(Server.ReadLine)	
		Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & strcomputer & "'")
		For Each objStatus in objPing
			strping = objStatus.protocoladdress
			if strping = "" then
				strconnection = 0
                strcpu = "N/A"
				strram = "N/A"
			else
' -----------------------------------
' Check The Server Halt OR NOT
' -----------------------------------					
				set objwmiservice=getobject("winmgmts:" & "\\" & strcomputer & "\root\cimv2")
				set colservices=objwmiservice.execquery("select * from win32_service where displayname='Terminal Services' or displayname='Remote Desktop Services'")
				for each objservice in colservices
					if objservice.state<>"Running" then
						strconnection = 0
						strcpu = "N/A"
						strram = "N/A"	
					else
' -----------------------------------
' Get Terminal Service Connections
' -----------------------------------						
											
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
' ----------------------------------- 
' Get The Server CPU Usage
' ----------------------------------- 			
						Set objProc = GetObject("winmgmts:\\" & strcomputer & "\root\cimv2:win32_processor='cpu0'")  
						strcpu = objProc.LoadPercentage & "%"
						'msgbox strcpu
' -----------------------------------
' Get The Server Memory Usage
' ----------------------------------- 			
						set objWMI = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
						set colOS = objWMI.InstancesOf("Win32_OperatingSystem") 
						for each objOS in colOS 
							strram = Round(((objOS.TotalVisibleMemorySize-objOS.FreePhysicalMemory)/objOS.TotalVisibleMemorySize)*100) & "%"
						next
						'msgbox strram
					End if	
				Next
			End If

			connWarning = 0	
			cpuWarning = 0
			ramWarning =0
			errorCount = 0
			
			If strcpu >= "85%" then			
				cpuWarning = 1
			end if		

			If strcpu = "100%" then			
				cpuWarning = 1
			end if			
						
			if strram >= "85%" then
				ramWarning = 2
			end If
			
			if strram = "100%" then
				ramWarning = 2
			end If
			
			If strconnection = 0 then							
				connWarning = 4
			end if
						
			errorCount = connWarning + cpuWarning + ramWarning
		Next
		If errorCount = 0 Then
			emailContent = emailContent & "<TR>"
			emailContent = emailContent & Show_Normal(strcomputer) & Show_Normal(strcpu) & Show_Normal(strram) & Show_Normal(strconnection)
		
		else 
			emailwarning = emailwarning & "<TR>"
			emailwarning = emailwarning & Show_Warning(strcomputer)
			emailContent = emailContent & "<TR>"
			emailContent = emailContent & Show_Warning(strcomputer)
			Select Case errorCount
			  Case 1
				emailwarning = emailwarning & Show_Warning(strcpu) & Show_Normal(strram) & Show_Normal(strconnection)
				emailContent = emailContent & Show_Warning(strcpu) & Show_Normal(strram) & Show_Normal(strconnection)
			  Case 2
				emailwarning = emailwarning & Show_Normal(strcpu) & Show_Warning(strram) & Show_Normal(strconnection)
				emailContent = emailContent & Show_Normal(strcpu) & Show_Warning(strram) & Show_Normal(strconnection)
			  Case 3
				emailwarning = emailwarning & Show_Warning(strcpu) & Show_Warning(strram) & Show_Normal(strconnection)
				emailContent = emailContent & Show_Warning(strcpu) & Show_Warning(strram) & Show_Normal(strconnection)
			  Case 4
				emailwarning = emailwarning & Show_Normal(strcpu) & Show_Normal(strram) & Show_Warning(strconnection)
				emailContent = emailContent & Show_Normal(strcpu) & Show_Normal(strram) & Show_Warning(strconnection)
			  Case 5
				emailwarning = emailwarning & Show_Warning(strcpu) & Show_Normal(strram) & Show_Warning(strconnection)
				emailContent = emailContent & Show_Warning(strcpu) & Show_Normal(strram) & Show_Warning(strconnection)
			  Case 6
				emailwarning = emailwarning & Show_Normal(strcpu) & Show_Warning(strram) & Show_Warning(strconnection)
				emailContent = emailContent & Show_Normal(strcpu) & Show_Warning(strram) & Show_Warning(strconnection)
			  Case 7
				emailwarning = emailwarning & Show_Warning(strcpu) & Show_Warning(strram) & Show_Warning(strconnection)
				emailContent = emailContent & Show_Warning(strcpu) & Show_Warning(strram) & Show_Warning(strconnection)
			  Case Else  					
			End Select	
		End if 

		intTotalLoad = intTotalLoad + strconnection
		intSvrCount = intSvrCount + 1		
	Loop
	fltAvgLoad = intTotalLoad / intSvrCount
	emailContent = emailContent & "<TR>"
	emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #D3D3D3;"">" & "<strong>" & "--  AVERAGE"_
								& "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #D3D3D3;"">"_
								& "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #D3D3D3;"">"
	emailContent = emailContent & "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #D3D3D3;"">" & "<strong>" & round(fltAvgLoad,0)
LOOP


emailContent = emailtitle & emailwarning & emailContent & "</table><h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">IT Global Response Team</h3>"_
														& "<h3 style=""font: bold 10px Verdana, Arial, Helvetica, sans-serif;"">Jabil - Confidential</h3>"_
  
Server.Close
Set fso = nothing
Serverfile.Close
Set fso = nothing
Serverlist.Close
Set fso = nothing


Function Show_Warning(objname)
	Show_Warning = "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px; background: #FFF2CC;"">" & "<font color=red>"  & objname
End Function

Function Show_Normal(objname)
	Show_Normal = "<td style = ""border: 1px solid #C1DAD7; font-size:11px; padding: 6px 6px 6px 12px;"">"  & objname
	
End Function	


' -----------------------------------
' Define Email Parameters
' -----------------------------------
NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
Set Email = CreateObject("CDO.Message")
Email.From = "ITGlobalResponseTeam@jabil.com"    
Email.To = "Leo_Yan@jabil.com"
'Email.To = "_f7736@jabil.com"
Email.Subject = "GRT - Terminal Server Load Radar"
Email.Htmlbody =emailContent
With Email.Configuration.Fields
.Item(NameSpace&"sendusing") = 2
.Item(NameSpace&"smtpserver") = "CORIMC04" 
.Item(NameSpace&"smtpserverport") = 25
.Item(NameSpace&"smtpauthenticate") = 1
.update
End With

' -----------------------------------
' Send Out The Email Report
' -----------------------------------
Email.Send
ws.popup "Completed! Please Check E-Mail!",10,"Notice",64