'###################################################################################################################################
'## This script was developed by Guberni and is part of Tellki's Monitoring Solution								              ##
'##																													              ##
'## December, 2014																									              ##
'##																													              ##
'## Version 1.0																										              ##
'##																													              ##
'## DESCRIPTION: Monitor memory utilization (physical and virtual)																  ##
'##																													              ##
'## SYNTAX: cscript "//Nologo" "//E:vbscript" "//T:90" "Memory.vbs" <HOST> <METRIC_STATE> <USERNAME> <PASSWORD> <DOMAIN>          ##
'##																													              ##
'## EXAMPLE: cscript "//Nologo" "//E:vbscript" "//T:90" "Memory.vbs" "10.10.10.1" "1,1,1,1,1,1,1,1,1,1,0,0" "user" "pwd" "domain" ##
'##																													              ##
'## README:	<METRIC_STATE> is generated internally by Tellki and its only used by Tellki default monitors. 						  ##
'##         1 - metric is on ; 0 - metric is off					              												  ##
'## 																												              ##
'## 	    <USERNAME>, <PASSWORD> and <DOMAIN> are only required if you want to monitor a remote server. If you want to use this ##
'##			script to monitor the local server where agent is installed, leave this parameters empty ("") but you still need to   ##
'##			pass them to the script.																						      ##
'## 																												              ##
'###################################################################################################################################

'Start Execution
Option Explicit
'Enable error handling
On Error Resume Next
If WScript.Arguments.Count <> 5 Then 
	CALL ShowError(3, 0) 
End If
'Set Culture - en-us
SetLocale(1033)

'METRIC_ID
Const FreePhysMem = "66:Free Physical Memory:4"
Const FreeVirtMem = "57:Free Virtual Memory:4"
Const UsedPhysMemPerc = "82:% Used Physical Memory:6"
Const UsedVirtMemPerc = "68:% Used Virtual Memory:6"
Const PagesSec = "129:Pages/Sec:4"
Const PagesFaultSec = "17:Page Faults/Sec:4"
Const PagesInSec = "149:Pages Input/Sec:4"
Const PagesOutSec = "117:Pages Output/Sec:4"
Const PagesWriteSec = "181:Page Writes/Sec:4"
Const CacheFaultSec = "136:Cache Faults/Sec:4"
Const TransFaultSec = "159:Transition Faults/Sec:4"
Const UsedPageFilePerc = "174:% Page File Usage:6"

'INPUTS
Dim Host, MetricState, Username, Password, Domain
Host = WScript.Arguments(0)
MetricState = WScript.Arguments(1)
Username = WScript.Arguments(2)
Password = WScript.Arguments(3)
Domain = WScript.Arguments(4)

Dim arrMetrics, arrProcesses, top
arrMetrics = Split(MetricState,",")
Dim objSWbemLocator, objSWbemServices, colItems
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")

Dim Counter, objItem, FullUserName, OS
Counter = 0

	If Domain <> "" Then
		FullUserName = Domain & "\" & Username
	Else
		FullUserName = Username
	End If
	
	Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", FullUserName, Password)
	If Err.Number = -2147217308 Then
		
		Set objSWbemServices = objSWbemLocator.ConnectServer(Host, "root\cimv2", "", "")
		Err.Clear
	End If
	if Err.Number = -2147023174 Then
		CALL ShowError(4, Host)
		WScript.Quit (222)
	End If
		if Err.Number = -2147024891 Then
		CALL ShowError(2, Host)
	End If
	If Err Then CALL ShowError(1, Host)
	
	if Err.Number = 0 Then
		objSWbemServices.Security_.ImpersonationLevel = 3
		OS = GetOSVersion(objSWbemServices)
		if OS >= 6000 Then
			Set colItems = objSWbemServices.ExecQuery( "select Name, WorkingSet FROM Win32_PerfFormattedData_PerfProc_Process WHERE Name<>'_Total' And Name<>'Idle'",,16) 
			If colItems.Count <> 0 Then
				For Each objItem in colItems 
					arrProcesses = arrProcesses & objItem.Name & ":" & objItem.WorkingSet + "|"
				Next
				top = GetTop(arrProcesses,5)
			End If
		End If
		if Len(top) = 0 Then top = "-"
		Set colItems = objSWbemServices.ExecQuery( "SELECT FreePhysicalMemory,FreeVirtualMemory,TotalVirtualMemorySize,TotalVisibleMemorySize FROM Win32_OperatingSystem",,16) 
		If colItems.Count <> 0 Then
			For Each objItem in colItems 
				'FreePhysicalMemory
				If arrMetrics(0)=1 Then _
				CALL Output(FreePhysMem,FormatNumber(objItem.FreePhysicalMemory/1024), "", top)
				'FreeVirtualMemory
				If arrMetrics(1)=1 Then _
				CALL Output(FreeVirtMem,FormatNumber(objItem.FreeVirtualMemory/1024), "", top)
				'%Physical Memory Used
				If arrMetrics(2)=1 Then _
				CALL Output(UsedPhysMemPerc,FormatNumber((1-objItem.FreePhysicalMemory/objItem.TotalVisibleMemorySize)*100), "",top)
				'%Virtual Memory Used
				If arrMetrics(3)=1 Then _
				CALL Output(UsedVirtMemPerc,FormatNumber((1-objItem.FreeVirtualMemory/objItem.TotalVirtualMemorySize)*100), "",top)
			Next
		Else
			'If there is no response in WMI query
			CALL ShowError(5, Host)
		End If
		Set colItems = objSWbemServices.ExecQuery( _
			"SELECT PagesPerSec,PageFaultsPersec,PagesInputPersec,PagesOutputPersec,PageWritesPersec,CacheFaultsPersec,TransitionFaultsPersec from Win32_PerfFormattedData_PerfOS_Memory",,16) 
		If colItems.Count <> 0 Then
			For Each objItem in colItems 
				'PagesPerSec
				If arrMetrics(4)=1 Then _
				CALL Output(PagesSec,objItem.PagesPerSec, "",top)
				'PageFaultsPersec 
				If arrMetrics(5)=1 Then _
				CALL Output(PagesFaultSec,objItem.PageFaultsPersec, "",top)
				'PagesInputPersec 
				If arrMetrics(6)=1 Then _
				CALL Output(PagesInSec,objItem.PagesInputPersec, "",top)
				'PagesOutputPersec
				If arrMetrics(7)=1 Then _
				CALL Output(PagesOutSec,objItem.PagesOutputPersec, "",top)
				'PageWritesPersec
				If arrMetrics(8)=1 Then _
				CALL Output(PagesWriteSec,objItem.PageWritesPersec, "",top)
				'CacheFaultsPersec
				If arrMetrics(9)=1 Then _
				CALL Output(CacheFaultSec,objItem.CacheFaultsPersec, "",top)
				'TransitionFaultsPersec
				If arrMetrics(10)=1 Then _
				CALL Output(TransFaultSec,objItem.TransitionFaultsPersec, "",top)
			Next
		Else
			'If there is no response in WMI query
			CALL ShowError(5, Host)
		End If
		if OS >= 6000 Then
			Set colItems = objSWbemServices.ExecQuery( _
				"SELECT PercentUsage FROM Win32_PerfFormattedData_PerfOS_PagingFile WHERE Name='_Total'",,16)
			If colItems.Count <> 0 Then
				For Each objItem in colItems 
					'Page File PercentUsage
					If arrMetrics(11)=1 Then _
					CALL Output(UsedPageFilePerc,FormatNumber(objItem.PercentUsage),"",top)
				Next
			Else
				'If there is no response in WMI query
				CALL ShowError(5, Host)
			End If
		Else
			Set colItems = objSWbemServices.ExecQuery("Select AllocatedBaseSize,CurrentUsage from Win32_PageFileUsage",,16)
			If colItems.Count <> 0 Then
			For Each objItem in colItems
				If arrMetrics(11)=1 Then _
					CALL Output(UsedPageFilePerc,round((objItem.CurrentUsage*100)/objItem.AllocatedBaseSize,2), "",top)
			Next
			Else
				'If there is no response in WMI query
				CALL ShowError(5, Host)
			End If
		End If
        If Err.number <> 0 Then
			CALL ShowError(5, Host)
			Err.Clear
		End If
	Else
		Err.Clear
	End If


If Err Then 
	CALL ShowError(1,0)
Else
	WScript.Quit(0)
End If

Sub ShowError(ErrorCode, Param)
	Dim Msg
	Msg = "(" & Err.Number & ") " & Err.Description
	If ErrorCode=2 Then Msg = "Access is denied"
	If ErrorCode=3 Then Msg = "Wrong number of parameters on execution"
	If ErrorCode=4 Then Msg = "The specified target cannot be accessed"
	If ErrorCode=5 Then Msg = "There is no response in WMI or returned query is empty"
	WScript.Echo Msg
	WScript.Quit(ErrorCode)
End Sub

Sub Output(MetricID, MetricValue, MetricObject, MetricData)
	if MetricData = "" Then 
		MetricData = "-"
	End If
	If MetricObject <> "" Then
		If MetricValue <> "" Then
			WScript.Echo MetricID & "|" & MetricValue & "|" & MetricObject & "=" & MetricData & "|" 
		Else
			CALL ShowError(5, Host) 
		End If
	Else
		If MetricValue <> "" Then
			WScript.Echo MetricID & "|" & MetricValue & "=" & MetricData & "|"
		Else
			CALL ShowError(5, Host)
		End If
	End If
End Sub

Function GetOSVersion(SWbem)
	Dim colItems, objItem
	Set colItems = SWbem.ExecQuery("select BuildVersion from Win32_WMISetting",,16)
	For Each objItem in colItems
		GetOSVersion = CInt(objItem.BuildVersion)
	Next
End Function

Function GetTop(ValueList, TotalRecords)
	Dim Val, rs, Counter, out, exists
	Set rs = CreateObject("ADODB.RECORDSET")
	rs.Fields.append "Property", 200, 255
	rs.Fields.append "Value", 20, 25
	rs.CursorType = 3
	rs.Open
	exists = 0
	For Each objItem in Split(ValueList,"|") 
		If (objItem<>"") Then
			Val = Split(objItem,":")
			rs.AddNew
			rs.Fields("Property") = Val(0)
			rs.Fields("Value") = Val(1)
			rs.Update
			exists = 1
		End if
	Next
	rs.Sort = "Value DESC, Property"
	if exists = 1 and not rs.EOF then
		rs.MoveFirst
		Counter = 0
		Do Until rs.EOF OR Counter = TotalRecords
			If out = "" Then
				out = rs.Fields(0) & ":" & FormatNumber(CLng(rs.Fields(1))/1048576)
			Else
				out = out & ";" & rs.Fields(0) & ":" & FormatNumber(CLng(rs.Fields(1))/1048576)
			End If
			rs.MoveNext
			Counter = Counter + 1
		Loop
		Set rs = Nothing
		GetTop = out
	Else
		GetTop = ""
	End If
End Function


