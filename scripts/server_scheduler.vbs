
Option Explicit

'------------------
' Declare variables
'------------------
Dim oConn, oRS
Dim fsObj, f1
Dim strSQL
Dim dtimStart, strFileStamp, strFileName
Dim intServerCounter
Dim strStatus, strCenterPrint

dtimStart     = Now()
strFileStamp  = DatePart("yyyy", dtimStart) & DatePart("m", dtimStart) & DatePart("d", dtimStart)
strFileName   = "C:\Logs\server_scheduler" & strFileStamp & ".log"
'---------------
' Set up objects
'---------------
Call Initialize_Conn(oConn)

Set oRS    = CreateObject("ADODB.Recordset")
Set fsObj  = CreateObject("Scripting.FileSystemObject")
If fsObj.FileExists(strFileName) Then
	Set f1   = fsObj.OpenTextFile(strFileName, 8, True)
Else
	Set f1   = fsObj.CreateTextFile(strFileName, True)
End If
f1.WriteLine("Beginning Log " & dtimStart)

'-------------
' Main Program
'-------------

'' Process the reservations, lock down servers
strSQL = "SELECT r.ReservationID, s.ServerID, r.SadPassword, r.JoinPassword, "
strSQL = strSQL & " s.ServerIP, s.TelnetPort, s.TelnetPassword "
strSQL = strSQL & " FROM tbl_reservations r, tbl_servers s "
strSQL = strSQL & " WHERE (DateDiff(n, StartTime, GetDate()) >= -15) "
strSQL = strSQL & " AND (DateDiff(n, StartTime, GetDate()) <= 0) "
strSQL = strSQL & " AND IsLockedDown = 'N' "
strSQL = strSQL & " AND s.ServerID = r.ServerID "
oRS.Open strSQL, oConn, 3, 3
If Not(oRS.EOF and oRS.BOF) Then
	strCenterPrint = "Server is automatically being locked down by TWL Reservation System. This is a prescheduled event, please drop from this server."
	Do While Not(oRS.EOF)
		strStatus = LockDownServer(oRS("ServerIP"), oRS("TelnetPort"), oRS("TelnetPassword"), oRS("JoinPassword"), oRS("SadPassword"), strCenterPrint)
		If Left(strStatus, 4) = "PASS" Then
			oConn.Execute ("UPDATE tbl_reservations SET IsLockedDown = 'Y' WHERE ReservationID = " & oRS("ReservationID"))
		End If
		oConn.Execute ("UPDATE tbl_servers SET ScriptStatus = '" & Replace(strStatus, "'", "''") & " - " & oRS("ReservationID") & "' WHERE ServerID = " & oRS("ServerID"))
		intServerCounter = intServerCounter + 1
		oRS.MoveNext
	Loop
	f1.WriteLine(intServerCounter & " server(s) found to set reservations")
Else
	f1.WriteLine("No Servers found to set reservations")
End If
oRS.NextRecordset

'' Process the reservations, un-lock down servers
strSQL = "SELECT r.ReservationID, s.ServerID, s.SadPassword, s.JoinPassword, "
strSQL = strSQL & " s.ServerIP, s.TelnetPort, s.TelnetPassword "
strSQL = strSQL & " FROM tbl_reservations r, tbl_servers s "
strSQL = strSQL & " WHERE (DateDiff(n, EndTime, GetDate()) <= 15) "
strSQL = strSQL & " AND (DateDiff(n, EndTime, GetDate()) >= 0) "
strSQL = strSQL & " AND IsLockedDown = 'Y' "
strSQL = strSQL & " AND IsUnLocked = 'N' "
strSQL = strSQL & " AND s.ServerID = r.ServerID "
oRS.Open strSQL, oConn, 3, 3
If Not(oRS.EOF and oRS.BOF) Then
	strCenterPrint = "Server is automatically being unreserved by TWL reservation system. Please be aware, original join password has been restored."
	Do While Not(oRS.EOF)
		strStatus = LockDownServer(oRS("ServerIP"), oRS("TelnetPort"), oRS("TelnetPassword"), oRS("JoinPassword"), oRS("SadPassword"), strCenterPrint)
		If Left(strStatus, 4) = "PASS" Then
			oConn.Execute ("UPDATE tbl_reservations SET IsUnLocked = 'Y' WHERE ReservationID = " & oRS("ReservationID"))
		End If
		oConn.Execute ("UPDATE tbl_servers SET ScriptStatus = '" & Replace(strStatus, "'", "''") & " - " & oRS("ReservationID") & "' WHERE ServerID = " & oRS("ServerID"))
		intServerCounter = intServerCounter + 1
		oRS.MoveNext
	Loop
	f1.WriteLine(intServerCounter & " server(s) found to unlock reservations")
Else
	f1.WriteLine("No Servers found to unlock reservations")
End If
oRS.NextRecordset

'--------------
' Close Objects
'--------------
f1.WriteLine("Ending Log " & Now())
f1.WriteLine("-----------------------------------------------------")
f1.WriteBlankLines(1)
f1.Close()
Set oRS = Nothing
Set fsObj = Nothing
Call Close_Conn(oConn)

'-------------------
' Functions and Subs
'-------------------
Function LockDownServer(byVal strServerIP, byVal strServerPort, byVal strServerPassword, byVal strJoinPassword, byVal strSADPassword, byVal strCenterPrintMsg)
	Dim sObj, intCounter
	Dim AllSuccess

	Set sObj = CreateObject("TWLControl.ServerControl")
	intCounter = 0
	Call sObj.ChangeSADJoinPW(strServerIP, strServerPort, strServerPassword, strJoinPassword, strSADPassword, strCenterPrintMsg)
	Do While sObj.StatusMessage() = "" AND intCounter < 15
		intCounter = intCounter + 1
		wScript.Sleep(250)
	Loop
	AllSuccess = sObj.StatusMessage()
	Set sObj = Nothing
	
	If AllSuccess Then
		LockDownServer = "PASS: Passwords changed successfully."
	Else
		LockDownServer = "FAIL: Possible connection problem, niether password set."
	End If
End Function

Sub Initialize_Conn(byRef cn)
	Set oConn = CreateObject("ADODB.Connection")
	oConn.Open "file name=c:\twl.udl"
End Sub

Sub Close_Conn(byRef cn)
	cn.Close
	Set cn = Nothing
End Sub
