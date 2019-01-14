<%
Option Explicit
%>
<!-- #include virtual="/include/xml.asp" -->
<%
Server.ScriptTimeout = 45
Dim strLadderName, intLadderID
Dim oRS, strSQL, oConn

strLadderName = Request.Form("ladder")
If Len(strLadderName) = 0 Then
	strLadderName = Request.QueryString("ladder")
End if
If (strLadderName = "") Then
	Response.Write "You must specify a ladder in the querystring, in the form of: filename.asp?ladder=Tribes+2+CTF"
	Response.End 
End If

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Set oRS = Server.CreateObject("ADODB.RecordSet")

' Get Proper casing of the ladder name, and verify existance
strSQL = "SELECT LadderName, LadderID FROM tbl_ladders WHERE LadderName = '" & CheckString(strLadderName) & "'"
oRS.Open strSQL, oConn
If Not(oRS.EOF And oRS.BOF) Then
	strLadderName	= oRS.Fields("LadderName").Value 
	intLadderID		= oRS.Fields("LadderID").Value 
Else
	' Invalid ladder
	Response.Write "Invalid ladder."
	oRS.Close
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
	Response.End 
End If
oRS.NextRecordset 

Dim intPendingDays, intRecentDays

intRecentDays = Trim(Request.QueryString("recentdays"))
If Len(intRecentDays) = 0 Or Not(IsNumeric(intRecentDays)) Or IsNull(intRecentDays) Then
	intRecentDays = 2
ElseIf cInt(intRecentDays) > 5 Then
	intRecentDays = 2
End If

intPendingDays = Trim(Request.QueryString("pendingdays"))
If Len(intPendingDays) = 0 Or Not(IsNumeric(intPendingDays)) Or IsNull(intPendingDays) Then
	intPendingDays = 2
ElseIf cInt(intPendingDays) > 5 Then
	intPendingDays = 2
End If

Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf

Response.Write "<ladder>" & vbCrLf

	'----------------------------------
	' Ladder Information Section
	'----------------------------------
	Response.Write vbTab & "<ladderinfo>" & vbCrLf
	Response.Write vbTab & vbTab & "<name>" & XMLEncode(Server.HTMLEncode("" & strLadderName)) & "</name>" & vbCrLf
	Response.Write vbTab & vbTab & "<date>" & FormatDateTime(date(), 2) & "</date>" & vbCrLf
	Response.Write vbTab & vbTab & "<time>" & FormatDateTime(now(), 3) & "</time>" & vbCrLf
	Response.Write vbTab & "</ladderinfo>" & vbCrLf
	'----------------------------------
	' End Ladder Information Section
	'----------------------------------
	
	'----------------------------------
	' Recent History Section 
	'----------------------------------
	Response.Write vbTab & "<results>" & vbCrLf
	strSQL = "SELECT HistoryID, MatchID, WinnerName, LoserName, WinnerRank, LoserRank, MatchDate, MatchForfeit "
	strSQL = strSQL & " FROM vHistory "
	strSQL = strSQL & " WHERE DateDiff(dd, MatchDate, GetDate()) <= " & intRecentDays
	strSQL = strSQL & " AND MatchLadderID = '" & intLadderID & "'"
	strSQL = strSQL & " ORDER BY MatchDate DESC "
	oRS.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		Do While Not(oRS.EOF)
			Response.Write vbTab & vbTab & "<result historyid=""" & oRs.Fields("HistoryID").Value & """ matchid=""" & oRs.Fields("MatchID").Value & """>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "<matchforfeit>" & Server.HTMLEncode("" & cBool(oRs.Fields("MatchForfeit").Value)) & "</matchforfeit>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "<rank>" & Server.HTMLEncode("" & oRS.Fields("WinnerRank").Value) & "</rank>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "<winner name=""" & XMLEncode(Server.HTMLEncode("" & oRS.Fields("WinnerName").Value)) & """ rank=""" & Server.HTMLEncode("" & oRs.FIelds("WinnerRank").Value) & """>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & XMLEncode(Server.URLEncode("" & oRS.Fields("WinnerName").Value)) & "</xmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLEncode("" & oRS.Fields("WinnerName").Value)) & "</httplink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "</winner>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<loser name=""" & XMLEncode(Server.HTMLEncode("" & oRS.Fields("LoserName").Value)) & """ rank=""" & Server.HTMLEncode("" & oRs.FIelds("LoserRank").Value) & """>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & XMLEncode(Server.URLEncode("" & oRS.Fields("LoserName").Value)) & "</xmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLEncode("" & oRS.Fields("LoserName").Value)) & "</httplink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "</loser>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<date>" & FormatDateTime(oRS.Fields("MatchDate").Value, 2) & "</date>" & vbCrLf
			Response.Write vbTab & vbTab & "</result>" & vbCrLF
			oRS.MoveNext
		Loop			
	End If
	oRS.NextRecordset
	Response.Write vbTab & "</results>" & vbCrLf
	'----------------------------------
	' End Recent History Section 
	'----------------------------------
	
	'----------------------------------
	' Upcoming Matches Section 
	'----------------------------------
	Response.Write vbTab & "<pending>" & vbCrLf
	strSQL = "SELECT MatchID, DefenderName, AttackerName, "
	strSQL = strSQL & " DefenderWins, DefenderLosses, DefenderForfeits, "
	strSQL = strSQL & " AttackerWins, AttackerLosses, AttackerForfeits, "
	strSQL = strSQL & " MatchDate, DefenderRank, AttackerRank, MatchTime "
	strSQL = strSQL & " FROM vDisplayPending "
	strSQL = strSQL & " WHERE DateDiff(d, MatchDate, GetDate()) <= " & intPendingDays
	strSQL = strSQL & " AND MatchLadderID = '" & intLadderID & "'"
	strSQL = strSQL & " ORDER BY MatchDate ASC "
	oRS.Open strSQL, oConn
	If Not(oRS.EOF and oRS.BOF) Then
		Do While Not(oRS.EOF)
			Response.Write vbTab & vbTab & "<match matchid=""" & oRs.Fields("MatchID").Value & """>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "<rank>" & Server.HTMLEncode("" & oRS.Fields("DefenderRank").Value) & "</rank>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "<defender name=""" & XMLEncode(Server.HTMLEncode("" & oRS.Fields("DefenderName").Value)) & """ rank=""" & Server.HTMLEncode("" & oRs.Fields("DefenderRank").Value) & """"
			Response.Write " wins=""" & Server.HTMLEncode("" & oRs.Fields("DefenderWins").Value) & """ losses=""" & Server.HTMLEncode("" & oRs.Fields("DefenderLosses").Value) & """ forfeits=""" & Server.HTMLEncode("" & oRs.Fields("DefenderForfeits").Value) & """"
			Response.Write ">" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & XMLEncode(Server.URLEncode("" & oRS.Fields("DefenderName").Value)) & "</xmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLEncode("" & oRS.Fields("DefenderName").Value)) & "</httplink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "</defender>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<attacker name=""" & XMLEncode(Server.HTMLEncode("" & oRS.Fields("AttackerName").Value)) & """ rank=""" & Server.HTMLEncode("" & oRs.Fields("AttackerRank").Value) & """"
			Response.Write " wins=""" & Server.HTMLEncode("" & oRs.Fields("AttackerWins").Value) & """ losses=""" & Server.HTMLEncode("" & oRs.Fields("AttackerLosses").Value) & """ forfeits=""" & Server.HTMLEncode("" & oRs.Fields("AttackerForfeits").Value) & """"
			Response.Write ">" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & XMLEncode(Server.URLEncode("" & oRS.Fields("AttackerName").Value)) & "</xmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLEncode("" & oRS.Fields("AttackerName").Value)) & "</httplink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "</attacker>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<date>" & FormatDateTime(oRS.Fields("MatchDate").Value, 2) & "</date>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<time>" & FormatDateTime(cDate(oRS.Fields("MatchTime").Value), 3) & "</time>" & vbCrLf
			Response.Write vbTab & vbTab & "</match>" & vbCrLf
			oRS.MoveNext
		Loop
	End If
	oRS.NextRecordset 
	Response.Write vbTab & "</pending>" & vbCrLf
	'----------------------------------
	' End Upcoming Matches Section 
	'----------------------------------
	
	
Response.Write "</ladder>" & vbCrLf

oConn.Close
Set oConn = Nothing
Set oRs = Nothing

Function CheckString(byVal strData)
	If Not(IsNull(strData)) Then
		CheckString = Replace(strData, "'", "''")
	Else
		CheckString = ""
	End If
End Function
%>