<% Option Explicit %>
<%
Response.Buffer = True

Dim strLadderName, intLadderID
Dim strSQL, oConn, oRS, oRS2

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLF
Response.Write "<tribesarenaxml>" & vbCrLF

ProcessLadder ( "Tribes Arena")
ProcessLadder ( "Tribes 2 Arena")
ProcessLadder ( "Tribes 2 Arena 2v2")
ProcessPlayerLadder ("Tribes 2 Duel")

LadderMatches ( "Tribes Arena")
LadderMatches ( "Tribes 2 Arena")
LadderMatches ( "Tribes 2 Arena 2v2")

Response.Write "</tribesarenaxml>" & vbCrLF

oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing

Sub ProcessLadder(strLadderName)
	strSQL = "SELECT LadderID FROM tbl_ladders WHERE LadderName = '" & Replace(strLadderName, "'", "''") & "'"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF and oRS.BOF) Then
		intLadderID = oRS.Fields("LadderID").Value 
	End If
	oRS.NextRecordset 
	
	Dim strTeamName, intMaps, strRules, strStatus
	
	Dim strEnemyName, strMatchDate, strMapArray(6), i
	Dim newMDate, mm, dd, pDate
	
	Response.Write "<ladderinformation>" & vbCrLF
	strSQL = "SELECT * FROM vLadder WHERE LadderName='" & Replace(strLadderName, "'", "''") & "' "
	strSQL = strSQL & " AND Rank <= 5"
	oRs.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		Response.Write vbTab & "<laddername>" & Server.HTMLEncode(strLadderName) & "</laddername>" & vbCrLF
		
		Response.Write vbTab & "<teams>" & vbCrLF
		Do While Not(oRS.EOF) 
			strTeamName		= oRS.Fields("TeamName").Value
			Response.Write vbTab & vbTab & "<team rank=""" & oRS.Fields("rank").Value & """>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<name>" & Server.HTMLEncode(strTeamName) & "</name>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<tag>" & Server.HTMLEncode(oRS.Fields("TeamTag").Value) & "</tag>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<wins>" & Server.HTMLEncode(oRS.Fields("Wins").Value) & "</wins>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<losses>" & Server.HTMLEncode(oRS.Fields("Losses").Value) & "</losses>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<forfeits>" & Server.HTMLEncode(oRS.Fields("Forfeits").Value) & "</forfeits>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<teamlink>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLencode(strTeamName) & "</httplink>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & Server.URLencode(strTeamName) & "</xmllink>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "</teamlink>" & vbCrLF
			' Write out the match status
			Response.Write vbTab & vbTab & "</team>" & vbCrLF
			oRS.MoveNext
		Loop
		Response.Write vbTab & "</teams>" & vbCrLF
	End If
	oRs.Close 
	Response.Write "</ladderinformation>" & vbCrLF
End Sub

Sub ProcessPlayerLadder(strLadderName)
	strSQL = "SELECT PlayerLadderID FROM tbl_playerladders WHERE PlayerLadderName = '" & Replace(strLadderName, "'", "''") & "'"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF and oRS.BOF) Then
		intLadderID = oRS.Fields("PlayerLadderID").Value 
	End If
	oRS.NextRecordset 
	

	Dim strTeamName, intMaps, strRules, strStatus
	
	Dim strEnemyName, strMatchDate, strMapArray(6), i
	Dim newMDate, mm, dd, pDate
	
	Response.Write "<duelladderinformation>" & vbCrLF
	strSQL = "SELECT * FROM vPlayerLadder WHERE PlayerLadderName='" & Replace(strLadderName, "'", "''") & "' "
	strSQL = strSQL & " AND Rank <= 5"
	oRs.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then

		Response.Write vbTab & "<laddername>" & Server.HTMLEncode(strLadderName) & "</laddername>" & vbCrLF
		
		Response.Write vbTab & "<members>" & vbCrLF
		Do While Not(oRS.EOF) 
			strTeamName		= oRS.Fields("PlayerHandle").Value
			Response.Write vbTab & vbTab & "<member rank=""" & oRS.Fields("rank").Value & """>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<name>" & Server.HTMLEncode(strTeamName) & "</name>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<wins>" & Server.HTMLEncode(oRS.Fields("Wins").Value) & "</wins>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<losses>" & Server.HTMLEncode(oRS.Fields("Losses").Value) & "</losses>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<forfeits>" & Server.HTMLEncode(oRS.Fields("Forfeits").Value) & "</forfeits>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<memberlink>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewplayer.asp?player=" & Server.URLencode(strTeamName) & "</httplink>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "</memberlink>" & vbCrLF
			' Write out the match status
			Response.Write vbTab & vbTab & "</member>" & vbCrLF
			oRS.MoveNext
		Loop
		Response.Write vbTab & "</members>" & vbCrLF
	End If
	oRs.Close 
	Response.Write "</duelladderinformation>" & vbCrLF
End Sub

Sub LadderMatches (strLaddername)
	Dim intRecentDays, intPendingDays
	intRecentDays = 21
	intPendingDays = 14

	strSQL = "SELECT LadderID FROM tbl_ladders WHERE LadderName = '" & Replace(strLadderName, "'", "''") & "'"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF and oRS.BOF) Then
		intLadderID = oRS.Fields("LadderID").Value 
	End If
	oRS.NextRecordset 

	'----------------------------------
	' Recent History Section 
	'----------------------------------
	Response.Write "<matches ladder=""" & Server.HTMLEncode(strLadderName & "") & """>" & vbCrLf
	Response.Write vbTab & "<results recentdays=""" & intRecentDays & """>" & vbCrLf
	strSQL = "SELECT WinnerName, LoserName, WinnerRank, MatchDate "
	strSQL = strSQL & " FROM vHistory "
	strSQL = strSQL & " WHERE DateDiff(dd, MatchDate, GetDate()) <= " & intRecentDays
	strSQL = strSQL & " AND MatchLadderID = '" & intLadderID & "'"
	strSQL = strSQL & " ORDER BY MatchDate DESC "
	oRS.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		Do While Not(oRS.EOF)
			Response.Write vbTab & vbTab & "<result>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "<rank>" & Server.HTMLEncode("" & oRS.Fields("WinnerRank").Value) & "</rank>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "<winner name=""" & Server.HTMLEncode("" & oRS.Fields("WinnerName").Value) & """>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & Server.URLEncode("" & oRS.Fields("WinnerName").Value) & "</xmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode("" & oRS.Fields("WinnerName").Value) & "</httplink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "</winner>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<loser name=""" & Server.HTMLEncode("" & oRS.Fields("LoserName").Value) & """>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & Server.URLEncode("" & oRS.Fields("LoserName").Value) & "</xmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode("" & oRS.Fields("LoserName").Value) & "</httplink>" & vbCrLf
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
	Response.Write vbTab & "<pending pendingdays=""" & intPendingDays & """>" & vbCrLf
	strSQL = "SELECT DefenderName, AttackerName, MatchDate, DefenderRank, MatchTime "
	strSQL = strSQL & " FROM vDisplayPending "
	strSQL = strSQL & " WHERE DateDiff(d, MatchDate, GetDate()) <= " & intPendingDays
	strSQL = strSQL & " AND MatchLadderID = '" & intLadderID & "'"
	strSQL = strSQL & " ORDER BY MatchDate ASC "
	oRS.Open strSQL, oConn
	If Not(oRS.EOF and oRS.BOF) Then
		Do While Not(oRS.EOF)
			Response.Write vbTab & vbTab & "<match>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "<rank>" & Server.HTMLEncode("" & oRS.Fields("DefenderRank").Value) & "</rank>" & vbCrLF
			Response.Write vbTab & vbTab & vbTab & "<defender name=""" & Server.HTMLEncode("" & oRS.Fields("DefenderName").Value) & """>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & Server.URLEncode("" & oRS.Fields("DefenderName").Value) & "</xmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode("" & oRS.Fields("DefenderName").Value) & "</httplink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "</defender>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<attacker name=""" & Server.HTMLEncode("" & oRS.Fields("AttackerName").Value) & """>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & Server.URLEncode("" & oRS.Fields("AttackerName").Value) & "</xmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode("" & oRS.Fields("AttackerName").Value) & "</httplink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "</attacker>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<date>" & FormatDateTime(oRS.Fields("MatchDate").Value, 2) & "</date>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<time>" & FormatDateTime(cDate(oRS.Fields("MatchTime").Value), 3) & "</time>" & vbCrLf
			Response.Write vbTab & vbTab & "</match>" & vbCrLf
			oRS.MoveNext
		Loop
	End If
	oRS.NextRecordset 
	Response.Write vbTab & "</pending>" & vbCrLf
	Response.Write "</matches>" & vbCrLf
	'----------------------------------
	' End Upcoming Matches Section 
	'----------------------------------
End Sub
%>