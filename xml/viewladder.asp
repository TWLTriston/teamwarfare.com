<% Option Explicit %>
<!-- #include virtual="/include/xml.asp" -->
<%
Response.Buffer = True

Dim strLadderName, intLadderID
strLadderName = Request.QueryString("ladder")
If Len(strLadderName) = 0 Then
	Response.Write "No ladder passed. No data returned."
	Response.End 
End If

Dim strSQL, oConn, oRS, oRS2

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

strSQL = "SELECT LadderID FROM tbl_ladders WHERE LadderName = '" & Replace(strLadderName, "'", "''") & "'"
oRS.Open strSQL, oConn
If Not(oRS.EOF and oRS.BOF) Then
	intLadderID = oRS.Fields("LadderID").Value 
Else
	oRS.Close
	Set oRS = Nothing
	Set oRS2 = Nothing
	oConn.Close
	Set oConn = Nothing
	Response.Write "Bad ladder name passed. No data returned."
	Response.End 
End If
oRS.NextRecordset 

Dim intTopRank, intBottomRank
Dim intMaxRanks, intDefaultRange

intDefaultRange = 25
intMaxRanks		= 100
intTopRank		= Request.QueryString("rank_start")
intBottomRank	= Request.QueryString("rank_end")

If Not(IsNumeric(intTopRank)) Or Len(intTopRank) = 0 Then
	intTopRank = 1
	intBottomRank = 25
Else
	intTopRank = cInt(intTopRank)
End If

If Not(IsNumeric(intBottomRank)) Or Len(intBottomRank) = 0 Then
	intBottomRank = intTopRank + intDefaultRange
Else
	intBottomRank = cInt(intBottomRank) 
	If intBottomRank - intTopRank > intMaxRanks Then
		intBottomRank = intTopRank + intDefaultRange
	End If
End If

Dim strTeamName, intMaps, strRules, strStatus

Dim strEnemyName, strMatchDate, strMapArray(6), i
Dim newMDate, mm, dd, pDate

Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLF
Response.Write "<ladderview>" & vbCrLF
strSQL = "SELECT * FROM vLadder WHERE LadderName='" & Replace(strLadderName, "'", "''") & "' "
strSQL = strSQL & " AND Rank >= " & intTopRank
strSQL = strSQL & " AND Rank <= " & intBottomRank
oRs.Open strSQL, oConn
If Not(oRS.EOF AND oRS.BOF) Then
	intMaps				= oRS.Fields("Maps").Value
	strRules			= oRS.Fields("LadderRules").Value 
	
	Response.Write vbTab & "<ladderinformation>" & vbCrLF
	Response.Write vbTab & vbTab & "<laddername>" & XMLEncode(Server.HTMLEncode(strLadderName)) & "</laddername>" & vbCrLF
	Response.Write vbTab & vbTab & "<ladderlink>" & vbCrLF
	Response.Write vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewladder.asp?ladder=" & XMLEncode(Server.URLencode(strLadderName)) & "</httplink>" & vbCrLF
	Response.Write vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewladder.asp?ladder=" & XMLEncode(Server.URLencode(strLadderName)) & "</xmllink>" & vbCrLF
	Response.Write vbTab & vbTab & vbTab & "<rulelink>http://www.teamwarfare.com/rules.asp?set=" & XMLEncode(Server.URLencode("" & strRules)) & "</rulelink>" & vbCrLF
	Response.Write vbTab & vbTab & "</ladderlink>" & vbCrLF
	Response.Write vbTab & "</ladderinformation>" & vbCrLF
	
	Response.Write vbTab & "<teams>" & vbCrLF
	Do While Not(oRS.EOF) 
		strTeamName		= oRS.Fields("TeamName").Value
		Response.Write vbTab & vbTab & "<team rank=""" & oRS.Fields("rank").Value & """>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<name>" & XMLEncode(Server.HTMLEncode(strTeamName)) & "</name>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<tag>" & XMLEncode(Server.HTMLEncode(oRS.Fields("TeamTag").Value)) & "</tag>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<wins>" & Server.HTMLEncode(oRS.Fields("Wins").Value) & "</wins>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<losses>" & Server.HTMLEncode(oRS.Fields("Losses").Value) & "</losses>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<forfeits>" & Server.HTMLEncode(oRS.Fields("Forfeits").Value) & "</forfeits>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<teamlink>" & vbCrLF
		Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLencode(strTeamName)) & "</httplink>" & vbCrLF
		Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & XMLEncode(Server.URLencode(strTeamName)) & "</xmllink>" & vbCrLF
		Response.Write vbTab & vbTab & vbTab & "</teamlink>" & vbCrLF
		' Write out the match status
		strStatus = oRS.Fields("Status").Value 
		Select Case Left(uCase(strStatus), 6)
			Case "DEFEND", "ATTACK"
				If  Left(uCase(strStatus), 6)  = "DEFEND" Then
					strsql = "select m.MatchAttackerID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID,  m.MatchMap4ID,  m.MatchMap5ID, t.teamname, t.teamtag "
					strsql = strsql & " from tbl_Matches m, tbl_teams t, lnk_t_l lnk "
					strsql = strsql & " where m.matchdefenderID = " & ors.Fields(4).Value 
					strsql = strsql & " and m.MatchLadderID=" & intLadderID
					strsql = strsql & " AND t.teamid = lnk.teamid "
					strsql = strsql & " AND lnk.tllinkid = m.MatchAttackerID "
					ors2.Open strSQL, oconn
					if not (ors2.EOF and ors2.BOF) then
						strEnemyName = ors2.Fields("TeamName").Value
						strMatchDate = ors2.Fields("MatchDate").Value
						strMapArray(1) = ors2.fields("MatchMap1ID").value
						strMapArray(2) = ors2.fields("MatchMap2ID").value
						strMapArray(3) = ors2.fields("MatchMap3ID").value
						strMapArray(4) = ors2.fields("MatchMap4ID").value
						strMapArray(5) = ors2.fields("MatchMap5ID").value
					end if
					ors2.nextrecordset 
				Else
					strsql = "select m.MatchDefenderID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID, m.MatchMap4ID,  m.MatchMap5ID, t.teamname, t.teamtag "
					strsql = strsql & " from tbl_Matches m, tbl_teams t, lnk_t_l lnk "
					strsql = strsql & " where m.MatchAttackerID = " & ors.Fields(4).Value 
					strsql = strsql & " and m.MatchLadderID=" & intLadderID
					strsql = strsql & " AND t.teamid = lnk.teamid "
					strsql = strsql & " AND lnk.tllinkid = m.matchdefenderID "
					ors2.Open strSQL, oconn
					if not (ors2.EOF and ors2.BOF) then
						strEnemyName = ors2.Fields("TeamName").Value
						strMatchDate = ors2.Fields("MatchDate").Value
						strMapArray(1) = ors2.fields("MatchMap1ID").value
						strMapArray(2) = ors2.fields("MatchMap2ID").value
						strMapArray(3) = ors2.fields("MatchMap3ID").value
						strMapArray(4) = ors2.fields("MatchMap4ID").value
						strMapArray(5) = ors2.fields("MatchMap5ID").value
					End If
					ors2.nextrecordset 				
				End If
				if strMatchDate <> "TBD" then
					newMDate = right(strMatchDate, len(strMatchDate)-instr(1, strMatchDate, ","))
					newMDate = Left(newmDate, (len(newMDate) - 4))
					newMDate = formatdatetime(newMDate, 2)
					mm=month(newmdate)
					dd=day(newmdate)
					pDate=mm & "/" & dd
				else
					pdate="TBD"
				End If
 				Response.Write vbTab & vbTab & vbTab & "<match status=""" & XMLEncode(Server.HTMLEncode(strStatus)) & """>" & vbCrLF
 				Response.Write vbTab & vbTab & vbTab & vbTab & "<opponentname>" & XMLEncode(Server.HTMLEncode(strEnemyName)) & "</opponentname>" & vbCrLF
				Response.Write vbTab & vbTab & vbTab & vbTab & "<opponentlink>" & vbCrLF
				Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<httplink>http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLencode(strEnemyName)) & "</httplink>" & vbCrLF
				Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & XMLEncode(Server.URLencode(strEnemyName)) & "</xmllink>" & vbCrLF
				Response.Write vbTab & vbTab & vbTab & vbTab & "</opponentlink>" & vbCrLF
 				Response.Write vbTab & vbTab & vbTab & vbTab & "<matchdate>" & Server.HTMLEncode(pDate) & "</matchdate>" & vbCrLF
				if pdate <> "TBD" then
					For i = 1 to intMaps
						Response.Write vbTab & vbTab & vbTab & vbTab & "<map name=""" & XMLEncode(Server.HTMLEncode(strMapArray(i))) & """/>" & vbCrLf
					Next
				end if
 				Response.Write vbTab & vbTab & vbTab & "</match>" & vbCrLF
			Case "IMMUNE", "DEFEAT", "RESTIN"
				Response.Write vbTab & vbTab & vbTab & "<match status=""" & XMLEncode(Server.HTMLEncode(strStatus)) & """/>" & vbCrLF
			Case Else
				Response.Write vbTab & vbTab & vbTab & "<match status=""" & XMLEncode(Server.HTMLEncode("Open")) & """/>" & vbCrLF
		End Select
		Response.Write vbTab & vbTab & "</team>" & vbCrLF
		oRS.MoveNext
	Loop
	Response.Write vbTab & "</teams>" & vbCrLF
End If
oRs.Close 
Response.Write "</ladderview>" & vbCrLF

oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>