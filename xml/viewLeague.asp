<% Option Explicit %>
<!-- #include virtual="/include/xml.asp" -->
<%
Server.ScriptTimeout = 45

Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Dim strLeagueName, intLeagueID
strLeagueName = Request.QueryString("League")
If Len(Trim(strLeagueName)) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If

strSQL = "SELECT LeagueID, LeagueName FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "' AND LeagueActive = 1"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
	strLeagueName = oRs.Fields("LeagueName").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
oRs.NextRecordSet
Dim intConferenceID, intDivisionID, strConferenceName, strDivisionName
Dim intDivisionsShown, intRank
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<%
Response.Write "<leagueview>" & vbCrLf 
Response.Write vbTab & "<leagueinformation>" & vbCrLf
Response.Write vbTab & vbTab & "<leaguename>" & XMLEncode(Server.HTMLEncode(strLeagueName)) & "</leaguename>" & vbCrLf
Response.Write vbTab & vbTab & "<leaguelink>" & vbCrLf
Response.Write vbTab & vbTab & vbTab & "<httplink>" & Server.HTMLEncode("http://www.teamwarfare.com/xml/viewleague.asp?league=" & XMLEncode(Server.URLEncode(strLeagueName))) & "</httplink>" & vbCrLf
Response.Write vbTab & vbTab & vbTab & "<xmllink>" & Server.HTMLEncode("http://www.teamwarfare.com/xml/viewleague.asp?league=" & XMLEncode(Server.URLEncode(strLeagueName))) & "</xmllink>" & vbCrLf
Response.Write vbTab & vbTab & "</leaguelink>" & vbCrLf
Response.Write vbTab & "</leagueinformation>" & vbCrLf

strSQL = "SELECT ConferenceName, c.LeagueConferenceID, DivisionName, d.LeagueDivisionID "
strSQL = strSQL & " FROM tbl_league_conferences c "
strSQL = strSQL & " INNER JOIN tbl_league_divisions d "
strSQL = strSQL & " ON d.LeagueConferenceID = c.LeagueConferenceID "
strSQL = strSQL & " WHERE d.LeagueID = '" & intLeagueID & "'"
strSQL = strSQL & " ORDER BY ConferenceSortOrder, ConferenceName, DivisionSortOrder, DivisionName"
oRS.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intConferenceID = -1
	Do While Not(oRs.EOF)
		If intConferenceID <> oRs.Fields("LeagueConferenceID").Value Then
			If intConferenceID <> -1 Then
				Response.Write vbTab & "</conference>" & vbCrLf
			End If
			intConferenceID = oRs.Fields("LeagueConferenceID").Value
			strConferenceName = oRs.Fields("ConferenceName").Value
			Response.Write vbTab & "<conference name=""" & XMLEncode(strConferenceName) & """>" & vbCrLf
			intDivisionsShown = 0
		End If
		intDivisionID  = oRs.Fields("LeagueDivisionID").Value
		strDivisionName = oRs.Fields("DivisionName").Value
		Response.Write vbTab & vbTab & "<division name=""" & XMLEncode(Server.HTMLEncode(strDivisionName)) & """>" & vbCrLf
'			strSQL = "SELECT Top 5 lnkLeagueTeamID, TeamName, LeaguePoints, Rank, Wins, Losses, Draws, WinPct FROM "
			strSQL = "SELECT lnkLeagueTeamID, TeamName, TeamTag, LeaguePoints, Rank, Wins, Losses, Draws, NoShows, WinPct, RoundsWon, RoundsLost FROM "
			strSQL = strSQL & " lnk_league_team l "
			strSQL = strSQL & " INNER JOIN tbl_teams T "
			strSQL = strSQL & " ON t.TeamID = l.TeamID "
			strSQL = strSQL & " WHERE LeagueDivisionID = '" & intDivisionID & "' "
			strSQL = strSQL & " AND Active = 1 "
			strSQL = strSQL & " ORDER BY LeaguePoints DESC, Wins DESC, WinPct DESC, RoundsWon DESC, TeamName ASC "
			oRs2.Open strSQL, oConn
			If Not(oRs2.EOF AND oRs2.BOF) Then
				intRank = 0
				Do While Not (oRs2.EOF)
					intRank = intRank + 1 
					Response.Write vbTab & vbTab & vbTab & "<team rank=""" & intRank & """>" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & "<name>" & XMLEncode(Server.HTMLEncode(oRs2.Fields("TeamName").Value)) & "</name>" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & "<tag>" & XMLEncode(Server.HTMLEncode(oRs2.Fields("TeamTag").Value)) & "</tag>" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & "<standing "
					Response.Write " points=""" & oRs2.Fields("LeaguePoints").Value & """"
					Response.Write " wins=""" & oRs2.Fields("Wins").Value & """"
					Response.Write " losses=""" & oRs2.Fields("losses").Value & """"
					Response.Write " draws=""" & oRs2.Fields("draws").Value & """"
					Response.Write " noshows=""" & oRs2.Fields("noshows").Value & """"
					Response.Write " winpct=""" & FormatNumber(cInt(oRs2.Fields("WinPct").Value) / 10000, 3, 0) & """"
					Response.Write " roundswon=""" & oRs2.Fields("roundswon").Value & """"
					Response.Write " roundslost=""" & oRs2.Fields("roundslost").Value & """"
					Response.Write " />" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & "<teamlink>" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<httplink>" & Server.HTMLEncode("http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLEncode(oRs2.Fields("TeamName").Value))) & "</httplink>" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<xmllink>" & Server.HTMLEncode("http://www.teamwarfare.com/xmp/viewteam_v2.asp?team=" & XMLEncode(Server.URLEncode(oRs2.Fields("TeamName").Value))) & "</xmllink>" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & "</teamlink>" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & "</team>" & vbCrLf
					oRs2.MoveNext
				Loop
			End If
			oRs2.NextRecordSet
			Response.Write vbTab & vbTab & "</division>" & vbCrLf
		oRs.MoveNext
	Loop
	Response.Write vbTab & "</conference>" & vbCrLf
End If
Response.Write "</leagueview>" & vbCrLf

oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>