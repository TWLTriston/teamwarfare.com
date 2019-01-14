<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("League") & ": " & Request.QueryString("Conference") & ": " & Request.QueryString("Division"), """", "&quot;") 

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strLeagueName, intLeagueID
Dim intConferenceID, intDivisionID, strConferenceName, strDivisionName
Dim intDivisionsShown, intRank, intLinkID

strLeagueName = Request.QueryString("League")
If Len(Trim(strLeagueName)) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If

strConferenceName = Request.QueryString("Conference")
If Len(Trim(strConferenceName)) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If

strDivisionName = Request.QueryString("Division")
If Len(Trim(strDivisionName)) = 0 Then
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

Dim bLeagueAdmin
bLeagueAdmin = IsLeagueAdminByID(intLeagueID)

strSQL = "SELECT LeagueConferenceID, ConferenceName FROM tbl_league_conferences WHERE ConferenceName= '" & CheckString(strConferenceName) & "' AND LeagueID = '" & intLeagueID & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intConferenceID = oRs.Fields("LeagueConferenceID").Value
	strConferenceName = oRs.Fields("ConferenceName").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
oRs.NextRecordSet

strSQL = "SELECT LeagueDivisionID, DivisionName FROM tbl_league_divisions WHERE DivisionName = '" & CheckString(strDivisionName) & "' AND LeagueID = '" & intLeagueID & "' AND LeagueConferenceID = '" & intConferenceID & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intDivisionID = oRs.Fields("LeagueDivisionID").Value
	strDivisionName = oRs.Fields("DivisionName").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
oRs.NextRecordSet
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("")
%>
<table BORDER="0" cellspacing="0" cellpadding="0" width="97%">
<tr>
	<td CLASS="pageheader">
		<a href="viewleague.asp?league=<%=Server.URLEncode(strLeagueName)%>"><%=strLeagueName%> League</a> &raquo;
		<a href="viewleagueconference.asp?league=<%=Server.URLEncode(strLeagueName)%>&conference=<%=Server.URLEncode(strConferenceName)%>"><%=strConferenceName%> Conference</a> &raquo;
		<%=strDivisionName%> Division</a>
		</td>
</tr>
<tr>
	<td>&nbsp;&nbsp;<a href="leaguehistory.asp?league=<%=Server.URLEncode(strLeagueName)%>">Last Week's Matches</a></td>
</tr>
</table>
<br />
<table border="0" cellspacing="0" cellpadding="0" width="97%" class="cssBordered"align="center">
<tr>
	<th bgcolor="#000000" width="20">R</th>
	<th bgcolor="#000000">Team</th>
	<th bgcolor="#000000" width="20">P</th>
	<th bgcolor="#000000" width="20">W</th>
	<th bgcolor="#000000" width="20">L</th>
	<th bgcolor="#000000" width="20">D</th>
	<th bgcolor="#000000" width="20">N</th>
	<th bgcolor="#000000" width="20">RW</th>
	<th bgcolor="#000000" width="20">RL</th>
	<th bgcolor="#000000" width="50">pct</th>
	<th bgcolor="#000000">Next Match</th>
	<% if bSysAdmin Or bLeagueAdmin THen %>
	<th bgcolor="#000000">Kick</th>
	<% End If %>
</tr>
<%
strSQL = "SELECT lnkLeagueTeamID, TeamName, TeamTag, LeaguePoints, Rank, Wins, Losses, Draws, NoShows, WinPct, RoundsWon, RoundsLost FROM "
strSQL = strSQL & " lnk_league_team l "
strSQL = strSQL & " INNER JOIN tbl_teams T "
strSQL = strSQL & " ON t.TeamID = l.TeamID "
strSQL = strSQL & " WHERE LeagueDivisionID = '" & intDivisionID & "' "
strSQL = strSQL & " AND Active = 1 "
strSQL = strSQL & " ORDER BY LeaguePoints DESC, Wins DESC, WinPct DESC, RoundsWon DESC, TeamName ASC"
oRs2.Open strSQL, oConn
If Not(oRs2.EOF AND oRs2.BOF) Then
	intRank = 0
	Do While Not (oRs2.EOF)
		intRank = intRank + 1
		intLinkID = oRs2.Fields("lnkLeagueTeamID").Value
		%>
		<tr>
			<td bgcolor="<%=bgcone%>"><%=intRank%>. </td>
			<td bgcolor="<%=bgctwo%>"><a href="viewteam.asp?team=<%=Server.URLEncode(orS2.Fields("TeamName").Value)%>"><%=Server.HTMLEncode(oRs2.Fields("TeamName").Value & " - " & oRS2.Fields("TeamTag").Value)%></a></td>
			<td bgcolor="<%=bgcone%>"align="center"><%=oRs2.Fields("LeaguePoints").Value%></td>
			<td bgcolor="<%=bgctwo%>"align="center"><%=oRs2.Fields("Wins").Value%></td>
			<td bgcolor="<%=bgcone%>"align="center"><%=oRs2.Fields("Losses").Value%></td>
			<td bgcolor="<%=bgctwo%>"align="center"><%=oRs2.Fields("Draws").Value%></td>
			<td bgcolor="<%=bgctwo%>"align="center"><%=oRs2.Fields("NoShows").Value%></td>
			<td bgcolor="<%=bgcone%>"align="center"><%=oRs2.Fields("RoundsWon").Value%></td>
			<td bgcolor="<%=bgctwo%>"align="center"><%=oRs2.Fields("RoundsLost").Value%></td>
			<td bgcolor="<%=bgcone%>"align="center"><%=FormatNumber(cInt(oRs2.Fields("WinPct").Value) / 10000, 3, 0)%></td>
			<td bgcolor="<%=bgcone%>"align="center"><%
			strSQL = "SELECT TOP 1 LeagueMatchID, MatchDate, "
			strSQL = strSQL & " OpponentName = CASE WHEN HomeTeamLinkID='" & intLinkID & "' THEN "
			strSQL = strSQL & "		(SELECT TeamName FROM tbl_teams t INNER JOIN lnk_league_team lnk ON t.TeamID = lnk.TeamID WHERE lnk.lnkLeagueTeamID=m.VisitorTeamLinkID)"
			strSQL = strSQL & "		ELSE (SELECT TeamName FROM tbl_teams t INNER JOIN lnk_league_team lnk ON t.TeamID = lnk.TeamID WHERE lnk.lnkLeagueTeamID=m.HomeTeamLinkID)"
			strSQL = strSQL & "		END"
			strSQL = strSQL & " FROM tbl_league_matches m"
			strSQL = strSQL & " WHERE (HomeTeamLinkID = '" & intLinkID & "' "
			strSQL = strSQL & " OR VisitorTeamLinkID = '" & intLinkID & "') "
			strSQL = strSQL & " ORDER BY MatchDate ASC"
			oRs.Open strSQL, oConn
			If Not(oRs.EOF AND oRs.BOF) Then
				Do While Not(oRs.EOF)
					%>
					<a href="viewteam.asp?team=<%=Server.URLEncode(oRS.Fields("OpponentName").Value & "")%>"><%=Server.HTMLEncode(oRS.Fields("OpponentName").Value & "")%></a><br />
					<%
					If Not(IsNull(oRS.Fields("MatchDate").Value)) Then
						Response.Write FormatDateTime(oRS.Fields("MatchDate").Value, 2)
					End if
					oRs.Movenext
				Loop
			Else
				%>
				<b>No matches currently scheduled</b>
				<%
			End If
			oRs.NextRecordSet
			%></td>
			<% if bSysAdmin Or bLeagueAdmin THen %>
			<td bgcolor=<%=bgcone%> align="center"><a href="saveitem.asp?League=<%=Server.URLEncode(strLeagueName & "")%>&Conference=<%=Server.URLEncode(strConferenceName & "")%>&Division=<%=Server.URLEncode(strDivisionName & "")%>&SaveType=LeagueBumpBack&lnkLeagueTeamID=<%=intLinkID%>&LeagueID=<%=intLeagueID%>">Kick</a></td>
			<% ENd If %>
		</tr>
		<%
		oRs2.MoveNext
	Loop
Else
	%>
	<tr>
		<td bgcolor="<%=bgcone%>" colspan="9"><i>No teams are in this division</td>
	</tr>
	<%
End If
oRs2.NextRecordSet
%>
</table>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>