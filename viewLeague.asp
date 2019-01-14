<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("League"), """", "&quot;") 

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
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("")
%>

<table BORDER="0" cellspacing="0" cellpadding="0" width="97%">
<tr>
	<td CLASS="pageheader"><%=strLeagueName%> League</td>
</tr>
<tr>
	<td>
		&nbsp;&nbsp;<a href="leaguehistory.asp?league=<%=Server.URLEncode(strLeagueName)%>">Last Week's Matches</a> / <a href="xml/leaguehistory.asp?league=<%=Server.URLEncode(strLeagueName)%>">XML Version</a> <br />
		&nbsp;&nbsp;<a href="viewleaguematches.asp?league=<%=Server.URLEncode(strLeagueName)%>">This Week's Matches</a> / <a href="xml/viewleaguematches.asp?league=<%=Server.URLEncode(strLeagueName)%>">XML Version</a><br />
		&nbsp;&nbsp;<a href="viewleaguematches.asp?league=<%=Server.URLEncode(strLeagueName)%>&x=2">Next Week's Matches</a> / <a href="xml/viewleaguematches.asp?league=<%=Server.URLEncode(strLeagueName)%>&x=2">XML Version</a><br />
		<br />
		&nbsp;&nbsp;<a href="xml/viewleague.asp?league=<%=Server.URLEncode(strLeagueName)%>">XML Standings</a></td>
</tr>
</table>
<%
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
			IF intDivisionsShown mod 2 = 1 THen
				Response.Write "<td bgcolor=""#000000"">&nbsp;</td></tr>"
			End If
			If intConferenceID <> -1 Then
				%>
					</table>
					</td></tr>
				</table>
				<%
			End If
			intConferenceID = oRs.Fields("LeagueConferenceID").Value
			strConferenceName = oRs.Fields("ConferenceName").Value
			%>
			<br /><br />
				<table class="cssBordered" width="100%">
					<tr>
						<th colspan="2" bgcolor="<%=bgcone%>"><a href="viewleagueconference.asp?league=<%=Server.URLEncode(strLeagueName)%>&conference=<%=Server.URLEncode(strConferenceName)%>"><%=strConferenceName%> Conference</a></th>
					</tr>
					</table>
					<table width="100%">
			<%
			intDivisionsShown = 0
		End If
		If (intDivisionsShown mod 2 = 0) Then
			If intDivisionsShown > 0 Then
				Response.Write "</tr>"
			End If
			Response.Write "<tr>"
		End If
		intDivisionsShown = intDivisionsShown + 1
		intDivisionID = oRs.Fields("LeagueDivisionID").Value
		strDivisionName = oRs.Fields("DivisionName").Value
		%>
		<td width="50%" valign="top">
			<table class="cssBordered" width="100%">
			<tr>
				<th colspan="8" bgcolor="#000000"><a href="viewleaguedivision.asp?league=<%=Server.URLEncode(strLeagueName)%>&conference=<%=Server.URLEncode(strConferenceName)%>&division=<%=Server.URLEncode(strDivisionName)%>"><%=strDivisionName%> Division</a></th>
			</tr>
			<tr>
				<th bgcolor="#000000" width="20">R</th>
				<th bgcolor="#000000">Team</th>
				<th bgcolor="#000000" width="20">P</th>
				<th bgcolor="#000000" width="20">W</th>
				<th bgcolor="#000000" width="20">L</th>
				<th bgcolor="#000000" width="20">D</th>
				<th bgcolor="#000000" width="20">N</th>
				<th bgcolor="#000000" width="50">pct</th>
			</tr>
			<%
'			strSQL = "SELECT Top 5 lnkLeagueTeamID, TeamName, LeaguePoints, Rank, Wins, Losses, Draws, WinPct FROM "
			strSQL = "SELECT lnkLeagueTeamID, TeamName, LeaguePoints, Rank, Wins, Losses, Draws, NoShows, WinPct FROM "
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
					%>
					<tr>
						<td bgcolor="<%=bgcone%>"><%=intRank%>. </td>
						<td bgcolor="<%=bgctwo%>"><a href="viewteam.asp?team=<%=Server.URLEncode(orS2.Fields("TeamName").Value)%>"><%=Server.HTMLEncode(oRs2.Fields("TeamName").Value & "")%></a></td>
						<td  bgcolor="<%=bgcone%>"align="center"><%=oRs2.Fields("LeaguePoints").Value%></td>
						<td  bgcolor="<%=bgctwo%>"align="center"><%=oRs2.Fields("Wins").Value%></td>
						<td  bgcolor="<%=bgcone%>"align="center"><%=oRs2.Fields("Losses").Value%></td>
						<td  bgcolor="<%=bgctwo%>"align="center"><%=oRs2.Fields("Draws").Value%></td>
						<td  bgcolor="<%=bgctwo%>"align="center"><%=oRs2.Fields("NoShows").Value%></td>
						<td  bgcolor="<%=bgcone%>"align="center"><%=FormatNumber(cInt(oRs2.Fields("WinPct").Value) / 10000, 3, 0)%></td>
					</tr>
					<%
					oRs2.MoveNext
				Loop
			Else
				%>
				<tr>
					<td bgcolor="<%=bgcone%>" colspan="8"><i>No teams are in this division</td>
				</tr>
				<%
			End If
			oRs2.NextRecordSet
			%>
			<tr>
				<td bgcolor="#000000" colspan="8" align="right"><a href="viewleaguedivision.asp?league=<%=Server.URLEncode(strLeagueName)%>&conference=<%=Server.URLEncode(strConferenceName)%>&division=<%=Server.URLEncode(strDivisionName)%>">view division &raquo;</a></td>
			</tr>
			</table></td>
			<%
		oRs.MoveNext
	Loop
	IF intDivisionsShown mod 2 = 1 THen
		Response.Write "<td bgcolor=""#000000"">&nbsp;</td></tr>"
	End If
	%>
	</table>
	<%
End If

	
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>