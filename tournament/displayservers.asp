<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Edit Server Information"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim X, tournamentid

If Not(bSysAdmin Or IsTournamentAdmin(Request.QueryString("Tournament"))) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim strTournamentName, intTournamentID, intDivisionID, intTeams
strTournamentName = Request.QueryString("tournament")
strSQL = "SELECT TournamentID, TeamsPerDiv FROM tbl_tournaments WHERE TournamentName = '" & CheckSTring(strTournamentName) & "'"
oRs.Open strSQL, oConn
If Not(ORs.Eof and Ors.BOF) Then
	intTournamentID = oRs.Fields("TournamentID").Value
End If
oRs.NextRecordSet

Dim strDivisionName 
intDivisionID = -1
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("" & Server.HTMLEncode(strTournamentName) & " Server Information")
%>
<%
strSQL = "SELECT Team1ID, Team2ID, Round, BracketBlurb, RoundsID, 'Team1Name' = t1.TeamName, 'Team2Name' = t2.TeamName, ServerName, ServerIP, ServerJoinPassword, ServerRConPassword, DivisionName, MatchTime, R.DivisionID "
strSQL = strSQL & " FROM tbl_rounds r "
strSQL = strSQL & " INNER JOIN lnk_t_m m1 ON m1.TMLinkID = Team1ID "
strSQL = strSQL & " INNER JOIN lnk_t_m m2 ON m2.TMLinkID = Team2ID "
strSQL = strSQL & " INNER JOIN tbl_teams t1 ON t1.TeamID = m1.TeamID "
strSQL = strSQL & " INNER JOIN tbl_teams t2 ON t2.TeamID = m2.TeamID "
strSQL = strSQL & " LEFT OUTER JOIN tbl_tdivisions td ON r.DivisionID = td.DivisionID AND td.TournamentID = r.TournamentID "
strSQL = strSQL & " WHERE r.TournamentID = '" & intTournamentID & "' ORDER BY r.DivisionID, Round "
'Response.Write strSQL
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	Do While Not(oRs.EOF)
		If intDivisionID <> oRs.Fields("DivisionID").Value Then
			If intDivisionID <> -1 Then
				%>
				</td></tr></table>
				</td></tr></table><br />
				<br />
				
				<%
			End If
			intDivisionID = oRs.Fields("DivisionID").Value 
			strDivisionName = oRs.Fields("DivisionName").Value 
			If intDivisionID = 0 Then
				strDivisionName = "Finals"
			End If
			%>
			<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" width="97%">
			<tr>
				<td>
					<table border="0" cellspacing="01" cellpadding="4" width="100%">
					<tr><th colspan="8" bgcolor="#000000"><%=strDivisionName%></th></tr>
					<tr>
						<th bgcolor="#000000" width="40">Round</th>
						<th bgcolor="#000000" width="200">Home Team</th>
						<th bgcolor="#000000" width="200">Visiting Team</th>
						<th bgcolor="#000000" width="200">Server</th>
						<th bgcolor="#000000" width="100">Join / RCon</th>
						<th bgcolor="#000000" width="75">Time</th>
						<th bgcolor="#000000" width="100">BracketBlurb</th>
						<th bgcolor="#000000">Edit</th>
					</tr>
			<%
		End If
		%>
		<tr>
			<td bgcolor="<%=bgcone%>" align="center"><%=oRs.Fields("Round").Value%></td>
			<td bgcolor="<%=bgctwo%>"><%=oRs.Fields("Team1Name").Value%></td>
			<td bgcolor="<%=bgcone%>"><%=oRs.Fields("Team2Name").Value%></td>
			<td bgcolor="<%=bgctwo%>"><%=oRs.Fields("ServerName").Value & " - " & oRs.Fields("ServerIP").Value%></td>
			<td bgcolor="<%=bgcone%>"><%=oRs.Fields("ServerJoinPassword").Value & " / " & oRs.Fields("ServerRconPassword").Value%></td>
			<td bgcolor="<%=bgcone%>"><%
				If Not(IsNull(oRs.Fields("MatchTime").Value)) Then
					Response.Write FormatDateTime(oRs.Fields("MatchTime").Value, 0)
				Else
					Response.Write "&nbsp;"
				End If
			%></td>
			<td bgcolor="<%=bgcone%>"><%=oRs.Fields("BracketBlurb").Value%></td>
			<td bgcolor="<%=bgcone%>"><a href="editServer.asp?Tournament=<%=Server.URLEncode(strTournamentName)%>&RoundsID=<%=oRs.Fields("RoundsID").Value%>">edit</a></td>
		</tr>
		<%
		oRs.MoveNext
	Loop
	'I moved the section just below this inside the IF statement to fix a borked table for tourneys with no server assignments.  Temp fix, clean as needed. - Will
%>
	</table>
	</td>
</tr>
</table>
<% 
End If
oRs.NextRecordSet

Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>