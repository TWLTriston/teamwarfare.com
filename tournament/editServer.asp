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
strDivisionName = ""
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("" & Server.HTMLEncode(strTournamentName) & " Server Information")
%>
<form name="frmEditServerInfo" id="frmEditServerInfo" action="savetournament.asp" method="post">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
<tr>
	<td>
<%
strSQL = "SELECT Team1ID, Team2ID, Round, RoundsID, BracketBlurb, 'Team1Name' = t1.TeamName, 'Team2Name' = t2.TeamName, ServerName, ServerIP, ServerJoinPassword, ServerRConPassword, DivisionName, MatchTime "
strSQL = strSQL & " FROM tbl_rounds r "
strSQL = strSQL & " INNER JOIN lnk_t_m m1 ON m1.TMLinkID = Team1ID "
strSQL = strSQL & " INNER JOIN lnk_t_m m2 ON m2.TMLinkID = Team2ID "
strSQL = strSQL & " INNER JOIN tbl_teams t1 ON t1.TeamID = m1.TeamID "
strSQL = strSQL & " INNER JOIN tbl_teams t2 ON t2.TeamID = m2.TeamID "
strSQL = strSQL & " LEFT OUTER JOIN tbl_tdivisions td ON r.DivisionID = td.DivisionID AND td.TournamentID = r.TournamentID "
strSQL = strSQL & " WHERE r.RoundsID = '" & Request.QueryString("RoundsID") & "' ORDER BY r.DivisionID, Round "
oRs.Open strSQL, oConn
%>
		<table border="0" cellspacing="01" cellpadding="4" width="100%">
		<tr><th colspan="2" bgcolor="#000000">Server Information</th></tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Round</td>
			<td bgcolor="<%=bgctwo%>"><%=oRs.FIelds("Round").Value%></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Home Team</td>
			<td bgcolor="<%=bgctwo%>"><%=oRs.FIelds("Team1Name").Value%></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Visiting Team</td>
			<td bgcolor="<%=bgctwo%>"><%=oRs.FIelds("Team2Name").Value%></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Server Name</td>
			<td bgcolor="<%=bgctwo%>"><input type="text" name="txtServerName" id="txtServerName" value="<%=Server.HTMLEncode(oRs.Fields("ServerName").Value & "")%>" /></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Server IP</td>
			<td bgcolor="<%=bgctwo%>"><input type="text" name="txtServerIP" id="txtServerIP" value="<%=Server.HTMLEncode(oRs.Fields("ServerIP").Value & "")%>" /></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Join Password</td>
			<td bgcolor="<%=bgctwo%>"><input type="text" name="txtServerJoinPassword" id="txtServerJoinPassword" value="<%=Server.HTMLEncode(oRs.Fields("ServerJoinPassword").Value & "")%>" /></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Rcon Password</td>
			<td bgcolor="<%=bgctwo%>"><input type="text" name="txtServerRConPassword" id="txtServerRConPassword" value="<%=Server.HTMLEncode(oRs.Fields("ServerRConPassword").Value & "")%>" /></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Match Time (MM/DD/YYYY HH:MM:SS) Hours in 24hr time</td>
			<td bgcolor="<%=bgctwo%>"><input type="text" name="txtMatchTime" id="txtMatchTime" value="<%=Server.HTMLEncode(oRs.Fields("MatchTime").Value & "")%>" /></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Bracket Blurb</td>
			<td bgcolor="<%=bgctwo%>"><input type="text" name="txtBracketBlurb" id="txtBracketBlurb" value="<%=Server.HTMLEncode(oRs.Fields("BracketBlurb").Value & "")%>" /></td>
		</tr>
		<tr>
			<td colspan="2" align="center" bgcolor="#000000"><input type="submit" value="Save Server Information" /></td>
		</tr>
		</tr>
		</table>
	</td>
</tr>
</table>
<input type="hidden" name="SaveType" id="SaveType" value="ServerInfo" />
<input type="hidden" name="RoundsID" id="RoundsID" value="<%=Request.QueryString("RoundSID")%>" />
<input type="hidden" name="Tournament" id="Tournament" value="<%=Server.HTMLEncode(Request.QueryString("Tournament") & "")%>" />

</form>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>