<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Create Tournament"

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

If Not(bSysAdmin) Then
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
	intTeams = oRs.Fields("TeamsPerDiv").Value
End If
oRs.NextRecordSet

Dim intDivisionNum
Dim arrTeams(32, 32)
intDivisionID = Request.QueryString("DivisionID")
arrTeams (0, 0) = 0
arrTeams (0, 1) = "Bye"
arrTeams (0, 2) = ""
Dim intCounter
intCounter = 1
If Len(intDivisionID) > 0 Then
	strSQL = "SELECT TeamName, TeamTag, TMLinkID, d.DivisionID FROM lnk_t_m m INNER JOIN tbl_teams t ON t.TeamID = m.TeamID INNER JOIN tbl_tdivisions d ON d.DivisionID = m.DivisionID  AND d.TournamentID = m.TournamentID WHERE d.DivisionID = '" & intDivisionID & "' AND m.TournamentID = '" & intTournamentID & "' ORDER BY TeamName ASC "
	'Response.Write strSQL
	'response.end
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intDivisionNum = oRs.Fields("DivisionID").Value
		Do While Not(oRs.EOF)
			arrTeams (intCounter, 0) = oRs.Fields("TMLinkID").Value
			arrTeams (intCounter, 1) = oRs.Fields("TeamName").Value
			arrTeams (intCounter, 2) = oRs.Fields("TeamTag").Value

			intCounter = intCounter + 1
			oRs.MoveNext
		Loop
	End If
	oRs.NextRecordSet
End If

	
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("Edit " & Server.HTMLEncode(strTournamentName) & " Seedings")

intDivisionID = Request.QueryString("DivisionID")
If Len(intDivisionID) = 0 Then
	%>
	<table ALIGN=CENTER border=0 cellpadding=0 cellspacing=0 BGCOLOR="#444444">
	<TR><TD align="center">
	<form action="editSeedings.asp" method="get">
	<input type="hidden" name="Tournament" id="Tournament" value="<%=Server.HTMLEncode(strTournamentName)%>" />
	<table ALIGN=CENTER border=0 cellpadding=4 cellspacing=1 WIDTH=100%>
	<tr><th bgcolor="#000000">Choose a Division</th></tr>
			<tr>
				<td bgcolor="<%=bgctwo%>" align="center">
					<select name="DivisionID" id="DivisionID">
	<%
	strSQL = "SELECT DivisionName, DivisionID, TDivisionID FROM tbl_tdivisions WHERE TournamentID = '" & intTournamentID & "'"
	oRs.Open strSQL, oConn
	If Not(oRS.EOF AND oRs.BOF) THen
		Do While Not (oRs.EOF)
			%>
				<option value="<%=oRs.Fields("DivisionID").Value%>"><%=oRs.Fields("DivisionName").Value%></option>
			<%
			oRs.MoveNext
		Loop
	End If
	oRs.NextRecordset
	%>
			</select>
				</td>
			</tr>
		<tr>
			<td width="100%" align="center" colspan="2" bgcolor="#000000"><input type="Submit" name="submit" value="Change Division"></td>
		</tr>
	</form>
	</table></td></tr>
	</table>
	<%
Else
	' DIvisionID has been choosen
	%>
	<table ALIGN=CENTER border=0 cellpadding=0 cellspacing=0 BGCOLOR="#444444">
	<TR><TD align="center">
	<form action="savetournament.asp" method="post">
	<input type="hidden" name="SaveType" id="SaveType" value="EditSeedings" />
	<input type="hidden" name="TournamentID" id="TournamentID" value="<%=intTournamentID%>" />
	<input type="hidden" name="DivisionID" id="DivisionID" value="<%=intDivisionNum%>" />
	<table ALIGN=CENTER border=0 cellpadding=4 cellspacing=1 WIDTH=100%>
	<%
	Dim i, j
	For i = 1 To intTeams / 2
		%>
		<tr>
			<td bgcolor="<%=bgcone%>">Round <%=i%> Team 1:</td>
			<td bgcolor="<%=bgcone%>">
				<input type="hidden" name="seedorder" id="seedorder" value="<%=i-1%>" />
				<select name="Team1" id="Team1">
					<%
					For j = 0 to 32
						If Len(arrTeams(j, 0)) > 0 Then
							Response.Write "<option value=""" & arrTeams(j, 0) & """>" & Server.HTMLEncode( arrTeams(j, 1) & " - " & arrTeams(j, 2)) & "</option>" & vbCrLf
						End If
					Next
					%>
				</select>
			</td>
		</tr>
		<tr>
			<td bgcolor="<%=bgctwo%>">Round <%=i%> Team 2:</td>
			<td bgcolor="<%=bgcone%>">
				<select name="Team2" id="Team2">
					<%
					For j = 0 to 32
						If Len(arrTeams(j, 0)) > 0 Then
							Response.Write "<option value=""" & arrTeams(j, 0) & """>" & Server.HTMLEncode( arrTeams(j, 1) & " - " & arrTeams(j, 2)) & "</option>" & vbCrLf
						End If
					Next
					%>
				</select>
			</td>
		</tr>
		<%
	Next
	%>
	<tr>
		<td colspan="2" align="center" bgcolor="#000000"><input type="submit" value="Save Seeding Information" /></td>
	</tr>
	</table>
	</td></tr>
	</table>
	<%
End if

Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>