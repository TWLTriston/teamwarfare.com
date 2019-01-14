<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: League Assign Divisions"

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
Dim bAnyLeagueAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
bAnyLeagueAdmin = IsAnyLeagueAdmin()

Dim strLeagueName, intLeagueID, intDivisions, intConferences, intLinkID
if not(bSysAdmin or bAnyLeagueAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("")
%>
<table border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#444444">
<tr>
	<td>
	<table border="0" cellspacing="1" cellpadding="4" width="400" align="center">
	<tr>
		<th bgcolor="#000000">Select a League</th>
	</tr>
<%
If bSysAdmin Then
	strSQL = "SELECT LeagueName, LeagueID FROM tbl_leagues WHERE LeagueActive = 1 "
Else
	strSQL = "SELECT LeagueName, l.LeagueID FROM tbl_leagues l INNER JOIN lnk_league_admin lnk ON lnk.LeagueID = l.LeagueID WHERE lnk.PlayerID='" & Session("PlayerID") & "' AND LeagueActive = 1 "
End If
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	bgc = bgcone
	Do While Not(oRs.EOF)
		%>
		<tr>
			<td bgcolor="<%=bgc%>"><a href="leagueassign.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value)%>"><%=Server.HTMLEncode(oRS.FIelds("LeagueName").value)%> League</a></td>
		</tr>
		<%
		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		oRs.MoveNext
	loop
End if
oRs.NextRecordSet
%>
	</table>
	</td>
</tr></table>
<%
Call ContentEnd()

strLeagueName = trim(Request.Querystring("League"))
strSQL = "SELECT LeagueID FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "' AND LeagueActive = 1"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
End If
oRs.NextRecordSet
If Len(intLeagueID) > 0 AND (bSYsAdmin Or IsLeagueAdminById(intLeagueID)) Then
	Call ContentStart(strLeagueName & " League Pending Joins")
	%>
	<table border="0" cellspacing="0" cellpadding="0" width="97%" align="center" bgcolor="#444444">
	<tr><td>
	<table border="0" cellspacing="1" cellpadding="4" width="100%" align="center">
	<tr>
		<th bgcolor="#000000">Team</th>
		<th bgcolor="#000000">Date</th>
		<th bgcolor="#000000" width="150">Conference</th>
		<th bgcolor="#000000" width="150">Division</th>
		<th bgcolor="#000000">Save</th>
		<th bgcolor="#000000">Kick</th>
	</tr>
	<script language="javascript" type="text/javascript">
	function ConfirmKick(objForm) {
		if (confirm("Are you sure you wish to kick this team from the league?")) {
			objForm.submit();
		}
	}
	var arrConferences = new Array(); // new Array(intConferenceID, strConferenceName);
	var arrDivisions = new Array(); // new Array(intConferenceID, intDivsionID, strDivisionName);

	function ChangeOptions(objConferences, objDivisions) 
	{
		objDivisions.length = 0;
		var intOptions = 0;
		for (var i=1;i<arrDivisions.length;i++) {
			if (arrDivisions[i][0] == objConferences.options[objConferences.selectedIndex].value) {
				objDivisions.options[intOptions] = new Option(arrDivisions[i][2], arrDivisions[i][1]);
				intOptions++;
			}
		}
	}   
   
	<%
	strSQL = "SELECT ConferenceName, c.LeagueConferenceID FROM tbl_league_conferences c WHERE c.LeagueID = '" & intLeagueID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intConferences = 0
		Do While Not(oRs.EOF)
			intConferences = intConferences + 1
			%>
			arrConferences[<%=intConferences%>] = new Array("<%=oRs.Fields("LeagueConferenceID").Value%>", "<%=Server.HTMLEncode(oRS.Fields("ConferenceName").value)%> Conference");
			<%
			oRs.MoveNext
		Loop
		%>
		var intConferences = <%=intConferences%>;
		<%
	End If
	oRs.NextRecordSet
	
	strSQL = "SELECT d.LeagueConferenceID, LeagueDivisionID, DivisionName FROM tbl_league_divisions d WHERE d.LeagueID = '" & intLeagueID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intDivisions = 0
		Do While Not(oRs.EOF)
			intDivisions = intDivisions + 1
			%>
			arrDivisions[<%=intDivisions%>] = new Array("<%=oRs.Fields("LeagueConferenceID").Value%>", <%=oRs.Fields("LeagueDivisionID").Value%>, "<%=Server.HTMLEncode(oRS.Fields("DivisionName").value)%> Division");
			<%
			oRs.MoveNext
		Loop
		%>
		var intDivisions = <%=intDivisions%>;
		<%
	End If
	oRs.NextRecordSet
	%>
	</script>
	<%
	strSQL = "SELECT t.TeamName, l.JoinDate, l.LeagueConferenceID, l.lnkLeagueTeamID, RosterCount = (SELECT COUNT(PlayerID) FROM lnk_league_team_player lltp WHERE lltp.lnkLeagueTeamID = l.lnkLeagueTeamId) FROM lnk_league_team l "
	strSQL = strSQL & " INNER JOIN tbl_teams t ON l.TeamID = t.TeamID WHERE LeagueDivisionID = 0 AND LeagueID='" & intLeagueID & "' AND l.Active = 1 ORDER BY JoinDate DESC"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		Do While Not(oRs.EOF)
			intLinkID = oRs.Fields("lnkLeagueTeamID").Value			
			%>
			<tr>
			<form name="frmAssignDivision<%=intLinkID%>" id="frmAssignDivision<%=intLinkID%>" action="saveitem.asp" method="post">
			<input type="hidden" name="SaveType" id="SaveType" value="LeagueAssignDivision">
			<input type="hidden" name="lnkLeagueTeamID" id="lnkLeagueTeamID" value="<%=intLinkID%>">
			<input type="hidden" name="LeagueID" id="LeagueID" value="<%=intLeagueID%>">
				<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(oRs.Fields("TeamName").Value)%>"><%=Server.HTMLEncode(oRs.Fields("TeamName").Value)%> (<%=oRS.Fields("RosterCount").Value%>)</a></td>
				<td align="center" bgcolor="<%=bgctwo%>"><%=FormatDateTime(oRs.Fields("JoinDate").Value, 2)%></td>
				<td align="center" bgcolor="<%=bgcone%>"><select name="selConferenceID" id="selConferenceID" onchange="ChangeOptions(frmAssignDivision<%=intLinkID%>.selConferenceID, frmAssignDivision<%=intLinkID%>.selDivisionID)">
						<%
						strSQL = "SELECT ConferenceName, c.LeagueConferenceID FROM tbl_league_conferences c WHERE c.LeagueID = '" & intLeagueID & "' ORDER BY conferencename asc"
						oRs2.Open strSQL, oConn
						If Not(oRs2.EOF AND oRs2.BOF) Then
							Do While Not(oRs2.EOF)
								Response.Write "<option value=""" & oRs2.Fields("LeagueConferenceID").Value & """ "
								If oRS.Fields("LeagueConferenceID").Value = oRs2.Fields("leagueConferenceID").Value Then
									Response.Write " selected=""selected"" "
								End If
								Response.Write ">" & Server.HTMLEncode(oRs2.Fields("ConferenceName").Value) & " Conference</option>" & vbCrlF
								oRs2.MoveNext
							Loop
						End If
						oRs2.NextRecordSet
						%>
					</select></td>
				<td align="center" bgcolor="<%=bgcone%>"><select name="selDivisionID" id="selDivisionID">
						<%
						strSQL = "SELECT DivisionName, d.LeagueDivisionID FROM tbl_league_divisions d WHERE d.LeagueConferenceID = '" & oRs.Fields("LeagueConferenceID").Value & "' ORDER BY DivisionName ASC"
						oRs2.Open strSQL, oConn
						If Not(oRs2.EOF AND oRs2.BOF) Then
							Do While Not(oRs2.EOF)
								Response.Write "<option value=""" & oRs2.Fields("LeagueDivisionID").Value & """ "
								Response.Write ">" & Server.HTMLEncode(oRs2.Fields("DivisionName").Value) & " Division</option>" & vbCrlF
								oRs2.MoveNext
							Loop
						End If
						oRs2.NextRecordSet
						%>
					</select></td>
				<td align="center" bgcolor="<%=bgctwo%>"><input type="submit" value="Assign Division"></td>
			</form>
			<form name="frmDecline<%=intLinkID%>" id="frmDecline<%=intLinkID%>" action="saveitem.asp" method="post">
			<input type="hidden" name="SaveType" id="SaveType" value="LeagueDecline">
			<input type="hidden" name="lnkLeagueTeamID" id="lnkLeagueTeamID" value="<%=intLinkID%>">
			<input type="hidden" name="LeagueID" id="LeagueID" value="<%=intLeagueID%>">
				<td align="center" bgcolor="<%=bgctwo%>"><input type="button" value="Kick" onclick="javascript:ConfirmKick(this.form);"></td>
			</form>
			</tr>	
			<%
			oRs.MoveNext
		Loop
	End If
	oRs.NextRecordSet
	%>
	</table></td></tr>
	</table>
	<%
	Call ContentEnd()
End If
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>