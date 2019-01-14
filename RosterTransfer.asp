<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Roster Transfer"

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo, bgc

bgcone = Application("bgcone")
bgctwo = Application("bgctwo")

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin, bAnyLeagueAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
bAnyLeagueAdmin = IsAnyLeagueAdmin()
Dim strLeagueName, intLeagueID, intDivisions, intConferences, intLinkID

if Not(bSysAdmin) Then
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

If Len(Request.QueryString("S")) > 0 Then
	%>
	Transfer Successful!<br /><br /><br />
	<%
End If

strSQL = "SELECT LeagueName, LeagueID FROM tbl_leagues WHERE LeagueActive = 1 ORDER BY LeagueName ASC "
oRs.Open strSQL, oConn, 3, 3
%>
<table border="0" cellspacing="0" cellpadding="0" align="center" class="cssBordered">
<form name="frmLeagueToLeague" id="frmLeagueToLeague" method="post" action="dbRosterTransfer.asp">
<input type="hidden" name="SaveType" id="SaveType" value="LeagueToLeague" />
<tr>
	<th colspan="2">League to League Transfer</th>
</tr>
<tr>
	<td bgcolor="<%=bgcone%>">Original League:</td>
	<td bgcolor="<%=bgcone%>"><select name="selFromLeagueID" id="selFromLeagueID">
		<option value="">Select a league</option>
		<%
		oRs.MoveFirst
		Do While Not(oRs.EOF)
			Response.Write "<option value=""" & oRs.Fields("LeagueID").Value & """>" & Server.HTMLEncode(oRs.Fields("LeagueName").Value & "") & "</option>" & vbCrLf
			oRs.MoveNext
		Loop
		%>
		</select>
	</td>
</tr>
<tr>
	<td bgcolor="<%=bgcone%>">Target League:</td>
	<td bgcolor="<%=bgcone%>"><select name="selToLeagueID" id="selToLeagueID">
		<option value="">Select a league</option>
		<%
		oRs.MoveFirst
		Do While Not(oRs.EOF)
			Response.Write "<option value=""" & oRs.Fields("LeagueID").Value & """>" & Server.HTMLEncode(oRs.Fields("LeagueName").Value & "") & "</option>" & vbCrLf
			oRs.MoveNext
		Loop
		%>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" align="center"><input type="submit" value="Perform League to League Transfer" /></td>
</tr>
</form>
</table>

<br /><br />
<table border="0" cellspacing="0" cellpadding="0" align="center" class="cssBordered">
<form name="frmLeagueToTournament" id="frmLeagueToTournament" method="post" action="dbRosterTransfer.asp">
<input type="hidden" name="SaveType" id="SaveType" value="LeagueToTournament" />
<tr>
	<th colspan="2">League to Tournament Transfer</th>
</tr>
<tr>
	<td bgcolor="<%=bgcone%>">Original League:</td>
	<td bgcolor="<%=bgcone%>"><select name="selFromLeagueID" id="selFromLeagueID">
		<option value="">Select a league</option>
		<%
		oRs.MoveFirst
		Do While Not(oRs.EOF)
			Response.Write "<option value=""" & oRs.Fields("LeagueID").Value & """>" & Server.HTMLEncode(oRs.Fields("LeagueName").Value & "") & "</option>" & vbCrLf
			oRs.MoveNext
		Loop
		%>
		</select>
	</td>
</tr>
<%
oRs.NextRecordSet

strSQL = "SELECT TournamentName, TournamentID FROM tbl_tournaments WHERE Active = 1 ORDER BY TournamentName ASC "
oRs.Open strSQL, oConn, 3, 3
%>
<tr>
	<td bgcolor="<%=bgcone%>">Target Tournament:</td>
	<td bgcolor="<%=bgcone%>"><select name="selToTournamentID" id="selToTournamentID">
		<option value="">Select a tournament</option>
		<%
		Do While Not(oRs.EOF)
			Response.Write "<option value=""" & oRs.Fields("TournamentID").Value & """>" & Server.HTMLEncode(oRs.Fields("TournamentName").Value & "") & "</option>" & vbCrLf
			oRs.MoveNext
		Loop
		%>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" align="center"><input type="submit" value="Perform League to Tournament Transfer" /></td>
</tr>
</form>
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