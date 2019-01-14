<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: League Admins"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strLeagueName

if not(bSysAdmin) Then
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
Call ContentStart("Assign Tournament Admins")
%>
<form name="frmTournamentAdmins" id="frmTournamentAdmins" action="savetournament.asp" method="post">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center">
<tr><td>
<table border="0" cellspacing="1" cellpadding="4" width="400">
<tr>
	<th colspan="2" bgcolor="#000000">Current Tournament Admins</th>
</tr>
<tr>
	<td align="right" bgcolor="<%=bgcone%>">Admin:</td>
	<td bgcolor="<%=bgcone%>">
		<select name="selTournamentAdminID" id="selTournamentAdminID">
		<option value="">Select an Admin</option>
		<%
		strSQL = "SELECT p.PlayerHandle, p.PlayerID, lnk.MALinkID, t.TournamentName "
		strSQL = strSQL & " FROM lnk_m_a lnk "
		strSQL = strSQL & " INNER JOIN tbl_players p ON lnk.PlayerID = p.PlayerID "
		strSQL = strSQL & " INNER JOIN tbl_tournaments t ON lnk.TournamentID = t.TournamentID "
		strSQL = strSQL & " WHERE t.Active = 1 ORDER BY p.PlayerHandle ASC "
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRs.BOF) Then
			Do While Not(oRs.EOF)
				%>
				<option value="<%=oRs.Fields("MALinkID").Value%>"><%=Server.HTMLENcode(oRs.Fields("PlayerHandle").Value) & " - " & Server.HTMLEncode(oRs.Fields("TournamentName").Value) %></option>
				<%
				oRs.MoveNext
			Loop
		End If
		oRs.NextRecordSet
		%>
		</select>
	</td>
</tr>
<tr>
	<td colspan="2" bgcolor="#000000" align="center"><input type="submit" value="Remove Admin"></td>
</tr>
</table>
</td></tr></table>
<input type="hidden" name="SaveType" id="SaveType" value="TournamentRemoveAdmin" />
</form>

<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center">
<tr><td>
<table border="0" cellspacing="1" cellpadding="4" width="400">
	<% if request("player") <> "" Then %>
		<form action=savetournament.asp method=post id=form4 name=form4>
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2>Select Player to Admin</TH>
		</TR>
		<tr bgcolor=<%=bgcone%>><td>&nbsp;&nbsp;Player: </td><td><select name=PlayerID class=bright style="width:200">
		<%
		strsql="select playerhandle, tbl_players.playerid from tbl_players WHERE playerHandle like '%" & CheckString(SearchString(request("player"))) & "%' order by playerhandle"
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			do while not ors.EOF
				Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value) & "</option>"
				ors.MoveNext 
			loop
		end if
		ors.NextRecordSet 
		%>
		</td></tr>
		<tr bgcolor=<%=bgctwo%> height=30><td>&nbsp;&nbsp;League</td><td><select name=LeagueID class=bright style="width:200">
		<%
		strsql="select TournamentName, TournamentID from tbl_Tournaments WHERE Active = 1 order by TournamentName"
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			do while not ors.EOF
				Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value) & " Tournament</option>"
				ors.MoveNext
			loop
		end if
		ors.NextRecordSet
		%>
	<% Else %>
		<form action=assignadmin.asp method=post id=form6 name=form6>
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2>Search for Player</TH>
		</TR>
		<tr bgcolor=<%=bgcone%> height=30><td>&nbsp;&nbsp;Player Name: </td><td><input type=text name=player class=bright style="width:200"></td></tr>
	<% End IF %>
	<tr bgcolor=<%=bgcone%>><td colspan=2 align=center><input type=submit value='Make it So' class=bright id=submit2 name=submit2>
		<input type=hidden name=SaveType value="TournamentAssignAdmin"></td>
	</tr>
	</form>
	</table>
</td></tr></table>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>