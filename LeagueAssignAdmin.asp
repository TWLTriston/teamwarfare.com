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
Call ContentStart("Assign League Admins")
%>
<form name="frmLeagueAdmins" id="frmLeagueAdmins" action="saveitem.asp" method="post">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center">
<tr><td>
<table border="0" cellspacing="1" cellpadding="4" width="400">
<tr>
	<th colspan="2" bgcolor="#000000">Current League Admins</th>
</tr>
<tr>
	<td align="right" bgcolor="<%=bgcone%>">Admin:</td>
	<td bgcolor="<%=bgcone%>">
		<select name="selLeagueAdminID" id="selLeagueAdminID">
		<option value="">Select an Admin</option>
		<%
		strSQL = "SELECT p.PlayerHandle, p.PlayerID, lnk.LeagueAdminID, l.LeagueName "
		strSQL = strSQL & " FROM lnk_league_admin lnk "
		strSQL = strSQL & " INNER JOIN tbl_players p ON lnk.PlayerID = p.PlayerID "
		strSQL = strSQL & " INNER JOIN tbl_leagues l ON lnk.LeagueID = l.LeagueID "
		strSQL = strSQL & " ORDER BY p.PlayerHandle ASC "
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRs.BOF) Then
			Do While Not(oRs.EOF)
				%>
				<option value="<%=oRs.Fields("LeagueAdminID").Value%>"><%=Server.HTMLENcode(oRs.Fields("PlayerHandle").Value) & " - " & Server.HTMLEncode(oRs.Fields("LeagueName").Value) %></option>
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
<input type="hidden" name="SaveType" id="SaveType" value="LeagueRemoveAdmin" />
</form>

<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center">
<tr><td>
<table border="0" cellspacing="1" cellpadding="4" width="400">
	<% if request("player") <> "" Then %>
		<form action=saveItem.asp method=post id=form4 name=form4>
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
		strsql="select LeagueName, LeagueID from tbl_leagues order by LeagueName"
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			do while not ors.EOF
				Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value) & " League</option>"
				ors.MoveNext
			loop
		end if
		ors.NextRecordSet
		%>
	<% Else %>
		<form action=leagueassignadmin.asp method=post id=form6 name=form6>
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2>Search for Player</TH>
		</TR>
		<tr bgcolor=<%=bgcone%> height=30><td>&nbsp;&nbsp;Player Name: </td><td><input type=text name=player class=bright style="width:200"></td></tr>
	<% End IF %>
	<tr bgcolor=<%=bgcone%>><td colspan=2 align=center><input type=submit value='Make it So' class=bright id=submit2 name=submit2>
		<input type=hidden name=SaveType value="LeagueAssignAdmin"></td>
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