<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Player Ladder Listing"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentSTart("Active Player  Ladders") %>
	<table border="0" cellspacing="0" cellpadding="0" ALIGN=CENTER BGCOLOR="#444444">
	<TR><TD>
	<table border="0" cellspacing="1" cellpadding="2">
	<TR BGCOLOR="#000000">
		<TH COLSPAN=<% If bSysAdmin Then Response.write "3" Else Response.Write "2" End If %>>Active Ladders</TH>
	</TR>
	<TR bgcolor=#000000><TH>Ladder Name</th><th>Players</th>
	<% If bSysAdmin Then %>
		<TH>Edit</TH>
	<% End If %>
	</tr>
	<%
	strSQL="select PlayerLadderName, PlayerLadderID, ActivePlayers = (SELECT COUNT(PlayerID) FROM lnk_p_pl lnk WHERE IsActive = 1 AND lnk.PlayerLadderID = l.PlayerLAdderID) from tbl_playerladders l where active=1 order by PlayerLadderName"
	oRs.Open strSQL, oConn
	bgc=bgctwo
	if not (ors.EOF and ors.BOF) then
		do while not ors.EOF
			%>
			<tr bgcolor=<%=bgc%> >
				<td align=center><a href=viewplayerladder.asp?ladder=<%=server.urlencode(ors.Fields("PlayerLadderName").Value) %> ><% =Server.HTMLEncode(ors.Fields("PlayerLadderName").Value) %></a></td>
				<td align=center><%=oRS.Fields("ActivePlayers").Value %></td>
				<% If bSysAdmin Then %>
					<TD align=center><a href="addplayerladder.asp?isedit=true&name=<%=ors.Fields ("PlayerLadderName").Value %>">Edit</A></TD>
				<% End IF %>
				</tr>
			<%
			ors.MoveNext
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
		loop
	else
		%>
		<tr><td>No ladders found.</td></tr>
		<%
	end if
%>
</table>
</TD></TR>
</TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>
