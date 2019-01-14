<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Full Game list"

Dim strSQL, oConn, oRS, oRS2, oCmd
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
set oCmd = server.CreateObject("adodb.command")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim intGameID
intGameID = -1

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
If Not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Set oRS2 = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=3"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Full Game List") 

strSQL = "SELECT GameName, GameID, tbl_games.ForumID, GameAbbreviation, f1.ForumName, f2.ForumName AS DisputeForum "
strSQL = strSQL & " FROM tbl_games "
strSQL = strSQL & " INNER JOIN tbl_forums f1 ON tbl_games.ForumID = f1.ForumID "
strSQL = strSQL & " LEFT OUTER JOIN tbl_forums f2 ON tbl_games.DisputeForumID = f2.ForumID "
strSQL = strSQL & " WHERE GameID > 0 ORDER BY GameName ASC "
oRs.Open strSQL, oConn
bgc=bgctwo
if not (ors.EOF and ors.BOF) then
	%>
		<table border="0" cellspacing="0" cellpadding="0" ALIGN=CENTER BGCOLOR="#444444" width="97%">
		<TR><TD>
		<table border="0" cellspacing="1" cellpadding="2" WIDTH=100%>
		<TR BGCOLOR="#000000">
			<TH COLSPAN=5>TWL Games</TH>
		</TR>
		<TR BGCOLOR="<%=bgcone%>"><td colspan="5" align="center"><a href="addgame.asp">New Game</a></td></tr>
		<TR BGCOLOR="#000000">
			<TH>Game Name</TH>
			<TH WIDTH=75>Game Abbr</TH>
			<TH WIDTH=100>Forum name</TH>
			<TH WIDTH=100>Dispute Forum</TH>
			<TH WIDTH=50>Edit</TH>
		</TR>
		<%
		do while not ors.EOF
			%>
			<tr bgcolor=<%=bgc%>>
				<td><%=oRS.Fields("GameName").Value%></td>
				<td><%=oRs.Fields("GameAbbreviation").Value %></td>
				<td><%=oRs.Fields("ForumName").Value %></td>
				<td><%=oRs.Fields("DisputeForum").Value %></td>
				<td align="center"><a href="addgame.asp?IsEdit=True&Game=<%=Server.URLEncode(oRS.Fields("GameName").Value)%>">edit</a></td>
			</tr>
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.MoveNext
		loop
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
Set oRs = Nothing
%>

