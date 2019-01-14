<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: League Admin"

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

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
<table border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#444444" width="97%">
<tr>
	<td>
	<table border="0" cellspacing="1" cellpadding="4" width="100%" align="center">
	<% If bSysAdmin Then %>
	<tr>
		<th colspan="7" bgcolor="#000000"><a href="leagueadd.asp">Add League</a></th>
	</tr>
	<% End If %>
	<tr>
		<th colspan="7" bgcolor="#000000">League Administration</th>
	</tr>
	<tr>
		<th bgcolor="#000000">Name</th>
		<th bgcolor="#000000">Pending Joins</th>
		<th bgcolor="#000000">Create Matches</th>
		<th bgcolor="#000000">Match Matrix</th>
		<th bgcolor="#000000">Map List</th>
		<% If bSysAdmin Then %>
		<th bgcolor="#000000">Divisions</th>
		<th bgcolor="#000000">Edit</th>
		<% End If %>
	</tr>
	
<%
If bSysAdmin Then
	strSQL = "SELECT LeagueName, LeagueID FROM tbl_leagues WHERE NOT(LeagueActive = 0 AND LeagueLocked = 1) ORDER BY LeagueName ASC" ' ORDER BY LeagueName ASC "
Else
	strSQL = "SELECT DISTINCT LeagueName, l.LeagueID FROM tbl_leagues l INNER JOIN lnk_league_admin lnk ON lnk.LeagueID = l.LeagueID WHERE lnk.PlayerID='" & Session("PlayerID") & "' AND NOT(LeagueActive = 0 AND LeagueLocked = 1) ORDER BY LeagueName ASC" ' AND LeagueActive = 1 "
End If
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	bgc = bgcone
	Do While Not(oRs.EOF)
		%>
		<tr>
			<td bgcolor="<%=bgc%>"><a href="viewleague.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value)%>"><%=Server.HTMLEncode(oRs.Fields("LeagueName").Value)%></a></td>
			<td bgcolor="<%=bgc%>"><a href="leagueassign.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value)%>">Pending Joins</a></td>
			<td bgcolor="<%=bgc%>"><a href="leaguematches.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value)%>">Create Matches</a></td>
			<td bgcolor="<%=bgc%>"><a href="viewleaguematrix.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value)%>">View Matrix</a></td>
			<td bgcolor="<%=bgc%>"><a href="leaguemaplist.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value)%>">Edit Map List</a></td>
			<% If bSysAdmin THen %>
			<td bgcolor="<%=bgc%>"><a href="leaguedivisionconfig.asp?intLeagueid=<%=oRs.Fields("LeagueID").Value%>">Edit Divisions</a></td>
			<td bgcolor="<%=bgc%>"><a href="leagueadd.asp?isedit=true&league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value)%>">Edit League</a></td>
			<% End If %>
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
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>