<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: RSS Feeds"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim intGameID
intGameID = -1

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("")  %>

<table class="cssBordered" align="center" width="60%">
<tr>
	<th>Link</th>
	<th>Available Feeds</th>
</tr>
<tr>
	<td bgcolor="<%=bgctwo%>" align="center"><a href="feed.asp">RSS</a></td>
	<td bgcolor="<%=bgctwo%>">News &amp; Announcements</a></td>
</tr>
<%
Dim intThisGame, arrGameIDs(50), iGameCounter
intThisGame = -1
iGameCounter = 0
strSQL = "SELECT DISTINCT GameID, GameName FROM tbl_games WHERE GameID > 0 ORDER BY GameName ASC "
oRs.Open strSQL, oConn
Do While Not(oRs.EOF)
	If intThisGame <> oRs.Fields("GameID").Value Then
		intThisGame = oRs.Fields("GameID").Value
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
		%>
		<tr>
			<td bgcolor="<%=bgc%>" align="center"><a href="feed.asp?Game=<%=Server.URLEncode(oRs.Fields("GameName").Value & "")%>">RSS</a></td>
			<td bgcolor="<%=bgc%>" ><%=Server.HTMLEncode(oRs.Fields("GameName").Value & "")%></a></td>
		</tr>
		<%
	End If
	oRs.MoveNext
Loop
oRs.NextRecordSet
%>
</table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

