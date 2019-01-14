<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: League History"

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
strLeagueName = Request.QueryString("League")

Dim intLeagueID
strSQL = "SELECT LeagueID FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
Else
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
oRs.NextRecordset
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("")
%>
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" width="97%">
<tr><td>
<table border="0" cellspacing="1" cellpadding="4" width="100%">
<tr>
	<th colspan="6" bgcolor="#000000"><%="<a href=""viewleague.asp?league=" & Server.URLEncode(strLeagueName) & """>" & strLeagueName & " League</a> Recent History"%></th>
<tr>
	<th bgcolor="#000000">Home Team</th>
	<th bgcolor="#000000">Visitor Team</th>
	<th bgcolor="#000000">Home Points</th>
	<th bgcolor="#000000">Visitor Points</th>
	<th bgcolor="#000000">Match Date</th>
	<th bgcolor="#000000">Map Info</th>
</tr>	
<%
Dim strWeek
strWeek = Request.Querystring("weekago")
if len(strWeek) = 0 Then
	strWeek = 0
End If
strSQL = "EXECUTE LeagueGetHistory @LeagueID = '" & intLeagueID & "', @XFactor='"  & strWeek  & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	Do While Not(oRs.EOF)
		%>
		<tr>
			<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(oRs.Fields("HomeTeamName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("HomeTeamName").Value & "")%></a></td>
			<td bgcolor="<%=bgctwo%>"><a href="viewteam.asp?team=<%=Server.URLEncode(oRs.Fields("VisitorTeamName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("VisitorTeamName").Value & "")%></a></td>
			<td bgcolor="<%=bgcone%>" align="center"><%=Server.HTMLEncode(oRs.Fields("HomeTeamPoints").Value & "")%></td>
			<td bgcolor="<%=bgctwo%>" align="center"><%=Server.HTMLEncode(oRs.Fields("VisitorTeamPoints").Value & "")%></td>
			<td bgcolor="<%=bgcone%>" align="center"><%=FormatDateTime(oRs.Fields("MatchDate").Value, 2)%></td>
			<td bgcolor="<%=bgcone%>" align="center"><%
			Dim i
			For i = 1 to 5
				If Len(oRs.Fields("Map" & i).Value) > 0 Then
					Response.Write oRs.Fields("Map" & i).Value & " (" & oRs.Fields("Map" & i & "HomeScore").Value & "-" & oRs.Fields("Map" & i & "VisitorScore").Value & ")<br />"
				End If
			Next
			%></td>
		
		</tr>
		<%
		oRs.MoveNext
	Loop
End If
oRs.NextRecordSet
%>
</table>
</td></tr>
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