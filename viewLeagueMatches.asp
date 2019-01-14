<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("League"), """", "&quot;") 

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strLeagueName, intLeagueID
strLeagueName = Request.QueryString("League")
If Len(Trim(strLeagueName)) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If

strSQL = "SELECT LeagueID, LeagueName FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "' AND LeagueActive = 1"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
	strLeagueName = oRs.Fields("LeagueName").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
oRs.NextRecordSet

Dim dtmDate, intTimeZoneDifference, strDate, strTime
intTimeZoneDifference = 0

Dim strDateMask, bln24HourTime, blnVerticalBars, strColumnColor1, strColumnColor2
strDateMask = "MM-DD-YYYY"
bln24HourTime = False

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("")
%>

<table BORDER="0" cellspacing="0" cellpadding="0" width="760">
<tr>
	<td CLASS="pageheader"><%=strLeagueName%> League Upcoming Matches</td>
</tr>
<tr>
	<td>&nbsp;&nbsp;<a href="viewleague.asp?league=<%=Server.URLEncode(strLeagueName)%>">View League Standings</a></td>
</tr>
</table>
<br /><br />
<%
strSQL = "EXECUTE LeagueGetMatches @LeagueID = '" & intLeagueID & "'"
If Len(Request.QueryString("X")) > 0 AND IsNumeric(Request.QueryString("X")) Then
	strSQL = strSQL & ", @XFactor = '-" & Request.QueryString("X") & "'"
End If
oRs.Open strSQL, oConn
If (oRs.State = 1) Then 
	If Not(oRs.EOF AND oRs.BOF) Then
		%>
		<table border="0" cellspacing="0" cellpadding="0" width="97%" align="center" bgcolor="#444444">
		<tr>
			<td>
			<table border="0" cellspacing="1" cellpadding="4" width="100%">
			<tr>
				<th bgcolor="#000000">Home Team</th>
				<th bgcolor="#000000">Visitor Team</th>
				<th bgcolor="#000000">Predicted Outcome</th>
				<th bgcolor="#000000">Rants</th>
				<th bgcolor="#000000">Last Rant</th>
				<th bgcolor="#000000"> </th>
			</tr>
		<%
		Do While Not oRs.EOF
			%>
			<tr>
				<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(oRs.Fields("HomeTeamName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("HomeTeamName").Value & "")%></a></td>
				<td bgcolor="<%=bgctwo%>"><a href="viewteam.asp?team=<%=Server.URLEncode(oRs.Fields("VisitorTeamName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("VisitorTeamName").Value & "")%></a></td>
				<%
				If oRs.Fields("HomeVotes").Value > oRs.Fields("VisitorVotes").Value Then
					%>
					<td align="center" bgcolor="<%=bgcone%>"><b>Home</b> (<%=oRs.Fields("HomeVotes").Value%> / <%=oRs.Fields("VisitorVotes").Value%>)</td>
					<%
				ElseIf oRs.Fields("HomeVotes").Value < oRs.Fields("VisitorVotes").Value Then
					%>
					<td align="center" bgcolor="<%=bgcone%>"><b>Visitor</b> (<%=oRs.Fields("HomeVotes").Value%> / <%=oRs.Fields("VisitorVotes").Value%>)</td>
					<%
				Else
					%>
					<td align="center" bgcolor="<%=bgcone%>"><b>Unknown</b> (<%=oRs.Fields("HomeVotes").Value%> / <%=oRs.Fields("VisitorVotes").Value%>)</td>
					<%
				End If
				%>
				<td align="center" bgcolor="<%=bgctwo%>"><%=oRs.Fields("Rants").Value%></td>
				<%
				If Not(IsNull(ors.fields("LastRantTime").value)) Then 
					Call FixDate(ors.fields("LastRantTime").value, intTimeZoneDifference, strDate, strTime, strDateMask, bln24HourTime)
					Response.Write "<TD BGCOLOR=""" & bgcone & """ align=right NOWRAP><span class=""smalldate"">" & strDate & "</span>"
					Response.Write "&nbsp;<span class=""smalltime"">" & strTime & "</span><BR><span class=""note"">by <B>" & Server.HTMLEncode (ors.fields("LastRanterName").value) & "</span></td>"
				Else
					Response.Write "<TD BGCOLOR=""" & bgcone & """ class=""note"" ALIGN=CENTER>Never</TD>"
				End If
				%>
				<td align="center" bgcolor="<%=bgctwo%>"><a href="viewleaguematch.asp?League=<%=Server.URLEncode(strLeagueName & "")%>&LeagueMatchID=<%=oRs.Fields("LeagueMatchID").Value%>">Rant Board</a></td>
				<%
			oRs.MoveNext
		Loop
			%>
			</table>
			</td>
		</tr>
		</table>
		<%
	End if
End if
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>