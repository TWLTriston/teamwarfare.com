<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add an Scrim Ladder"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Scrim Ladder Administration")
%>

	<table align=center bgcolor="#444444" cellspacing=0 cellpadding=0 width="550">
	<tr><td>
	<table align=center cellspacing=1 cellpadding=4 width=100%>
		<tr><th colspan=4 bgcolor="#000000"><a href="LadderAdd.asp">Add a new ladder</a></th></tr>
		<tr><th colspan=4 bgcolor="#000000">Choose a Ladder to Admin</th></tr>
		<%
		strSQL = "SELECT EloLadderID, EloLadderName FROM tbl_elo_ladders ORDER BY EloLadderName ASC "
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRs.BOF) Then
			Do While Not(oRs.EOF)
				If bgc = bgcone Then
					bgc = bgctwo
				Else
					bgc = bgcone
				End If
				%>
				<tr>
					<td bgcolor="<%=bgc%>"><a href="/viewscrimladder.asp?ladder=<%=Server.URLEncode(oRs.Fields("EloLadderName").Value)%>"><%=Server.HTMLEncode(oRs.Fields("EloLadderName").Value)%> Ladder</a></td>
					<td align="center" bgcolor="<%=bgc%>"><a href="LadderAdd.asp?isedit=true&Ladder=<%=Server.URLEncode(oRs.Fields("EloLadderName").Value)%>">edit</a></td>
					<td align="center" bgcolor="<%=bgc%>"><a href="MapList.asp?isedit=true&Ladder=<%=Server.URLEncode(oRs.Fields("EloLadderName").Value)%>">Map List</a></td>
					<td align="center" bgcolor="<%=bgc%>"><a href="/reports/scrimidentifierreport.asp?lname=<%=Server.URLEncode(oRs.Fields("EloLadderName").Value)%>">GUID Report</a></td>
				</tr>					
				<%
				oRs.MoveNext
			Loop
		End If
		oRs.NextRecordSet
		%>
</table>
</td></tr></table>
<%
Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>