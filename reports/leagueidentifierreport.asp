<% 'Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "Roster Report"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Set oRS2 = Nothing
	Response.Clear 
	Response.Redirect "/errorpage.asp?error=3"
End If
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("GUID Report") %>
<form name=ladder id=ladder method=get action=Leagueidentifierreport.asp>
	<table align=center>
<TR><TD>Ladder:</TD><TD>
<select name=lname id=lname>
<%
strsql = "select LeagueName, LeagueID from tbl_Leagues where LeagueActive = 1 AND IdentifierID <> 0 AND IdentifierID IS NOT NULL order by LeagueName asc"
ors.open strsql, oconn
if not(ors.eof and ors.bof) then
	do while not(ors.eof)
		Response.Write "<option value=""" & ors.fields("LeagueName").value & """"
		if Request.QueryString("lname") = ors.fields("LeagueName").value then
			Response.Write " selected "
		end if
		Response.Write ">" & ors.fields("LeagueName").value & "</option>" & vbcrlf
		ors.movenext
	loop
end if
ors.nextrecordset
%>
</select>
</td></TR>
<tr><TD align=center colspan=2><input type=submit id=submit name=submit value="Run Query"></TD></TR>
</form>
</Table>
<BR><BR>
<%
bgc = bgcone
totalteam = 0
undermin = 0
overmin = 0
overdate = 0
underdate = 0
violation = 0
dim strTeam
strTeam = ""
if Len(Request.QueryString("lname")) > 0 then
	strsql = "EXECUTE LeagueReportCheckIdentifier @LeagueName = '" & CheckString(Request.QueryString("lname")) & "'"
	ors.open strsql, oconn
'	Response.Write strSQL
	if not(ors.eof and ors.bof) then
		do while not(ors.eof)
			If strTeam <> oRs.Fields("TeamName") Then
				If strTeam <> "" Then
					Response.Write "</table></td></tr></table><br /><br />"
				End If
				%>
				<table border="0" cellspacing="0" cellpadding="0" width="400" bgcolor="#444444">
				<tr><td>
				<table border="0" cellspacing="1" cellpadding="4" width="100%">
				<tr><th bgcolor="#000000"><a href="/viewteam.asp?team=<%=Server.URLEncode(oRs.Fields("TeamName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("TeamName").Value)%></a></th></tr>
				<%
				strTeam = oRs.FieldS("teamName").Value
			End If
			%>
			<tr>
				<td bgcolor="<%=bgcone%>"><a href="/viewplayer.asp?player=<%=Server.URLEncode(oRs.Fields("PlayerHandle").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("PlayerHandle").Value)%></a></td></tr>
			</tr>
			<%
			if bgc = bgcone then
				bgc = bgctwo
			else
				bgc = bgcone
			end if
			ors.movenext
		loop
		%>
		</table></td></tr></table>		
		<%
	Else
		%>
		<b>There are no players who are in violation.</b>
		<%
	end if
	ors.nextrecordset	    
end if
%>

<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->

<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>