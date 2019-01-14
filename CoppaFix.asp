<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Coppa Fixer"

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

If Not(bSysAdmin AND Session("uName") = "Triston") Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

If Request.Querystring("PlayerID") <> "" Then
	oConn.Execute ("UPDATE tbl_players SET PlayerCoppa=1 WHERE PlayerID='" & Request.Querystring("PlayerID") & "'")
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Errors Desk") %>
<table width=50% border=0 cellspacing=0 cellpadding=0 align=center BGCOLOR="#444444">
<TR><TD>
<table width=100% border=0 cellspacing=1 cellpadding=4 align=center>
<tr BGCOLOR="#000000">
	<TH>ID</th>
	<TH>Handle</th>
</tr>
<%
	strSQL = "SELECT PlayerHandle, PlayerID FROM tbl_players WHERE PlayerCoppa = 0 ORDER BY PlayerHandle ASC "
	oRs.Open strSQL, oConn
	bgc = bgctwo
	if not (ors.eof and ors.bof) then
		do while not ors.EOF 
			Response.write "<tr bgcolor=" & bgc & ">"
			Response.write "<td>" & Server.HTMLEncode(ors.Fields("PlayerID").Value & "") & "</TD>"
			Response.write "<td><a href=""coppaFix.asp?PlayerID=" & oRs.Fields("PlayeriD").Value & """>" & Server.HTMLEncode(ors.Fields("PlayerHandle").Value & "") & "</a></TD>"
			ors.MoveNext 
			if bgc = bgcone then
				bgc = bgctwo
			else
				bgc = bgcone
			end if 
		loop
	end if
	ors.close
%>
</table></TD></TR>
</TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

