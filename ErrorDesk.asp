<% Option Explicit %>
<%
Response.Buffer = False

Dim strPageTitle

strPageTitle = "TWL: Errors Desk"

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

If Request.Querystring("ErrorID") <> "" Then
	oConn.Execute ("UPDATE tbl_errors SET status='" & Request.Querystring("status") & "' WHERE ErrorID='" & Request.Querystring("ErrorID") & "'")
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Errors Desk") %>
<table width=97% border=0 cellspacing=0 cellpadding=0 align=center BGCOLOR="#444444">
<TR><TD>
<table width=100% border=0 cellspacing=1 cellpadding=2 align=center>
<tr BGCOLOR="#000000">
	<TH>ID</th>
	<TH>Page</th>
	<TH>Line</th>
	<TH>Player</th>
	<TH>Status</th>
	<TH>Date</th>
	<TH>Close</th>
</tr>
<%
	strSQL = "SELECT * FROM tbl_errors WHERE Status IS NULL ORDER BY ErrorID DESC"
	oRs.Open strSQL, oConn
	bgc = bgctwo
	if not (ors.eof and ors.bof) then
		do while not ors.EOF 
			Response.write "<tr bgcolor=" & bgcone & ">"
			Response.write "<td>" & Server.HTMLEncode(ors.Fields("ErrorID").Value & "") & "</TD>"
			Response.write "<td>" & Server.HTMLEncode(ors.Fields("Page").Value & "") & "</TD>"
			Response.write "<td>" & Server.HTMLEncode(ors.Fields("Line").Value & "") & "</TD>"
			Response.write "<td>" & Server.HTMLEncode(ors.Fields("Player").Value & "") & "</TD>"
			Response.write "<td>" & Server.HTMLEncode(ors.Fields("Status").Value & "") & "</TD>"
			Response.write "<td>" & Server.HTMLEncode(ors.Fields("ins_dtim").Value & "") & "</TD>"
			Response.write "<td align=center><a href=errordesk.asp?errorid=" & oRS.Fields("ErrorID").Value & "&status=1>close</A></TD></TR>"
			Response.write "<TR bgcolor=" & bgctwo & "><TD COLSPAN=7>" & Replace(Replace(Server.HTMLEncode(oRS.Fields("ErrorMessage").Value),"&amp;", vbCrLf & "&nbsp;&nbsp;&amp;"), vbCrLf, "<BR>") & "</TD></TR>"
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

