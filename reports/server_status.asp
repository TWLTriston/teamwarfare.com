<% Option Explicit %>
<%
Server.ScriptTimeout = 6000
Response.Buffer = False

Dim strPageTitle

strPageTitle = "TWL: Server Status"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.CommandTimeout = 0
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin) then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim intSpot
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Server Status")

strSQL = "EXECUTE sp_server_status"
oRs.Open strSQL, oConn
If Not(oRS.EOF AND oRS.BOF) Then
	%>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER>
	<TR><TD>
		<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2>
		<TR BGCOLOR="#000000">
			<TH>Statistic</TH>
			<TH>Count</TH>
		</TR>
		<%
		For intSpot = 0 To (oRS.Fields.Count - 1) Step 1
			If bgc = bgcOne then
				bgc = bgctwo
			Else
				bgc = bgcone
			End If
			Response.Write "<TR BGCOLOR=" & bgc & ">"
			Response.Write "<TD><B>" & oRS.Fields(intSpot).Name & "</B></TD>"
			Response.Write "<TD ALIGN=RIGHT>" & oRs.Fields(intSpot).Value & "</TD>"
			Response.Write "</TR>"
		Next
		%>
		</TABLE>
	</TD></TR>
	</TABLE>
	<%
End If
oRs.Close
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>