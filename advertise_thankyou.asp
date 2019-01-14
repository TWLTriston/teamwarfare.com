<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "Advertising on TeamWarfare"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgcblack, bgcheader

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

    <% Call ContentStart("Advertise on TeamWarfare") %>
		<table border="0" cellspacing="0" cellpadding="0" width="97%" align="center">
		<tr>
			<td>
				Thank you for your interest in advertising with TeamWarfare.<br />
				<br />
				Your request has been submitted and we will be in contact shortly. If you would like to provide any additional information, please send it to <a href="mailto:ads@teamwarfare.com?Subject=Advertising+on+TeamWarfare">ads@teamwarfare.com</a><br />
				<br />
				
			</td>
		</tr>
		</table>
	<% Call ContentEnd() %>
	<!-- #include virtual="/include/i_footer.asp" //-->
	
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Response.End
%>