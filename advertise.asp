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
				Fill out the form below, and we will get in contact with you shortly!<br />
				<br />
				<table border="0" cellspacing="0" cellpadding="0" align="center" class="cssBordered">
				<form name="frmAdvertise" id="frmAdvertise" method="post" action="dbAdvertise.asp">
				<tr>
					<td align="right" bgcolor="<%=bgcone%>">Name:</td>
					<td bgcolor="<%=bgctwo%>"><input type="text" name="txtName" id="txtName" style="width: 150px;" /></td>
				</tr>
				<tr>
					<td align="right" bgcolor="<%=bgcone%>">Email Address:</td>
					<td bgcolor="<%=bgctwo%>"><input type="text" name="txtEmail" id="txtEmail" style="width: 250px;" /></td>
				</tr>
				<tr>
					<td align="right" bgcolor="<%=bgcone%>">Contact Phone:</td>
					<td bgcolor="<%=bgctwo%>"><input type="text" name="txtPhone" id="txtPhone" style="width: 100px;" /></td>
				</tr>
				<tr>
					<td align="right" bgcolor="<%=bgcone%>">Company Name:</td>
					<td bgcolor="<%=bgctwo%>"><input type="text" name="txtCompany" id="txtCompany" style="width: 200px;" /></td>
				</tr>
				<tr>
					<td align="right" bgcolor="<%=bgcone%>">Company URL:</td>
					<td bgcolor="<%=bgctwo%>"><input type="text" name="txtCompanyURL" id="txtCompanyURL" style="width: 150px;" /></td>
				</tr>
				<tr>
					<td bgcolor="<%=bgcone%>" colspan="2">Please describe the product or service you would like to advertise on TeamWarfare.<br />
					Be sure to include links to relevant web pages and any exist ad material, if available:</td>
				</tr>
				<tr>
					<td colspan="2" bgcolor="<%=bgctwo%>"><textarea name="txtDetails" id="txtDetails" cols="60" rows="9"></textarea></td>
				</tr>
				<tr>
					<td bgcolor="<%=bgctwo%>" colspan="2" align="right"><input type="submit" value="Submit" /></td>
				</tr>
				</form>
				</table>
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