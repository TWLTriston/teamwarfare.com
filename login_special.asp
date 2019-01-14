<% Option Explicit %>
<%
Response.Buffer=True
Dim strPageTitle

strPageTitle = "TWL: Login"

Dim bgcone, bgctwo
Dim strSQL, oConn, oRS, oRS2

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()
Dim strFromURL, strError

strFromURL	= Request("url")
strError	= Request("Error")
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<HTML>
<HEAD>
<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
<title><%=strPageTitle%></title>
<SCRIPT Language="javascript">
<!--
	function register() {
		window.opener.location.href='/addplayer.asp';
		window.close();
	}
//-->
</SCRIPT>
</HEAD>
<body bgcolor="#000000" leftmargin="0" topmargin="00" marginwidth="000" marginheight="0000" ONLOAD="self.document.frmCallSec.uName.focus();">
<%
If Not(IsSysAdmin()) Then 
	Response.Cookies("User")("UName")=""
	Response.Cookies("User")("UserInfo")=""
	Session.Abandon
	If strFromURL = "" Then
		strFromURL = "/"
	end if
    Response.Clear
	%>
	<script> 
		if (window.opener) {
			window.opener.location.href='<%=strFromURL%>';
		}
		window.close();
	</script>
	<%
	Response.End 
End If
%>
<TABLE height=100% width=100% border=0 cellspacing=0 cellpadding=0 valign=center align=center>
<tr valign=center>
	<form name=frmCallSec action=security.asp method=post>
	<td align=center>
	<input type=hidden name=SecType value=SpecLogin>
	<input type=hidden name=fromurl value="<%=strFromURL%>">
	<TABLE border=0 cellspacing=0 cellpadding=0 BGCOLOR="#444444">
	<tr valign=center>
		<td align=center>
	<table align=center border=0 cellspacing=1 CELLPADDING=3>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=2>Login as Someone</TH>
	</TR>
	<% If strError = "true" Then %>
	<tr height=18 bgcolor=<%=bgcone%>><td colspan=2 align=center><font color=red>Incorrect name or password</font></td></tr>
	<%
	end if
	%>

	<tr height=30 bgcolor=<%=bgcone%>>
		<td align=right><b>Player Name:</b></td>
		<td>&nbsp;<INPUT id=uName name=uName style="WIDTH: 150px" class=text></td>
	</tr>
	<tr height=30 bgcolor=<%=bgctwo%>>
		<td align=right><b>Password:</b></td>
		<td>&nbsp;<INPUT id=uPassword name=uPassword type=password style=" WIDTH: 150px" class=text></td>
	</tr>
	<tr height=30 bgcolor=<%=bgcone%>>
		<td colspan=2 align=center><INPUT id=submit1 name=submit1 type=submit value=Login class=bright></td>
	</tr>
	<tr height=30 bgcolor=<%=bgctwo%>>
		<td colspan=2 align=center>
			<input type=button class=bright value="Register" onclick="javascript:register();"></td></tr>
	</table>
	</td>
	</tr>
	</table>
	</td>
	</tr>
	</form>
	</table>
</BODY>
</HTML>
