<% Option Explicit %>
<%
Response.Buffer=True

Dim strPageTitle

strPageTitle = "TWL: Login"

Dim bgcone, bgctwo
			bgcone = "#3C0000"
			bgctwo = "#2B0000"

Dim strFromURL, strError

strFromURL	= Request("url")
strError	= Request("Error")
If (InStr(strFromURL, "activate") <> 0) Then
	strFromURL = "/"
End If

%>
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
If Session("LoggedIn") = True then
  Response.Clear

	'Response.Cookies("User").domain = Request.ServerVariables("SERVER_NAME")
	Response.Cookies("User").path = "/"
	Response.Cookies("User")("uName")=""
	Response.Cookies("User")("UserInfo")=""
	Response.Cookies("User")("UserInfo2")=""
	Response.Cookies("User").expires = "1/1/2038"
	Session("LoggedIn") = False
	session("uName")=""
	Session.Abandon
	If strFromURL = "" Then
		strFromURL = "/"
	end if
	%>
	<script> 
		window.opener.location.href='<%=strFromURL%>';
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
	<input type=hidden name=SecType value=login>
	<input type=hidden name=fromurl value="<%=strFromURL%>">
	<TABLE border=0 cellspacing=0 cellpadding=0 BGCOLOR="#444444">
	<tr valign=center>
		<td align=center>
	<table align=center border=0 cellspacing=1 CELLPADDING=3>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=2>Login to TeamWarfare</TH>
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
