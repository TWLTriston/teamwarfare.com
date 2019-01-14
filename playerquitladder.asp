<% Option Explicit %>
<%
Response.Buffer=True

Dim strPageTitle

strPageTitle = "TWL: Quit Ladder"

Dim bgcone, bgctwo

Dim strSQL, oConn, oRS, oRS2

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()
Call SetStyle()

Dim strFromURL, strError
strFromURL	= Request("url")
strError	= Request("Error")

Dim strPlayer, bSysAdmin, strLadder, bLadderAdmin, intPlayerID

intPlayerID = Request.QueryString("playerid")
strLadder = Request.QueryString("ladder")
bLadderAdmin = IsPlayerLadderAdmin(strLadder)
bSysAdmin = IsSysAdmin()
'Response.write intPlayerID & "--" & Session("PlayerID")
'Response.End
if not(bSysAdmin or bLadderAdmin or cstr(Session("PlayerID")) = cstr(intPlayerID)) then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear 
	Response.Redirect  "/errorpage.asp?error=3"
end if

%>
<!-- #include virtual="/include/i_funclib.asp" -->
<HTML>
<HEAD>
<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
<title><%=strPageTitle%></title>
</HEAD>
<body bgcolor="#000000" leftmargin="0" topmargin="00" marginwidth="000" marginheight="0000">
<table width=100% height=100%><tr valign=center><td align=center>
	<table align=center border=0 cellspacing=2 cellpadding=0 width=100%>
	<%
	if strError = 1 then
		response.write "<tr bgcolor=" & bgcone & " height=30><td align=center><p class=small>Unable to process request. Player currently has an active challenge.</p></td></tr>"
	else
		%>	
		<form method=post action=saveitem.asp>
		<tr bgcolor=<%=bgcone%> height=25><td align=center><p class=small><font color=ffcf3f><b>WARNING THIS IS NOT REVERSABLE</b></font></p></td></tr>
		<tr bgcolor=<%=bgcone%> height=25>
		<td valign=center align=center>
		<input type=hidden name=SaveType value="PlayerQuitLadder">
		<input type=hidden name=playerid value="<%=intPlayerID%>">
		<input type=hidden name=laddername value="<%=strLadder%>">
		<input id=hidden name=fromurl type=hidden value=<%=strFromURL%>>
		<input type=submit name=submit style="width:250" value="Confirm Quit <%=Server.HTMLEncode(strLadder)%> Ladder" class=bright>
		</td>
		</tr>
		<tr bgcolor=<%=bgctwo%> height=25><td align=center><p class=small><font color=ffcf3f><b>WARNING THIS IS NOT REVERSABLE</b></font></p></td></tr>
		</form>
	<% end if %>
	</table>
</td></tr>
</table>
</body>
<% 
oConn.Close 
set ors = nothing
set oConn = nothing	
%>
</html>