<% Option Explicit %>
<%
Response.Buffer=True
Dim strPageTitle

strPageTitle = "TWL: Quit Ladder"

Dim bgcone, bgctwo, bgcblack, bgcheader

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

Dim bSysAdmin, strLadder, bLadderAdmin, strTeam
bSysAdmin = IsSysAdmin()
strTeam = Request.QueryString("team")
strLadder = Request.QueryString("ladder")

Dim bTeamCaptain, bTeamFounder, intTeamID
bSysAdmin = IsSysAdmin()
bLadderAdmin = IsEloLadderAdmin(strLadder)
bTeamFounder = IsTeamFounder(strTeam)
bTeamCaptain = IsTeamCaptain(strTeam, strLadder)

if not(bSysAdmin or bLadderAdmin or bTeamFounder or bTeamCaptain) then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
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
		response.write "<tr bgcolor=" & bgcone & " height=30><td align=center>Unable to process request. Team currently has an active challenge or is on rest.</td></tr>"	
	else
		strsql = "select TeamID from tbl_teams where teamname='" & CheckString(strTeam) & "'"
		ors.open strsql, oconn
		if not (ors.bof and ors.eof) then
			intTeamID=ors.fields(0).value
		end if
		ors.close
	
		%>	
		<form method=post action=saveitem.asp>
		<tr bgcolor=<%=bgcone%>><td align=center><font color=ffcf3f><b>WARNING THIS IS NOT REVERSABLE</b></font></td></tr>
		<tr bgcolor=<%=bgctwo%>><td align=center><font color=ffcf3f><b><%=Server.HTMLEncode(strTeam)%></b></font></td></tr>
		<tr bgcolor=<%=bgcone%>>
		<td valign=center align=center>
		<input type=hidden name=SaveType value="QuitLadder">
		<input type=hidden name=teamname value="<%=Server.HTMLEncode (strTeam)%>">
		<input type=hidden name=teamid value="<%=intTeamID%>">
		<input type=hidden name=laddername value="<%=Server.HTMLEncode(strLadder)%>">
		<input id=hidden name=fromurl type=hidden value="<%=Server.HTMLEncode(strFromURL & "")%>">
		<input type=submit name=submit style="width:200" value="Confirm Quit <%=Server.HTMLEncode(strLadder)%> Ladder" class=bright>
		</td>
		</tr>
		<tr bgcolor=<%=bgctwo%> height=25><td align=center><font color=ffcf3f><b>WARNING THIS IS NOT REVERSABLE</b></font></td></tr>
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