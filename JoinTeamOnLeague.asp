<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Join New Competition"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()
Call SetStyle()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim intTeamID, intLeagueID, strJoinType, strFromURL, strError
Dim strTeamName, strLeagueName, intIdentifierID, needGUID
intTeamID = Request("TeamID")
intLeagueID = Request("LeagueID")
strJoinType = Request("type")
strFromURL = Request("url")
strError = Request("Error")
%>
<!-- #Include virtual="/include/i_funclib.asp" -->

<HTML>
<HEAD>
<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
<title><%=strPageTitle%></title>
</HEAD>
<HTML>
<body bgcolor="#000000" leftmargin="0" topmargin="00" marginwidth="000" marginheight="0000" <% If strError = "" Then %> ONLOAD="self.document.frmCallSec.password.focus();" <% End If %>>
<table width=100% height=100%><tr valign=center><td align=center>
<% if strError = "2" then %>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER>
	<TR><TD>
	<table align=center border=0 cellspacing=1 cellpadding=2 width=100%>
		<tr bgcolor=#000000><tH><font color=red>Unable to process request</font></tH></tr>
		<tr bgcolor=<%=bgctwo%>><td>You cannot join two teams on the same league.</td></tr>
		<tr bgcolor=<%=bgcone%>><td align=center><input type="button" class=bright value="Close" onclick="window.close();"></td></tr>
	</table>
	</TD></TR>
	</TABLE>
	<%
else
	strSQL="select teamname from tbl_teams where teamid='" & intTeamID & "'"
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		strTeamName=ors.Fields(0).Value
	end if
	ors.NextRecordset 
	strSQL="select LeagueName, IdentifierID from tbl_leagues where LeagueID='" & intLeagueID & "'"
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		strLeagueName = ors.Fields(0).Value
		intIdentifierID = ors.Fields(1).Value
	end if
	oRs.NextRecordset 
	if intIdentifierID > 0 then
		needGUID = 1
	else
		needGUID = 0
	end if
	%>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER>
	<TR><TD>
	<table align=center border=0 cellspacing=1 cellpadding=2 width=100%>
	<FORM action=security.asp method=post name=frmCallSec>
	<input type=hidden name=LeagueToJoin value="<%=Server.HTMLEncode (strLeagueName)%>">
	<input id=SecType name=SecType type=hidden value="TeamLeagueJoin">
	<input id=hidden name=LeagueID type=hidden value=<%=intLeagueID%>>
	<input id=hidden name=Teamid type=hidden value=<%=intTeamID%>>
	<input id=hidden name=fromurl type=hidden value="<%=strFromURL%>">
	<input id=KeyData name=KeyData type=hidden value="<% =Server.HTMLEncode (strTeamName)%>">
	<input id=IdentifierID name=IdentifierID type=hidden value=<%=intIdentifierID%>>
	<input id=needGUID name=needGUID type=hidden value=<%=needGUID%>>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=2>Join a Team</TH>
	</TR>
	<TR BGCOLOR="<%=bgcone%>">
		<TD><B>Team:</B></TD><TD><%=strTeamName%></TD>
	</TR>
	<TR BGCOLOR="<%=bgctwo%>">
		<TD><B>League:</B></TD><TD><%=strLeagueName%> League</TD>
	</TR>
	<tr bgcolor=<%=bgcone%> height=25>
		<td align=right>Join Password:</td>
		<td align=left>&nbsp;&nbsp;<INPUT id=password name=password type=password class=text style="WIDTH: 150px"></td></tr>
	<tr bgcolor=<%=bgctwo%>>
		<td align=center colspan=2><INPUT type="submit" value="Join" id=submit1 name=submit1 class=bright></td></tr>
	<% if strError=1 then %>
		<tr valign=center bgcolor=<%=bgcone%>><td align=center colspan=2><font color=red><b>Incorrect Password</b></font></td></tr>
	<%end if%>
	<% if strError=9 then %>
		<tr valign=center bgcolor=<%=bgcone%>><td align=center colspan=2><font color=red><b>Must Have Anti-Smurf/GUID entered to Join this Competition</b></font></td></tr>
	<%end if%>
	</FORM>
	</table>
	</td></tr>
	</TABLE>
<% end if %>
</TD></TR>
</table>
</body>
</html>
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>