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

Dim bSysAdmin, strLadder, bLadderAdmin, strTeam
bSysAdmin = IsSysAdmin()

Dim strFromURL, jointype, intTeamID, intLeagueID
strFromURL = Request("url")
intTeamID = Request("TeamID")
intLeagueID = Request.QueryString("LeagueID")

Dim pid, tlName, lName
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<HTML>
<HEAD>
<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
<title><%=strPageTitle%></title>
</HEAD>
<body bgcolor="#000000" leftmargin="0" topmargin="00" marginwidth="000" marginheight="0000">
<table width=100% height=100%><tr valign=center><td align=center>
<%
	strSQL = "select PlayerID from tbl_players where PlayerHandle = '" & CheckString(Session("uName")) & "'" 
	ors.open strsql, oconn
	if not (ors.eof and ors.eof) then
		pID = ors.fields(0).value
	end if
	ors.NextRecordset

	strSQL="select teamname from tbl_teams where teamid='" & intTeamID & "'"
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		tlName=ors.Fields(0).Value
	end if
	ors.NextRecordset
	strSQL="select LeagueName from tbl_leagues where LeagueID='" & intLeagueID & "'"
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		lName=ors.Fields(0).Value
	end if
	oRs.NextRecordset 
if not(pID = "" or lname = "" or tlname = "") then
	%>
	<form action=security.asp method=post name=frmQuitTeam>
	<table align=center border=0 cellspacing=2 cellpadding=0 width=100%>
	<tr bgcolor=<%=bgcone%> height=30 valign=center><td align=center><p class=text><font color=red><b>Click quit again to quit:</b></font></p></td></tr>
	<tr bgcolor=<%=bgctwo%> valign=center height=30><td align=center><p class=small>Team:&nbsp;<%=Server.HTMLEncode (tlname)%></p></td></tr>
	<tr bgcolor=<%=bgcone%> valign=center height=30><td align=center><p class=small>League:&nbsp;<%=Server.HTMLEncode (lname)%>&nbsp;League</p></td></tr>
	<tr bgcolor=<%=bgctwo%> valign=center height=30><td align=center>
	<input id=KeyData name=TeamID type=hidden value=<%=intTeamID%>>
	<input id=SecType name=SecType type=hidden value="TeamLeagueQuit">
	<input id=PlayerID name=PlayerID type=hidden value=<%=pID%>>
	<input id=hidden name=LeagueToQuit type=hidden value="<%=Server.HTMLEncode (lname)%>">
	<input id=hidden name=LeagueID type=hidden value=<%=intLeagueID%>>
	<input id=hidden name=fromurl type=hidden value="<%=strFromURL%>">
	<INPUT type="submit" value="Quit" id=submit1 name=submit1 class=bright></td></tr>
	</table>
	</FORM>
<% else
	response.write "Invalid data passed, please try again."
end if
%>
</td></tr>
</table>
</body>
</html>
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>