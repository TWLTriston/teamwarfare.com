<% Option Explicit %>
<%
Response.Buffer=True
Dim strPageTitle

strPageTitle = "TWL: Report Loss"

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

Dim bSysAdmin, strLadder, bLadderAdmin, strTeam, intTeamID, intTournamentID
bSysAdmin = IsSysAdmin()
intTeamID = Request.QueryString("teamid")
intTournamentID = Request.QueryString("tournamentid")
Dim mname, tlName
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
	strSQL="select teamname from tbl_teams where teamid='" & intTeamID & "'"
	'strSQL="select teamname from tbl_teams inner join lnk_T_L on lnk_T_L.TeamID=tbl_teams.teamID where lnk_T_L.TLLinkID=" & Request.QueryString("teamid")
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		tlName=ors.Fields(0).Value
	end if
	ors.NextRecordset 
	strSQL="select TournamentName from tbl_tournaments where tournamentid='" & intTournamentID & "'"
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		mName=ors.Fields(0).Value
	end if
	oRs.NextRecordset 
	%>
<FORM action=/tournament/savetournament.asp method=post name=reportloss>
<table align=center border=0 cellspacing=2 cellpadding=0 width=100%>
<tr bgcolor=<%=bgcone%> valign=center><td colspan=2 align=center height=30>
<p class=small><b>Report Loss: <% =Server.HTMLEncode (tlname) %> on the <%=Server.HTMLEncode(mname)%></b></p></td></tr>
<input type=hidden name=jdate value=<%=now()%>>
<input id=RoundsID name=RoundsID type=hidden value=<%=Request.QueryString("RoundsID")%>>
<input id=LinkID name=LinkID type=hidden value=<%=Request.QueryString("LinkID")%>>
<input id=savetype name=savetype type=hidden value="ReportLoss">
<input id=hidden name=TournamentID type=hidden value=<%=intTournamentID%>>
<input id=hidden name=Teamid type=hidden value=<%=intTeamID%>>
<input id=hidden name=fromurl type=hidden value="<%=Server.HTMLEncode(strFromURL & "")%>">

<tr bgcolor=<%=bgctwo%> height=25>
<td align=center>&nbsp;&nbsp;<INPUT type=submit value="Verify Report Loss" id=submit name=submit></td></tr>
</table>
</FORM>
</td></tr>
</table>
</body>
</html>
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing	
%>