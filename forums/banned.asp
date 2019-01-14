<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Banned from Forums"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("") %>
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR=#444444>
<TR>
	<TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=4>
	<TR>
		<TH BGCOLOR="#000000">Unable to Post on Forums</TH>
	</TR>
	<TR>
		<TD BGCOLOR=<%=bgcone%>>Your posting privledges have been revoked.</TD>
	</TR>
	</TABLE>
	</TD>
</TR>
</TABLE>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>