<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Prize Winners"

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
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("TWL Awarded Prizes") %>

	<table border="0" cellspacing="0" cellpadding="0" ALIGN=CENTER BGCOLOR="#444444" WIDTH=97%>
	<TR><TD>
	<table border="0" cellspacing="1" cellpadding="2" WIDTH=100%>
	<TR BGCOLOR="#000000">
		<TH WIDTH="75">Date</TH>
		<TH>Winner</TH>
		<TH>Prize</TH>
		<TH>Sponsor</TH>
	</TR>
	<TR BGCOLOR="<%=bgcone%>" VALIGN=TOP>
		<TD>05/01/2001</TD>
		<TD><a href="http://www.kungfuclan.com">Kung Fu Clan kfc.</A>
			<BR>Tribes 2 CTF Ladder</TD>
		<TD>$2,000</TD>
		<TD ALIGN=CENTER><a href="http://www.teamsound.com"><IMG SRC="/images/teamsound.jpg" BORDER=0></A><BR>
		<A href="http://www.wiredred.com">WiredRed Software</A></TD>
	</TR>
	<TR BGCOLOR="<%=bgctwo%>" VALIGN=TOP>
		<TD>06/01/2001</TD>
		<TD><a href="http://www.teamfusion.org">Team Fusion TF_</A>
			<BR>Tribes 2 CTF Ladder</TD>
		<TD>$2,000</TD>
		<TD ALIGN=CENTER><a href="http://www.teamsound.com"><IMG SRC="/images/teamsound.jpg" BORDER=0></A><BR>
		<A href="http://www.wiredred.com">WiredRed Software</A></TD>
	</TR>
	<TR BGCOLOR="<%=bgcone%>" VALIGN=TOP>
		<TD>07/01/2001</TD>
		<TD><a href="http://www.teamfusion.org">Team Fusion TF_</A>
			<BR>Tribes 2 CTF Ladder</TD>
		<TD>$2,000</TD>
		<TD ALIGN=CENTER><a href="http://www.teamsound.com"><IMG SRC="/images/teamsound.jpg" BORDER=0></A><BR>
		<A href="http://www.wiredred.com">WiredRed Software</A></TD>
	</TR>
	<TR BGCOLOR="<%=bgctwo%>" VALIGN=TOP>
		<TD>08/01/2001</TD>
		<TD><a href="http://www.team5150.com">Team 5150 |5150|</A>
			<BR>Tribes 2 CTF Ladder</TD>
		<TD>$2,000</TD>
		<TD ALIGN=CENTER><a href="http://www.teamsound.com"><IMG SRC="/images/teamsound.jpg" BORDER=0></A><BR>
		<A href="http://www.wiredred.com">WiredRed Software</A></TD>
	</TR>
	<TR BGCOLOR="<%=bgcone%>" VALIGN=TOP>
		<TD>09/01/2001</TD>
		<TD><a href="http://www.team5150.com">Team 5150 |5150|</A>
			<BR>Tribes 2 CTF Ladder</TD>
		<TD>$2,000</TD>
		<TD ALIGN=CENTER><a href="http://www.teamsound.com"><IMG SRC="/images/teamsound.jpg" BORDER=0></A><BR>
		<A href="http://www.wiredred.com">WiredRed Software</A></TD>
	</TR>
	<TR BGCOLOR="<%=bgctwo%>" VALIGN=TOP>
		<TD>10/01/2001</TD>
		<TD><a href="http://www.team5150.com">Team 5150 |5150|</A>
			<BR>Tribes 2 CTF Ladder</TD>
		<TD>$2,000</TD>
		<TD ALIGN=CENTER><a href="http://www.teamsound.com"><IMG SRC="/images/teamsound.jpg" BORDER=0></A><BR>
		<A href="http://www.wiredred.com">WiredRed Software</A></TD>
	</TR>
	<TR BGCOLOR="<%=bgcone%>" VALIGN=TOP>
		<TD>11/01/2001</TD>
		<TD><a href="http://www.team5150.com">Team 5150 |5150|</A>
			<BR>Tribes 2 CTF Ladder</TD>
		<TD>$2,000</TD>
		<TD ALIGN=CENTER><a href="http://www.teamsound.com"><IMG SRC="/images/teamsound.jpg" BORDER=0></A><BR>
		<A href="http://www.wiredred.com">WiredRed Software</A></TD>
	</TR>
	<TR BGCOLOR="<%=bgctwo%>" VALIGN=TOP>
		<TD>12/01/2001</TD>
		<TD><a href="http://tribes.tribalwar.com/vanguard/">Team Vanguard =V=</A>
			<BR>Tribes 2 CTF Ladder</TD>
		<TD>$2,000</TD>
		<TD ALIGN=CENTER><a href="http://www.teamsound.com"><IMG SRC="/images/teamsound.jpg" BORDER=0></A><BR>
		<A href="http://www.wiredred.com">WiredRed Software</A></TD>
	</TR>
	<TR BGCOLOR="<%=bgcone%>" VALIGN=TOP>
		<TD>12/01/2001</TD>
		<TD>-translucint-
			<BR>GeeZer's Tribes 2 CTF 7 Man Tournament</TD>
		<TD>GeForce 3 TI</TD>
		<TD ALIGN=CENTER><A href="http://www.northstarpc.com">NorthStar PC</A></TD>
	</TR>
</TABLE>
</TD></TR>
</TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

