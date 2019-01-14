<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Account Saved"

Dim strSQL, oConn, oRS
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

<%
Call ContentStart("Account Saved")
%>
<TABLE BORDER=0 ALIGN=CENTER WIDTH=90%>
<TR>
	<TD>Your account has been saved. If you are a new registrant, or you changed your email address for your profile, you will be receiving an email with an activation code shortly.
	<br /><br />
	If you do not receive this email, your account may have an invalid email address. Please visit #teamwarfare on irc.gamesurge.net:6667 or email us at sitesupport at teamwarfare.com and we will assist you.
	</td>
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