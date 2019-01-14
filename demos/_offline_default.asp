<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Demos"

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
Dim strCategory, oFile

If Not(Session("LoggedIn")) Then
	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=2"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("TWL Demo Library") %>
<table width=97% cellspacing=0 cellpadding=0 border=0 ALIGN=CENTER>
<TR><TD>
 Due to excessive bandwidth issues, the TeamWarfare demo section has been taken offline for the time being.  We are sorry for the
 inconvenience.  We will keep you posted on the status of the demo section.
</TD></TR>
</TABLE>

<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
