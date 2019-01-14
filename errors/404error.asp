<% Option Explicit %>
<%
Response.Buffer = True
Response.Status = "404 Not Found"

Dim strPageTitle

strPageTitle = "TWL: File Not Found"

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

<% Call ContentStart("File Not Found") %>

	<table width=760 border="0" cellspacing="0" cellpadding="2">
	<tr><td>
	                  <p class=small>You requested a file that is not on Teamwarfare. Please check your URL 
	                  and notify the linking webmaster.</p>
	</td></tr>
	</table>            
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>