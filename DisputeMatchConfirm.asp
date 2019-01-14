<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle, intDisputeTeamID
intDisputeTeamID = Request.QueryString("DisputeTeamID")

strPageTitle = "TWL: Dispute Match Submitted" 

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo, strHeaderColor
strHeaderColor	= Application("HeaderColor")

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLoggedIn
bLoggedIn = Session("LoggedIn")
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Match Dispute Submitted")
%>	
	<table border="0" align="center" width="52%"><tr><td>
	Your dispute has been submitted. Please allow 24 hours for a response from your ladder admin. If no
	response is received, you can follow up via <a href="staff.asp">email</a>, or irc.<br /><br />
	<center><b>Do not submit more than once. It will only delay the processing of the dispute.</b></center>
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