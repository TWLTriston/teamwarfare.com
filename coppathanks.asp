<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Age Verification"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open
Session("LoggedIn") = ""
Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Age Verification - Confirmation")
Dim intPlayerCoppa
strSQL = "SELECT PlayerCoppa FROM tbL_players WHERE PlayerID ='" & Session("PlayeriD") & "'"
oRs.Open strSQl, oConn
If Not(oRs.EOF AND oRS.BOF) Then
	intPlayerCoppa = oRs.Fields("PlayerCoppa").Value
End If

If intPlayerCoppa = 1 Then
	%>
	Thank you, you may now proceed to your intended destination.
	<%
Else
	%>
	Unfortunately, at this time, we do not allow registration from persons under the age of 13.<br />
	 Please return to TeamWarfare when you are 13 years or older.<br />
	<br />
	Questions on this policy can be directed at <a href="mailto:triston@teamwarfare.com">triston@teamwarfare.com</a>.<br />
	<br />
	<a href="coppachange.asp">If you are over the age of 13, click here to update your account.</a>
	<%
End if
%>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>