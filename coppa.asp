<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Age Verification"
Session("LoggedIn") = ""
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
'Response.Write "--" & ()
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Age Verification")
%>
In order to become COPPA compliant, we are now requiring age verification before<br />
 access to the members portion of the site can be obtained. We have made this <br />
 process as painless as possible.
 <br />
<br />
<b>Please click the link that corresponds with your birth date:</b><br />
<br />
<br /><br /><br /><br /><b>IT IS IMPERATIVE THAT YOU CLICK THE CORRECT LINK BELOW<br /><br /><br /><br /><br />
<font color="#ff0000"><b>Currently 13 and Older<b></font>: <a href="SaveItem.asp?SaveType=Coppa&Age=1">If you were born on or before <%=FormatDateTime(Now() - (365*13), 2)%>, click here.</a> <br />
<br /> 
<font color="#ff0000"><b>Currently under 13 years of age</b></font>: <a href="SaveItem.asp?SaveType=Coppa&Age=0">If you were born after <%=FormatDateTime(Now() - (365*13), 2)%>, click here. </a><br />
<br />
For information about how this site uses personal information, please read the <a href="/privacy.asp">TeamWarfare Privacy Policy</a>

<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>