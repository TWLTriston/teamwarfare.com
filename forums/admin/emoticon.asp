<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Emoticons Administration"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

if not bSysAdmin then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Set oRS2 = Nothing
	Response.Clear
	Response.Redirect "/default.asp"
end if

Dim oFS
Set oFS = Server.CreateObject("Scripting.FileSystemObject")

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<% 
Call ContentStart("Emoticon Administration")

Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oFS = Nothing
Set oConn = Nothing
Set oRS2 = Nothing
Set oRS = Nothing
%>