<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add Questions"

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

if not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If

Dim bTitle, NumQs, bType, bText, i, j, bdesc, bLadder
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Activate a Ballot")

bTitle=Request.Form("bTitle")
NumQs=Request.Form("qNum")
btype=Request.Form("bType")
bdesc=Request.Form("bdesc")
bLadder=Request.Form("ladder")
if btype=0 then 
	btext="General"
else
	btext="Founders"
end if

Response.Write btitle & "<br>" & numqs & " questions<br>" & bText & " voting<br><br>"
%>
<form name=Qs action=VBallot.asp method=post>
<%
for i = 1 to numqs
	Response.Write "Question " & i & ": <input type=text name=q_" & i & "></br><ul>"
	for j=1 to 5
		Response.Write "<li>Choice " & j & ": <input type=text name=q_" & i & "_A_" & j & "></br>"
	next 
	Response.Write "</ul>"
next 
%>
<input type=hidden name=btitle value="<%=Server.HTMLEncode(btitle)%>">
<input type=hidden name=numqs value=<%=NumQs%>>
<input type=hidden name=btype value=<%=bType%>>
<input type=hidden name=btext value="<%=btext%>">
<input type=hidden name=bdesc value="<%=bdesc%>">
<input type=hidden name=bLadder value="<%=bLadder%>">
<input type=submit>
</form>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>