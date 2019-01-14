<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Verify Ballot"

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
Dim i, j, numqs
Dim formItem

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Verify Ballot")

Response.Write Request.Form("btitle") & "<br>"
Response.Write Request.Form("numqs") & " questions<br>"
Response.Write Request.Form("btext") & " voting<br>"
numqs=Request.Form("numqs")
for i = 1 to numqs
	Response.Write "<b>" & Request.Form("q_" & i) & "</b><ul>"
	for j = 1 to 5
		if trim(Request.Form("q_" & i & "_a_" & j)) <> "" then
			Response.Write "<li>" & Request.Form("q_" & i & "_a_" & j)
		end if
	next
	Response.Write "</ul>"
next

%>
<FORM NAME="ballot" ACTION="saveBallot.asp" METHOD="post">
<%
For Each formItem In Request.Form
	Response.write "<INPUT TYPE=""HIDDEN"" NAME=""" & formItem & """ ID=""" & formItem & """ VALUE=""" & Request.Form(formItem) & """>" & vbCrLf
Next
%>
<INPUT TYPE=SUBMIT VALUE="Save This Ballot">
</FORM>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>