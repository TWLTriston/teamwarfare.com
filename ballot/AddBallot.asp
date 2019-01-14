<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Make Ballot"

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
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Activate a Ballot")
%>

<form name=Ballot action=AddQs.asp method=post>
    <table width="90%" border="0">
<table width=50% border=0 cellpadding=2 cellspacing=0>
<tr><td>Ballot Title</td><td><input type=text name=bTitle></td></tr>
<tr><td>Number of Questions</td><td><input type=text name=qNum></td></tr>
<tr><td>Ballot Type</td><td><select name=bType><option value=0>Global</option><option value=1>Founders</option></select></td></tr>
<TR><TD>Long Description</TD><TD><TEXTAREA NAME=bDesc ROWS=3 COLS=30></TEXTAREA></TD></TR>
<TR><TD>Associated Ladder / League</td><td><select name=ladder><option value=0>Select a ladder</option>
<%
dim oCmd
set oCmd = server.CreateObject("adodb.command")
ocmd.ActiveConnection = oconn
	ocmd.CommandText = "usp_Select_LadderList"
	ocmd.CommandType = &H0004 
	set oRs = ocmd.Execute
	if not (ors.EOF and ors.BOF) then
		do while not ors.EOF
			If oRS.Fields ("LadderType").Value = "T" Then
				Response.Write "<option value=T" & ors.Fields("ladderid").Value & ">" & ors.Fields("laddername").Value & " Ladder</option>"
			ElseIf oRS.Fields ("LadderType").Value = "P" Then
'				Response.Write "<option value=P" & ors.Fields("ladderid").Value & ">" & ors.Fields("laddername").Value & " Player Ladder</option>"
			ElseIf oRS.Fields ("LadderType").Value = "L" Then
				Response.Write "<option value=L" & ors.Fields("ladderid").Value & ">" & ors.Fields("laddername").Value & " League</option>"
			End If
			ors.MoveNext
		loop
	end if
%>
</select></td></tr>
<tr><td colspan=2 align=center><input type=submit></td></tr>
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