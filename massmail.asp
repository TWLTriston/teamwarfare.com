<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Mass Mailer"

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

If Not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If	
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Send out a mass e-mail...") %>
<table align=center CELLPADDING=4 CELLSPACING=1 BORDER=0>
<form name=frmMassMailer action=../saveitem.asp method=post>
<tr bgcolor=<%=bgcone%>><td align=center COLSPAN=2><p>Enter the e-mail below</p></td></tr>
<tr bgcolor=<%=bgctwo%>><td COLSPAN=2>Your E-mail: <input type=text name=Email maxlength=50 class=text></td></tr>
<tr bgcolor=<%=bgcone%>><td COLSPAN=2>Subject: <input type=text name=Subject maxlength=50 class=text></td></tr>
<tr bgcolor=<%=bgctwo%>><td align=center height=200 COLSPAN=2>
<textarea name=MailBody cols=40 rows=10></textarea>
</td></tr>
<tr bgcolor=<%=bgcone%>><td align=center COLSPAN=2>
<input type=hidden name=SaveType value="SendMail" class=text>
<input type=submit name=submit1 value=submit class=bright>
</td></tr>
</table>

<% 
dim oCmd
set oCmd = server.CreateObject("adodb.command")
ocmd.ActiveConnection = oconn
ocmd.CommandText = "usp_Select_LadderList2"
ocmd.CommandType = &H0004 
set oRs = ocmd.Execute
Dim intGameID
intGameID = -1
bgc = bgcone
'strsql = "select LadderID, LadderName from tbl_ladders order by laddername asc "
'ors.open strsql, oconn
if not(ors.eof and ors.bof) then
	do while not ors.EOF
			if intGameID <> oRS.Fields("GameID") Then
				If intGameID <> -1 Then
					Response.Write "</TABLE></TD></TR></TABLE><BR><BR>"
				End If
				intGameID = oRS.Fields("GameID").Value
				%>
				<a name="Game<%=intGameID%>"></a>
				<table border="0" cellspacing="0" cellpadding="0" ALIGN=CENTER BGCOLOR="#444444">
				<TR><TD>
				<table border="0" cellspacing="1" cellpadding="2" WIDTH=600>
				<TR BGCOLOR="#000000">
					<TH COLSPAN=2><%=oRS.Fields("GameName").Value%> ( <%=oRs.Fields("GameAbbreviation").Value%> )</TH>
				</TR>
				<%
			End If
		If oRS.Fields ("LadderType").Value = "T" Then
			%>
			<tr bgcolor=<%=bgc%>><td><%=ors.fields("LadderName").value%> Ladder Captains: </TD><TD><input type=radio name=mailto class=borderless value="t<%=ors.fields("LadderID").value%>"></td></tr>
			<% 
		ElseIf oRS.Fields ("LadderType").Value = "P" Then
			%>
			<tr bgcolor=<%=bgc%>><td><%=ors.fields("LadderName").value%> Players: </TD><TD><input type=radio name=mailto class=borderless value="p<%=ors.fields("LadderID").value%>"></td></tr>
			<% 
		ElseIf oRS.Fields ("LadderType").Value = "L" Then
			%>
			<tr bgcolor=<%=bgc%>><td><%=ors.fields("LadderName").value%> League Captains: </TD><TD><input type=radio name=mailto class=borderless value="l<%=ors.fields("LadderID").value%>"></td></tr>
			<% 
		ElseIf oRS.Fields ("LadderType").Value = "A" Then
			%>
			<tr bgcolor=<%=bgc%>><td><%=ors.fields("LadderName").value%> Tournament Captains: </TD><TD><input type=radio name=mailto class=borderless value="a<%=ors.fields("LadderID").value%>"></td></tr>
			<% 
		End If
		if bgc = bgcone then
			bgc = bgctwo
		else 
			bgc = bgcone
		end if
		ors.movenext
	loop
end if
%>
</table>
</TD></TR>
</TABLE>
<%
ors.nextrecordset
bgc = bgcone
%>
</form>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

