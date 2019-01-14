<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Edit Tournament Content"

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

Dim X, tournamentid

If Not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim strTournamentName, intTournamentID, intForumID, intGameID, intActive, intLocked, intSignup
Dim strRulesName, intRosterLock, strHeaderURL, strContentMain, strContentPrizes, strContentSponsors, intHasPrizes, intHasSponsors

Dim strPage, strContent
strPage = Request.QueryString("page")
If Len(strPage) = 0 Then
	strPage = "main"
End If

strTournamentName = Request.QuerySTring("Tournament")
strSQL = "SELECT TournamentID, Signup, Locked, Active, RosterLock, GameID, ForumID, RulesName, HeaderURL, HasSponsors, HasPrizes, ContentMain, ContentPrizes, ContentSponsors, ContentSchedule "
strSQL = strSQL & " FROM tbl_tournaments WHERE TournamentName = '" & CheckString(strTournamentName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intTournamentID = oRs.Fields("TournamentID").Value '
	intForumID = oRs.Fields("ForumID").Value '
	intGameID = oRs.Fields("GameID").Value '
	intActive = oRs.Fields("Active").Value '
	intLocked = oRs.Fields("Locked").Value '
	intSignup = oRs.Fields("Signup").Value ' 
	strRulesName = oRs.Fields("RulesName").Value '
	intRosterLock = oRs.Fields("RosterLock").Value '
	strHeaderURL = oRs.Fields("HeaderURL").Value
	intHasPrizes = oRs.Fields("HasPrizes").Value
	intHasSponsors = oRs.Fields("HasSponsors").Value	
	strContent = oRs.Fields("Content" & strPage).Value
End If
oRs.NextRecordSet

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("Edit Content for " & strTournamentName & " Tournament")
%>
Choose a content:<br />
<a href="EditContent.asp?Tournament=<%=Server.URLEncode(strTournamentName)%>&page=main">Main</a> / 
<a href="EditContent.asp?Tournament=<%=Server.URLEncode(strTournamentName)%>&page=prizes">Prizes</a> / 
<a href="EditContent.asp?Tournament=<%=Server.URLEncode(strTournamentName)%>&page=sponsors">Sponsors</a> / 
<a href="EditContent.asp?Tournament=<%=Server.URLEncode(strTournamentName)%>&page=schedule">Schedule</a>

<script language="javascript">
<!--
	function writeheadline(strData) {
		divheadline.innerHTML = strData
	}
	function writecontent(strData) {
		divcontent.innerHTML = strData
	}
//-->
</script>
<table border=0 cellspacing=0 cellpadding=0 align=center bgcolor="#444444">
<tr><td>
	<table align=center width=400 border=0 cellspacing=1 cellpadding=4>
	<form name=frmEditContent action=saveTournament.asp method="post">
	<input type="hidden" name="SaveType" value="EditContent" />
	<input type="hidden" name="ContentName" value="<%=strPage%>" />
	<input type="hidden" name="TournamentID" value="<%=intTournamentID%>" />
	<input type="hidden" name="TournamentName" value="<%=Server.HTMLEncode(strTournamentName)%>" />
	<tr bgcolor="#000000"><th colspan=2>Edit <%=strPage%> Content</th></tr>
	<tr bgcolor=<%=bgctwo%> height=200><td align=center><textarea name="Content" id="Content" cols=80 rows=15><%=Server.HTMLEncode(strContent & "")%></textarea></td></tr>
	<tr bgcolor=<%=bgctwo%>><td align=center>
		<input type=button name=Preview value="Preview" class=bright onclick="javascript:writecontent(this.form.Content.value);">&nbsp;&nbsp;&nbsp;
		<input type=submit name=submit1 value=submit class=bright>
	</td></tr>
	</form>
	</table>
</td></tr>
</table>
<%
Call ContentEnd()

Call ContentStart("Preview Content")
%>
	<table width="97%" align=center border=0 cellpadding="0">
	  <tr>
		<td><div name="divcontent" id="divcontent"></div></td>
	  </tr>
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