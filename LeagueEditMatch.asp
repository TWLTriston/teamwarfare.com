<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Edit League Match"

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

Dim intLeagueMatchID, strFromURL
intLeagueMatchID = Request.QueryString("LeagueMatchID")

If Not(IsNumeric(intLeagueMatchID)) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing	
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End if

Dim intHomeLinkID, intVisitorLinkID, strMatchDate, strMap1, strMap2, strMap3, strMap4, strMap5
Dim intLeagueID
Dim strMaps(6)
Dim strHomeName, strVisitorName
Dim strDivisionName, strLeagueName
'strFromURL = Request.QueryString("f")
strFromURL = "/LeagueMatches.asp"


strSQL = "SELECT 'HomeTeamName' = th.TeamName, 'VisitorTeamName' = tv.TeamName,"
strSQL = strSQL & "	LeagueName,"
strSQL = strSQL & "	LeagueMatchID, lm.LeagueID, HomeTeamLinkID, VisitorTeamLinkID,"
strSQL = strSQL & "	Map1, Map2, Map3, Map4, Map5,"
strSQL = strSQL & "	MatchDate"
strSQL = strSQL & "	FROM tbl_league_matches lm"
strSQL = strSQL & "	INNER JOIN lnk_league_team lth ON HomeTeamLinkID = lth.lnkLeagueTeamID"
strSQL = strSQL & "	INNER JOIN lnk_league_team ltv ON VisitorTeamLinkID = ltv.lnkLeagueTeamID"
strSQL = strSQL & "	INNER JOIN tbl_teams th ON lth.TeamID = th.TeamID"
strSQL = strSQL & "	INNER JOIN tbl_teams tv ON ltv.TeamID = tv.TeamID"
strSQL = strSQL & "	INNER JOIN tbl_leagues l ON l.LeagueID = lm.LeagueID"
strSQL = strSQL & "	WHERE LeagueMatchID = '" & CheckString(intLeagueMatchID) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intHomeLinkID = oRs.Fields("HomeTeamLinkID").Value
	intVisitorLinkID = oRs.Fields("VisitorTeamLinkID").Value
	strMatchDate = oRs.Fields("MatchDate").Value
	strMaps(1) = oRs.Fields("Map1").Value
	strMaps(2) = oRs.Fields("Map2").Value
	strMaps(3) = oRs.Fields("Map3").Value
	strMaps(4) = oRs.Fields("Map4").Value
	strMaps(5) = oRs.Fields("Map5").Value
	intLeagueID = oRs.Fields("LeagueID").Value
	strMatchDate = oRs.Fields("MatchDate").Value
	strHomeName = oRs.Fields("HomeTeamName").Value
	strVisitorName = oRs.Fields("VisitorTeamName").Value
	strLeagueName = oRs.Fields("LeagueName").Value
Else
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End if
oRS.NextRecordSet

if not(bSysAdmin OR IsLeagueAdminByID(intLeagueID)) Then
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
Call ContentStart("Edit Match on " & Server.HTMLEncode(strLeagueName & "") & " League")
%>

<form name="frmLeagueEditMatch" id="frmLeagueEditMatch" action="SaveLeague.asp" method="post">
<input type="hidden" name="LeagueMatchID" id="LeagueMatchID" value="<%=intLeagueMatchID%>" />
<input type="hidden" name="LeagueID" id="LeagueID" value="<%=intLeagueID%>" />
<input type="hidden" name="SaveType" id="SaveType" value="LeagueEditMatch" />
<input type="hidden" name="FromURL" id="FromURL" value="<%=Server.HTMLEncode(strFromURL & "")%>" />


<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center" WIDTH="75%">
<tr><td>
    <table width="100%" border="0" cellpadding="4" cellspacing="1">
	<tr> 
		<td colspan="2" bgcolor="#000000">Edit Match</td>
	</tr>
	<tr>
		<th bgcolor="#000000">Current Values</th>
		<th bgcolor="#000000">Change to</th>
	</tr>
    <tr>
    	<td align="Left" bgcolor="<%=bgcone%>" width="70%">Match Date:<%=strMatchDate%></td>
		<td bgcolor="<%=bgctwo%>"><input type="text" name="HMatchDate" id="HMatchDate" value="<%=FormatDateTime(strMatchDate, 2)%>" maxlength="10" /></td>
    </tr>
    <tr>
    	<td colspan="2" align="center" bgcolor="#000000"><input type="submit" value="Save Changes"></td>
    </tr>
  	</table>
</td></tr>
</table>
</form>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>