<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Report League Match"

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

Dim intMatchID, strFromURL
intMatchID = Request.QueryString("MatchID")
strFromURL = Trim(Request.QueryString("f"))
If Len(strFromURL) = 0 Then
	strFromURL = "default.asp"
End If

If Not(IsNumeric(intMatchID)) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing	
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End if

Dim intHomeLinkID, intVisitorLinkID, strMatchDate, strMap1, strMap2, strMap3, strMap4, strMap5
Dim intLeagueID, intConferenceID, intDivisionID
Dim strMaps(6)
Dim strHomeName, strVisitorName
Dim strDivisionName, strLeagueName, strConferenceName

strSQL = "SELECT HomeTeamLinkID, VisitorTeamLinkID, MatchDate, "
strSQL = strSQL & " Map1, Map2, Map3, Map4, Map5, "
strSQL = strSQL & " LeagueID, LeagueConferenceID, LeagueDivisionID "
strSQL = strSQL & " FROM tbl_league_matches "
strSQL = strSQL & " WHERE LeagueMatchID = '" & intMatchID & "'"
'Response.Write strSQL & "<br />"
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
	intConferenceID = oRs.Fields("LeagueConferenceID").Value
	intDivisionID = oRs.Fields("LeagueDivisionID").Value
Else
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect strFromURL
End if
oRS.NextRecordSet

if not(bSysAdmin OR IsLeagueTeamCaptainByLinkID(intHomeLinkId) OR IsLeagueTeamCaptainByLinkID(intVisitorLinkId) OR IsLeagueAdminByID(intLeagueID)) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If

strSQL = "SELECT TeamName FROM tbl_teams t "
strSQL = strSQL & " INNER JOIN lnk_league_team l "
strSQL = strSQL & " ON l.TeamID = t.TeamID "
strSQL = strSQL & " WHERE l.lnkLeagueTeamID = '" & intHomeLinkID & "'"
'Response.Write strSQL
oRS.Open strSQL, oConn
If Not(ors.eof and ors.bof) Then
	strHomeName = oRS.Fields("TeamName").Value	
Else
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.End
	Response.Redirect "errorpage.asp?error=7"
End if
oRS.NextRecordSet
 
strSQL = "SELECT TeamName FROM tbl_teams t "
strSQL = strSQL & " INNER JOIN lnk_league_team l "
strSQL = strSQL & " ON l.TeamID = t.TeamID "
strSQL = strSQL & " WHERE l.lnkLeagueTeamID = '" & intVisitorLinkID & "'"
oRS.Open strSQL, oConn
If Not(ors.eof and ors.bof) Then
	strVisitorName = oRS.Fields("TeamName").Value	
Else
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	'Response.Write "Who the fuck is the visitor"
	Response.Redirect "errorpage.asp?error=7"
End If
oRS.NextRecordSet

Dim intScoring
If intDivisionID = 0 Then
	' Interconference game...
	If intConferenceID = 0 Then
		' interleague game
		strSQL = "SELECT LeagueName, l.Scoring, '', '' FROM tbl_leagues l "
		strSQL = strSQL & " WHERE l.LeagueID = '" & intLeagueID & "'"		
	Else
		strSQL = "SELECT l.LeagueName, l.Scoring, c.ConferenceName, '' FROM tbl_leagues l "
		strSQL = strSQL & " INNER JOIN tbl_league_conferences c ON c.LeagueID = l.LeagueID "		
		strSQL = strSQL & " WHERE c.LeagueConferenceID = '" & intConferenceID & "'"		
	End If
Else 
	strSQL = "SELECT l.LeagueName, l.Scoring, c.ConferenceName, d.DivisionName FROM tbl_leagues l "
	strSQL = strSQL & " INNER JOIN tbl_league_conferences c ON c.LeagueID = l.LeagueID "		
	strSQL = strSQL & " INNER JOIN tbl_league_divisions d ON d.LeagueConferenceID = c.LeagueConferenceID "		
	strSQL = strSQL & " WHERE d.LeagueDivisionID = '" & intDivisionID & "'"		
End If
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	strLeagueName = oRS.FIelds(0).Value	
	intScoring = oRs.Fields("Scoring").Value
	strConferenceName = oRS.FIelds(2).Value	
	strDivisionName = oRS.FIelds(3).Value	
End If 
oRs.NextRecordSet%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Report Match on " & Server.HTMLEncode(strLeagueName & "") & " League")
%>

<form name="frmLeagueReportMatch" id="frmLeagueReportMatch" action="LeagueReportMatchValidate.asp" method="post">
<input type="hidden" name="MatchID" id="MatchID" value="<%=intMatchID%>" />
<input type="hidden" name="SaveType" id="SaveType" value="LeagueReportMatch" />
<input type="hidden" name="FromURL" id="FromURL" value="<%=Server.HTMLEncode(strFromURL & "")%>" />

<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center" WIDTH="75%">
<tr><td>
    <table width="100%" border="0" cellpadding="4" cellspacing="1">
    <tr>
    	<th colspan="2" bgcolor="#000000">Report Match Results</th>
    </tr>
    <tr>
    	<td align="right" bgcolor="<%=bgcone%>" width="70%">Match Date:</td>
    	<td bgcolor="<%=bgcone%>"><%=strMatchDate%></td>
    </tr>
    <tr>
    	<td colspan="2" bgcolor="#000000">If either team did not show up for the match, check one of the boxes below.</td>
    </tr>
    <% If intScoring = 0 Then %>
    <tr>
    	<td align="right" bgcolor="<%=bgcone%>" width="70%">Did <%=strHomeName%> fail to show up for match? (check for yes):</td>
    	<td bgcolor="<%=bgcone%>"><input type="checkbox" name="HNoShow" id="HNoShow" value="1" /></td>
    </tr>
    <tr>
    	<td align="right" bgcolor="<%=bgcone%>" width="70%">Did <%=strVisitorName%> fail to show up for match? (check for yes):</td>
    	<td bgcolor="<%=bgcone%>"><input type="checkbox" name="VNoShow" id="VNoShow" value="1" /></td>
    </tr>
    <tr>
    	<td colspan="2" bgcolor="#000000">Otherwise, type in the round points for each team, on the map / all maps played.</td>
    </tr>
    <% End If %>
    <%
    Dim i
    For i = 1 to 5
    	If Len(Trim(strMaps(i))) > 0 AND Not(IsNull(strMaps(i))) Then
    		%>
    		<tr>
    			<td colspan="2" bgcolor="#000000"><%=strMaps(i)%></td>
    		</tr>
    		<tr>
    			<td align="right" bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(strHomeName & "")%> Round Points:</td>
    			<td bgcolor="<%=bgctwo%>"><input name="HMapScore<%=i%>" id="HMapScore<%=i%>" maxlength="4" size="5" type="text" value="0" /></td>
    		</tr>
    		<tr>
    			<td align="right" bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(strVisitorName & "")%> Round Points:</td>
    			<td bgcolor="<%=bgctwo%>"><input name="VMapScore<%=i%>" id="VMapScore<%=i%>" maxlength="4" size="5" type="text" value="0" /></td>
    		</tr>
    		<% If intScoring = 1 Then %>
		    <tr>
		    	<td align="right" bgcolor="<%=bgcone%>" width="70%">Did <%=strHomeName%> fail to show up for this map? (check for yes):</td>
		    	<td bgcolor="<%=bgcone%>"><input type="checkbox" name="HMap<%=i%>NoShow" id="HMap<%=1%>NoShow" value="1" /></td>
		    </tr>
		    <tr>
		    	<td align="right" bgcolor="<%=bgcone%>" width="70%">Did <%=strVisitorName%> fail to show up for this map? (check for yes):</td>
		    	<td bgcolor="<%=bgcone%>"><input type="checkbox" name="VMap<%=i%>NoShow" id="VMap<%=1%>NoShow" value="1" /></td>
		    </tr>
		    <% End If %>    		
    		<%
    	End If
	Next
    %>
    <tr>
    	<td colspan="2" align="center" bgcolor="#000000"><input type="submit" value="Report Results"></td>
    </tr>
  	</table>
</td></tr>
</table>
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