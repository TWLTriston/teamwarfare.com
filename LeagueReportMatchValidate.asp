<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Report League Match Validation"

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
intMatchID = Request.Form("MatchID")
strFromURL = Trim(Request.Form("FromURL"))
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
		strSQL = "SELECT LeagueName, '', '', l.Scoring, l.WinPoints, l.LossPoints, l.DrawPoints, l.NoShowPoints FROM tbl_leagues l "
		strSQL = strSQL & " WHERE l.LeagueID = '" & intLeagueID & "'"		
	Else
		strSQL = "SELECT l.LeagueName, c.ConferenceName, '', l.Scoring, l.WinPoints, l.LossPoints, l.DrawPoints, l.NoShowPoints FROM tbl_leagues l "
		strSQL = strSQL & " INNER JOIN tbl_league_conferences c ON c.LeagueID = l.LeagueID "		
		strSQL = strSQL & " WHERE c.LeagueConferenceID = '" & intConferenceID & "'"		
	End If
Else 
	strSQL = "SELECT l.LeagueName, c.ConferenceName, d.DivisionName, l.Scoring, l.WinPoints, l.LossPoints, l.DrawPoints, l.NoShowPoints FROM tbl_leagues l "
	strSQL = strSQL & " INNER JOIN tbl_league_conferences c ON c.LeagueID = l.LeagueID "		
	strSQL = strSQL & " INNER JOIN tbl_league_divisions d ON d.LeagueConferenceID = c.LeagueConferenceID "		
	strSQL = strSQL & " WHERE d.LeagueDivisionID = '" & intDivisionID & "'"		
End If
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	strLeagueName = oRS.FIelds(0).Value	
	strConferenceName = oRS.FIelds(1).Value	
	strDivisionName = oRS.FIelds(2).Value	
	intScoring = oRs.Fields(3).Value
End If 
oRs.NextRecordSet

Dim strHScores(6), strVScores(6)
Dim intHScore, intVScore
Dim i
Dim strOutcome
Dim intHMapWins, intVMapWins, intVMapLosses, intHMapLosses, intHMapDraws, intVMapDraws, intVMapNoShows, intHMapNoShows

If intScoring = 0 Then
	' Add up all the rounds, then find a winner from teh total	
	intHScore = 0 
	intVScore = 0
	strOutcome = "Unknown."
	For i = 1 to 5
		strHScores(i) = 0
		strVScores(i) = 0
	Next 
	If Request.Form("HNoShow") = "1" Then
		If Request.Form("VNoShow") = "1" Then
			strOutcome = "Both teams forfeit for failure to show."
		Else
			strOutcome = strHomeName & " forfeits for failure to show."
		End If
	Else
		If Request.Form("VNoShow") = "1" Then
			strOutcome = strVisitorName  & " forfeits for failure to show."
		Else
			For i = 1 to 5
			   	If Len(Trim(strMaps(i))) > 0 AND Not(IsNull(strMaps(i))) Then
					strHScores(i) = Request.Form("HMapScore" & i)
					strVScores(i) = Request.Form("VMapScore" & i)
					If IsNumeric(strHScores(i)) Then
						intHScore = intHScore + strHScores(i)
					End If
					If IsNumeric(strVScores(i)) Then
						intVScore = intVScore + strVScores(i)
					End If
				End If
			Next	
			If intHScore > intVScore Then
				strOutcome = strHomeName & " wins. "
			ElseIf intHScore = intVScore Then
				strOutcome = "Tie game."
			Else
				strOutCome = strVisitorName & " wins."
			End If
		End If
	End If
ElseIf intScoring = 1 Then
	' Each map is different
	intHMapWins = 0 
	intVMapWins = 0
	intVMapLosses = 0
	intHMapLosses = 0
	intHMapDraws = 0
	intVMapDraws = 0
	intHMapNoShows = 0 
	intVMapNoShows = 0
	For i = 1 to 5
		If Len(Trim(strMaps(i))) > 0 AND Not(IsNull(strMaps(i))) Then
			strHScores(i) = Request.Form("HMapScore" & i)
			strVScores(i) = Request.Form("VMapScore" & i)
			If IsNumeric(strHScores(i)) Then
				strHScores(i) = cint(strHScores(i))
			End If
			If IsNumeric(strVScores(i)) Then
				strVScores(i) = Cint(strVScores(i))
			End If
			
			If Request.Form("HMap" & i & "NoShow") = "1" AND Request.Form("VMap" & i & "NoShow") = "1" Then
				intHMapNoShows = intHMapNoShows + 1
				intVMapNoShows = intVMapNoShows + 1
				strOutCome = strOutCome & "Both teams forfeit " & strMaps(i) & "<br />"
			ElseIf Request.Form("HMap" & i & "NoShow") = "1" Then
				intHMapNoShows = intHMapNoShows + 1
				intVMapWins = intVMapWins + 1
				strOutCome = strOutCome & strHomeName & " forfeits " & strMaps(i) & "<br />"
			ElseIf Request.Form("VMap" & i & "NoShow") = "1" Then
				intHMapWins = intHMapWins + 1
				intVMapNoShows = intVMapNoShows + 1
				strOutCome = strOutCome & strVisitorName & " forfeits " & strMaps(i) & "<br />"
			ElseIf strVScores(i) > strHScores(i) Then
				intHMapLosses = intHMapLosses + 1
				intVMapWins = intVMapWins + 1
				strOutCome = strOutCome & strVisitorName & " wins " & strMaps(i) & "<br />"
			ElseIf strVScores(i) < strHScores(i) Then
				intHMapWins = intHMapWins + 1
				intVMapLosses = intVMapLosses + 1
				strOutCome = strOutCome & strHomeName & " wins " & strMaps(i) & "<br />"
			ElseIf strVScores(i) = strHScores(i) Then
				intHMapDraws = intHMapDraws + 1
				intVMapDraws = intVMapDraws + 1
				strOutCome = strOutCome & "Draw on " & strMaps(i) & "<br />"
			End If
		End If
	Next
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Validate Results for match on " & Server.HTMLEncode(strLeagueName & "") & " League")
%>

<form name="frmLeagueReportMatch" id="frmLeagueReportMatch" action="SaveItem.asp" method="post">
<input type="hidden" name="MatchID" id="MatchID" value="<%=intMatchID%>" />
<input type="hidden" name="SaveType" id="SaveType" value="LeagueReportMatch" />
<input type="hidden" name="FromURL" id="FromURL" value="<%=Server.HTMLEncode(strFromURL & "")%>" />
<input type="hidden" name="HNoShow" id="HNoShow" value="<%=Request.Form("HNoShow")%>" />
<input type="hidden" name="VNoShow" id="VNoShow" value="<%=Request.Form("VNoShow")%>" />

<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center" WIDTH="50%">
<tr><td>
    <table width="100%" border="0" cellpadding="4" cellspacing="1">
    <tr>
    	<th colspan="2" bgcolor="#000000">Validate Match Results</th>
    </tr>
    <tr>
    	<th colspan="2" bgcolor="#000000"><%=strOutcome%></th>
    </tr>
    <tr>
    	<td align="right" bgcolor="<%=bgcone%>" width="70%">Match Date:</td>
    	<td bgcolor="<%=bgcone%>"><%=strMatchDate%></td>
    </tr>
    <% If Request.Form("HNoShow") = "1" Then %>
    <tr>
    	<td align="center" bgcolor="<%=bgcone%>" colspan="2"><%=strHomeName%> failed to show up for match.</td>
    </tr>
    <% End If %>
    <% If Request.Form("VNoShow") = "1" Then %>
    <tr>
    	<td align="center" bgcolor="<%=bgcone%>" colspan="2"><%=strVisitorName%> failed to show up for match.</td>
    </tr>
    <% End If %>
    <%
    For i = 1 to 5
    	If Len(Trim(strMaps(i))) > 0 AND Not(IsNull(strMaps(i))) Then
    		%>
    		<tr>
    			<td colspan="2" bgcolor="#000000"><%=strMaps(i)%></td>
    		</tr>
    		<tr>
    			<td align="right" bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(strHomeName & "")%> Round Points:</td>
    			<td bgcolor="<%=bgctwo%>"><input name="HMapScore<%=i%>" id="HMapScore<%=i%>" maxlength="3" size="5" type="hidden" value="<%=strHScores(i)%>" /><%=strHScores(i)%></td>
    		</tr>
    		<tr>
    			<td align="right" bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(strVisitorName & "")%> Round Points:</td>
    			<td bgcolor="<%=bgctwo%>"><input name="VMapScore<%=i%>" id="VMapScore<%=i%>" maxlength="3" size="5" type="hidden" value="<%=strVScores(i)%>" /><%=strVScores(i)%></td>
    		</tr>
    		<input type=hidden name="VMap<%=i%>NoShow" id="VMap<%=i%>NoShow" value="<%=Request.Form("VMap" & i & "NoShow")%>">
    		<input type=hidden name="HMap<%=i%>NoShow" id="HMap<%=i%>NoShow" value="<%=Request.Form("HMap" & i & "NoShow")%>">
    		<%
    	End If
	Next
    %>
    <tr>
    	<td colspan="2" align="center" bgcolor="#000000"><input type="button" value="Go Back" onclick="GoBack();" />&nbsp;&nbsp;&nbsp;<input type="submit" value="Report Results" /></td>
    </tr>
  	</table>
</td></tr>
</table>
</form>
<script language="javascript">
<!--
function GoBack() {
	window.location="LeagueReportMatch.asp?matchid=<%=intMatchID%>&f=<%=Server.URLEncode(strFromURL)%>";
}
//-->
</script>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>