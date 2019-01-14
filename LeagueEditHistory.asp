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

Dim intHistoryID, strFromURL
intHistoryID = Request.QueryString("HistoryID")

If Not(IsNumeric(intHistoryID)) Then
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
Dim intMapVisitorScore(6), intMapHomeScore(6)
strFromURL = Request.QueryString("f")
If Len(strFromURL) = 0 Then
	strFromURL = "/default.asp"
End If

strSQL = "SELECT 'HomeTeamName' = th.TeamName, 'VisitorTeamName' = tv.TeamName,"
strSQL = strSQL & "	LeagueName,"
strSQL = strSQL & "	lh.LeagueID, HomeTeamLinkID, VisitorTeamLinkID,"
strSQL = strSQL & "	Map1, Map1HomeScore, Map1VisitorScore,"
strSQL = strSQL & "	Map2, Map2HomeScore, Map2VisitorScore,"
strSQL = strSQL & "	Map3, Map3HomeScore, Map3VisitorScore,"
strSQL = strSQL & "	Map4, Map4HomeScore, Map4VisitorScore,"
strSQL = strSQL & "	Map5, Map5HomeScore, Map5VisitorScore,"
strSQL = strSQL & "	MatchDate"
strSQL = strSQL & "	FROM tbl_league_history lh"
strSQL = strSQL & "	INNER JOIN lnk_league_team lth ON HomeTeamLinkID = lth.lnkLeagueTeamID"
strSQL = strSQL & "	INNER JOIN lnk_league_team ltv ON VisitorTeamLinkID = ltv.lnkLeagueTeamID"
strSQL = strSQL & "	INNER JOIN tbl_teams th ON lth.TeamID = th.TeamID"
strSQL = strSQL & "	INNER JOIN tbl_teams tv ON ltv.TeamID = tv.TeamID"
strSQL = strSQL & "	INNER JOIN tbl_leagues l ON l.LeagueID = lh.LeagueID"
strSQL = strSQL & "	WHERE LeagueHistoryID = '" & CheckString(intHistoryID) & "'"
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
	intMapHomeScore(1) = oRs.Fields("Map1HomeScore").Value
	intMapVisitorScore(1) = oRs.Fields("Map1VisitorScore").Value
	intMapHomeScore(2) = oRs.Fields("Map2HomeScore").Value
	intMapVisitorScore(2) = oRs.Fields("Map2VisitorScore").Value
	intMapHomeScore(3) = oRs.Fields("Map3HomeScore").Value
	intMapVisitorScore(3) = oRs.Fields("Map3VisitorScore").Value
	intMapHomeScore(4) = oRs.Fields("Map4HomeScore").Value
	intMapVisitorScore(4) = oRs.Fields("Map4VisitorScore").Value
	intMapHomeScore(5) = oRs.Fields("Map5HomeScore").Value
	intMapVisitorScore(5) = oRs.Fields("Map5VisitorScore").Value
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
Call ContentStart("Edit History on " & Server.HTMLEncode(strLeagueName & "") & " League")
%>

<form name="frmLeagueEditHistory" id="frmLeagueEditHistory" action="SaveLeague.asp" method="post">
<input type="hidden" name="HistoryID" id="HistoryID" value="<%=intHistoryID%>" />
<input type="hidden" name="LeagueID" id="LeagueID" value="<%=intLeagueID%>" />
<input type="hidden" name="SaveType" id="SaveType" value="LeagueEditHistory" />
<input type="hidden" name="FromURL" id="FromURL" value="<%=Server.HTMLEncode(strFromURL & "")%>" />


<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center" WIDTH="75%">
<tr><td>
    <table width="100%" border="0" cellpadding="4" cellspacing="1">
    <tr>
    	<th colspan="2" bgcolor="#000000">Edit History</th>
    </tr>
    <tr>
    	<td align="right" bgcolor="<%=bgcone%>" width="70%">Match Date:</td>
    	<td bgcolor="<%=bgcone%>"><%=strMatchDate%></td>
    </tr>
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
    			<td bgcolor="<%=bgctwo%>"><input name="HMapScore<%=i%>" id="HMapScore<%=i%>" maxlength="4" size="5" type="text" value="<%=intMapHomeScore(i)%>" /></td>
    		</tr>
    		<tr>
    			<td align="right" bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(strVisitorName & "")%> Round Points:</td>
    			<td bgcolor="<%=bgctwo%>"><input name="VMapScore<%=i%>" id="VMapScore<%=i%>" maxlength="4" size="5" type="text" value="<%=intMapVisitorScore(i)%>" /></td>
    		</tr>
    		<%
    	End If
		Next
    %>
    <tr>
    	<td colspan="2" align="center" bgcolor="#000000"><input type="submit" value="Save History"></td>
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