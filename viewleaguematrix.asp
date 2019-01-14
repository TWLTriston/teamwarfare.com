<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("League") & ": " & Request.QueryString("Conference") & ": " & Request.QueryString("Division"), """", "&quot;") & " Match Matrix"

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strLeagueName, intLeagueID
Dim intConferenceID, intDivisionID, strConferenceName, strDivisionName
Dim intDivisionsShown, intRank, intLinkID

strLeagueName = Request.QueryString("League")
If Len(Trim(strLeagueName)) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If

strConferenceName = Request.QueryString("Conference")
strDivisionName = Request.QueryString("Division")

strSQL = "SELECT LeagueID, LeagueName FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "' AND LeagueActive = 1"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
	strLeagueName = oRs.Fields("LeagueName").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
oRs.NextRecordSet

If Len(Trim(strConferenceName)) > 0 Then
	strSQL = "SELECT LeagueConferenceID, ConferenceName FROM tbl_league_conferences WHERE ConferenceName= '" & CheckString(strConferenceName) & "' AND LeagueID = '" & intLeagueID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intConferenceID = oRs.Fields("LeagueConferenceID").Value
		strConferenceName = oRs.Fields("ConferenceName").Value
	Else
		oRs.Close
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Response.Clear
		Response.Redirect "errorpage.asp?error=7"
	End If
	oRs.NextRecordSet
End If

If Len(strDivisionName) > 0 Then
	strSQL = "SELECT LeagueDivisionID, DivisionName FROM tbl_league_divisions WHERE DivisionName = '" & CheckString(strDivisionName) & "' AND LeagueID = '" & intLeagueID & "' AND LeagueConferenceID = '" & intConferenceID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intDivisionID = oRs.Fields("LeagueDivisionID").Value
		strDivisionName = oRs.Fields("DivisionName").Value
	Else
		oRs.Close
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Response.Clear
		Response.Redirect "errorpage.asp?error=7"
	End If
	oRs.NextRecordSet
End If

Dim intConferences, intDivisions
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("")

%>
	<script language="javascript" type="text/javascript">
	var arrConferences = new Array(); // new Array(intConferenceID, strConferenceName);
	var arrDivisions = new Array(); // new Array(intConferenceID, intDivsionID, strDivisionName);

	function ConferenceChange() {
		objConferences = eval("self.document.frmMatrix.ConferenceID");
		objDivisions = eval("self.document.frmMatrix.Division");
		objConferenceName = eval("self.document.frmMatrix.Conference");
		strConfName = objConferences.options[objConferences.selectedIndex].text.substr(0, objConferences.options[objConferences.selectedIndex].text.length - 11);
		objConferenceName.value = strConfName;
		objDivisions.length = 0;
		var intOptions = 0;
		for (var i=1;i<arrDivisions.length;i++) {
			if ((arrDivisions[i][0] == objConferences.options[objConferences.selectedIndex].value) || (arrDivisions[i][1].length == 0)|| (objConferences.options[objConferences.selectedIndex].value.length == 0)) {
				objDivisions.options[intOptions] = new Option(arrDivisions[i][2], arrDivisions[i][1]);
				intOptions++;
			}
		}
	}
	<%
	Dim intTopConfID
	strSQL = "SELECT ConferenceName, c.LeagueConferenceID FROM tbl_league_conferences c WHERE c.LeagueID = '" & intLeagueID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intConferences = 0
		intTopConfID = oRs.Fields("LeagueConferenceID").Value
		Do While Not(oRs.EOF)
			intConferences = intConferences + 1
			%>
			arrConferences[<%=intConferences%>] = new Array("<%=oRs.Fields("LeagueConferenceID").Value%>", "<%=Server.HTMLEncode(oRS.Fields("ConferenceName").value)%> Conference");
			<%
			oRs.MoveNext
		Loop
		%>
		var intConferences = <%=intConferences%>;
		<%
	End If
	oRs.NextRecordSet
	
	strSQL = "SELECT d.LeagueConferenceID, LeagueDivisionID, DivisionName FROM tbl_league_divisions d WHERE d.LeagueID = '" & intLeagueID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intDivisions = 0
		Do While Not(oRs.EOF)
			intDivisions = intDivisions + 1
			%>
			arrDivisions[<%=intDivisions%>] = new Array("<%=oRs.Fields("LeagueConferenceID").Value%>", "<%=oRs.Fields("DivisionName").Value%>", "<%=Server.HTMLEncode(oRS.Fields("DivisionName").value)%> Division");
			<%
			oRs.MoveNext
		Loop
		%>
		var intDivisions = <%=intDivisions%>;
		<%
	End If
	oRs.NextRecordSet
	%>
	</script>
<form name="frmMatrix" id="frmMatrix" action="viewleaguematrix.asp" method="get">
<input type="hidden" name="League" id="League" value="<%=Server.HTMLEncode(strLeagueName & "")%>" />
<input type="hidden" name="Conference" id="Conference" value="" />
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
<tr>
<td>
	<table border="0" cellspacing="1" cellpadding="4">
	<tr>
		<th colspan="2" bgcolor="#000000">Conferences and Division</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right" width="100"><b>Conference:</b></td>
		<td bgcolor="<%=bgcone%>">&nbsp;
			<select name="ConferenceID" id="ConferenceID" onchange="ConferenceChange();">
				<option value="">--- Select a Conference ---</option>
				<%
				strSQL = "SELECT ConferenceName, c.LeagueConferenceID FROM tbl_league_conferences c WHERE c.LeagueID = '" & intLeagueID & "'"
				oRs.Open strSQL, oConn
				If Not(oRs.EOF AND oRs.BOF) Then
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & oRs.Fields("LeagueConferenceID").Value & """>" & Server.HTMLEncode(oRs.Fields("ConferenceName").Value) & " Conference</option>" & vbCrlF
						oRs.MoveNext
					Loop
				End If
				oRs.NextRecordSet
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right"><b>Division:</b></td>
		<td bgcolor="<%=bgcone%>">&nbsp;
			<select name="Division" id="Division">
				<option value="">--- Select a Conference ---</option>
			</select>
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#000000" align="center"><input type="submit" value="Change Division" /></td>
	</tr>
	</table>
</td>
</tr>
</table>
</form>

<%
If Len(strDivisionName) > 0 Then
	
	Dim strKey, intHLinkID, intVLinkID
	Dim dicMatches
	Set dicMatches = Server.CreateObject("Scripting.Dictionary")
	Dim dicTeams
	Set dicTeams = Server.CreateObject("Scripting.Dictionary")
	
	strSQL = "SELECT HomeTeamLinkID, VisitorTeamLinkID "
	strSQL = strSQL & " FROM tbl_league_history "
	strSQL = strSQL & " WHERE LeagueID = '" & intLeagueID & "' "
	strSQL = strSQL & " AND LeagueConferenceID = '" & intConferenceID & "' "
	strSQL = strSQL & " AND LeagueDivisionID = '" & intDivisionID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		Do While Not (oRs.EOF)
			intHLinkID = oRs.Fields("HomeTeamLinkID").Value
			intVLinkID = oRs.Fields("VisitorTeamLinkID").Value
			strKey = intHLinkID & intVLinkID
			If Not(dicMatches.Exists (strKey)) Then
				dicMatches.Add (strKey), "P"
			Else
				dicMatches.Item (strKey) = dicMatches.Item(strKey) & "P"
			End If
			strKey = intVLinkID & intHLinkID
			If Not(dicMatches.Exists (strKey)) Then
				dicMatches.Add (strKey), "P"
			Else
				dicMatches.Item (strKey) = dicMatches.Item(strKey) & "P"
			End If
			oRs.Movenext
		Loop
	End If
	ors.NextRecordSet
	
	
	strSQL = "SELECT HomeTeamLinkID, VisitorTeamLinkID "
	strSQL = strSQL & " FROM tbl_league_matches "
	strSQL = strSQL & " WHERE LeagueID = '" & intLeagueID & "' "
	strSQL = strSQL & " AND LeagueConferenceID = '" & intConferenceID & "' "
	strSQL = strSQL & " AND LeagueDivisionID = '" & intDivisionID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		Do While Not (oRs.EOF)
			intHLinkID = oRs.Fields("HomeTeamLinkID").Value
			intVLinkID = oRs.Fields("VisitorTeamLinkID").Value
			strKey = intHLinkID & intVLinkID
			If Not(dicMatches.Exists (strKey)) Then
				dicMatches.Add (strKey), "S"
			Else
				dicMatches.Item (strKey) = dicMatches.Item(strKey) & "S"
			End If
			strKey = intVLinkID & intHLinkID
			If Not(dicMatches.Exists (strKey)) Then
				dicMatches.Add (strKey), "S"
			Else
				dicMatches.Item (strKey) = dicMatches.Item(strKey) & "S"
			End If
			oRs.Movenext
		Loop
	End If
	ors.NextRecordSet
	
	strSQL = "SELECT lnkLeagueTeamID, TeamName, TeamTag, LeaguePoints, Rank, Wins, Losses, Draws, WinPct, RoundsWon, RoundsLost FROM "
	strSQL = strSQL & " lnk_league_team l "
	strSQL = strSQL & " INNER JOIN tbl_teams T "
	strSQL = strSQL & " ON t.TeamID = l.TeamID "
	strSQL = strSQL & " WHERE LeagueDivisionID = '" & intDivisionID & "' "
	strSQL = strSQL & " AND Active = 1 "
	strSQL = strSQL & " ORDER BY LeaguePoints DESC, Rank ASC"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		Do While Not(oRs.EOF)
			dicTeams.Add oRs.FIelds("TeamName").Value & "~|~" & oRs.Fields("TeamTag").Value, oRs.FieldS("lnkLeagueTeamID").Value
			oRs.MoveNext
		Loop
	End if
	oRs.NextRecordSet
	
	Dim dic2Teams
	Set dic2Teams = dicTeams
	Dim intTeams, oItem, oItem2
	Dim strTeamName, strTeamTag, intTLinkID, intT2LinkID
	intTeams = dicTeams.Count
	%>
	<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
	<tr>
		<td>
		<table border="0" cellspacing="1" cellpadding="4">
		<tr>
			<th colspan="<%=intTeams+1%>" bgcolor="#000000">Match Matrix</th>
		</tr>
		<tr>
			<td bgcolor="#000000">&nbsp;</td>
		<%
		dim i
		For Each oItem In dicTeams
			strTeamName = Left(oItem, inStr(oItem, "~|~") - 1)
			strTeamTag =  Right(oItem, Len(oItem) - inStr(oItem, "~|~")- 2)
			%>
			<td bgcolor="#000000" width="25" align="center"><a href="viewteam.asp?team=<%=Server.URLEncode(strTeamName & "")%>"><%
				For i = 1 to len(strTeamTag)
					Response.write mid(Server.HTMLencode(strTeamTag), i, 1) & "<br />"
				Next 
				%></a></td>
			<%
		Next
		%>
		</tr>
		<%
		For Each oItem In dicTeams
			strTeamName = Left(oItem, inStr(oItem, "~|~") - 1)
			strTeamTag =  Right(oItem, Len(oItem) - inStr(oItem, "~|~") - 2)
			intTLinkID = dicTeams.Item(oItem)
			%>
			<tr>
				<td bgcolor="#000000"><a href="viewteam.asp?team=<%=Server.URLEncode(strTeamName & "")%>"><%=Server.HTMLencode(strTeamName & " - " & strTeamTag & "")%></a></td>
				<%
				For Each oItem2 In dic2Teams
					intT2LinkID = dic2Teams.Item(oItem2)
					If intTLinkID = intT2LinkID Then
						Response.Write "<td bgcolor=""" & bgcone & """>&nbsp;</td>"
					Else
						%>
						<td bgcolor="#000000" align="center"><%
						Response.Write dicMatches.Item(intTLinkID & intT2LinkID)
						%></td>
					<%
					End If
				Next
				%>
			</tr>	
			<%
		Next
		%>
		</table>
		</td>
	</tr>
	</table>
	<br /><br />
	P = Played a match<br />
	S = Match currently scheduled
	<%
End If
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>