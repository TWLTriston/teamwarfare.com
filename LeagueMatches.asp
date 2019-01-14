<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: League Create Matches"

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin, bAnyLeagueAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
bAnyLeagueAdmin = IsAnyLeagueAdmin()

Dim strLeagueName, intLeagueID, intDivisions, intConferences, intLinkID, intTeams
Dim dtmMatchDate, strHomeTeamName, strVisitorTeamName, strDivisionName, strConferenceName
dtmMatchDate = Request.QueryString("matchDate")
If IsDate(dtmMatchDate) Then
	dtmMatchDate = cDate(dtmMatchDate)
Else
	dtmMatchDate = Now()
End If

if not(bSysAdmin or bAnyLeagueAdmin) Then
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
Call ContentStart("")
%>
<table border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#444444">
<tr>
	<td>
	<table border="0" cellspacing="1" cellpadding="4" width="400" align="center">
	<tr>
		<th bgcolor="#000000">Select a League</th>
	</tr>
<%
If bSysAdmin Then
	strSQL = "SELECT LeagueName, LeagueID FROM tbl_leagues WHERE LeagueActive = 1 "
Else
	strSQL = "SELECT LeagueName, l.LeagueID FROM tbl_leagues l INNER JOIN lnk_league_admin lnk ON lnk.LeagueID = l.LeagueID WHERE lnk.PlayerID='" & Session("PlayerID") & "' AND LeagueActive = 1 "
End If
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	bgc = bgcone
	Do While Not(oRs.EOF)
		%>
		<tr>
			<td bgcolor="<%=bgc%>"><a href="leaguematches.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value)%>"><%=Server.HTMLEncode(oRS.FIelds("LeagueName").value)%> League</a></td>
		</tr>
		<%
		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		oRs.MoveNext
	loop
End if
oRs.NextRecordSet
%>
	</table>
	</td>
</tr></table>
<%
Call ContentEnd()

strLeagueName = trim(Request.Querystring("League"))
strSQL = "SELECT LeagueID FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "' AND LeagueActive = 1"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
End If
oRs.NextRecordSet
If Len(intLeagueID) > 0  AND (bSYsAdmin Or IsLeagueAdminById(intLeagueID)) Then
	Call ContentStart(strLeagueName & " League Create Matches")
	%>
	<script language="javascript" type="text/javascript">
	var arrConferences = new Array(); // new Array(intConferenceID, strConferenceName);
	var arrDivisions = new Array(); // new Array(intConferenceID, intDivsionID, strDivisionName);
	var arrTeams = new Array(); // new Array(intConferenceID, intDivsionID, intLeagueTeamID, strTeamName);

	function ChangeDate() {
		window.location="leaguematches.asp?league=<%=Server.URLEncode(strLeagueName)%>&matchdate="+self.document.frmLeagueMatches.txtMatchDate.value;
	}
	function DeleteLeagueMatch(intMatchID) {
		if (confirm("Are you sure you want to delete this match?")) {
			window.location = "saveitem.asp?league=<%=Server.URLEncode(strLeagueName)%>&matchdate=<%=Server.URLEncode(dtmMatchDate)%>&savetype=LeagueDeleteMatchAdmin&matchid="+intMatchID;
		}
	}

	function ConferenceChange(strPosition) {
		objConferences = eval("self.document.frmLeagueMatches.sel"+strPosition+"ConferenceID");
		objDivisions = eval("self.document.frmLeagueMatches.sel"+strPosition+"DivisionID");
		objTeams = eval("self.document.frmLeagueMatches.sel"+strPosition+"LeagueTeamID");
		
		objDivisions.length = 0;
		var intOptions = 0;
		for (var i=0;i<arrDivisions.length;i++) {
			if ((arrDivisions[i][0] == objConferences.options[objConferences.selectedIndex].value) || (arrDivisions[i][1].length == 0)|| (objConferences.options[objConferences.selectedIndex].value.length == 0)) {
				objDivisions.options[intOptions] = new Option(arrDivisions[i][2], arrDivisions[i][1]);
				intOptions++;
			}
		}
		DivisionChange(strPosition);
	}

	function DivisionChange(strPosition) {
		objConferences = eval("self.document.frmLeagueMatches.sel"+strPosition+"ConferenceID");
		objDivisions = eval("self.document.frmLeagueMatches.sel"+strPosition+"DivisionID");
		objTeams = eval("self.document.frmLeagueMatches.sel"+strPosition+"LeagueTeamID");

		objTeams.length = 0;
		var intOptions = 0;
		for (var i=0;i<arrTeams.length;i++) {
			if ((arrTeams[i][1] == objDivisions.options[objDivisions.selectedIndex].value) 
					|| (arrTeams[i][1].length == 0) 
					|| ((objConferences.options[objConferences.selectedIndex].value.length == 0) && (objDivisions.options[objDivisions.selectedIndex].value.length == 0))
					|| ((objDivisions.options[objDivisions.selectedIndex].value.length == 0) && (arrTeams[i][0] == objConferences.options[objConferences.selectedIndex].value))) {
				objTeams.options[intOptions] = new Option(arrTeams[i][3], arrTeams[i][2]);
				intOptions++;
			}
		}		
	}
	arrConferences[0] = new Array("", "All Conferences");
	arrDivisions[0] = new Array("", "", "All Divisions");
	arrTeams[0] = new Array("", "", "", "Select a Team");
   
	<%
	strSQL = "SELECT ConferenceName, c.LeagueConferenceID FROM tbl_league_conferences c WHERE c.LeagueID = '" & intLeagueID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intConferences = 0
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
			arrDivisions[<%=intDivisions%>] = new Array("<%=oRs.Fields("LeagueConferenceID").Value%>", <%=oRs.Fields("LeagueDivisionID").Value%>, "<%=Server.HTMLEncode(oRS.Fields("DivisionName").value)%> Division");
			<%
			oRs.MoveNext
		Loop
		%>
		var intDivisions = <%=intDivisions%>;
		<%
	End If
	oRs.NextRecordSet
	
	strSQL = "SELECT t.TeamName, l.LeagueConferenceID, l.lnkLeagueTeamID, l.LeagueDivisionID FROM lnk_league_team l "
	strSQL = strSQL & " INNER JOIN tbl_teams t ON l.TeamID = t.TeamID WHERE LeagueDivisionID > 0 AND LeagueID = '" & intLeagueID & "'"
	strSQL = strSQL & " ORDER BY t.TeamName "
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intTeams = 0
		Do While Not(oRs.EOF)
			intTeams = intTeams + 1
			%>
			arrTeams[<%=intTeams%>] = new Array("<%=oRs.Fields("LeagueConferenceID").Value%>", <%=oRs.Fields("LeagueDivisionID").Value%>, "<%=Server.HTMLEncode(oRS.Fields("lnkLeagueTeamID").value)%>", "<%=Server.HTMLEncode(oRS.Fields("TeamName").value)%>");
			<%
			oRs.MoveNext
		Loop
		%>
		var intTeams = <%=intTeams%>;
		<%
	End If
	oRs.NextRecordSet
	%>
	</script>
	<form name="frmLeagueMatches" id="frmLeagueMatches" action="saveitem.asp" method="post">
	<input type="hidden" name="SaveType" id="SaveType" value="LeagueAddMatch">
	<input type="hidden" name="LeagueID" id="LeagueID" value="<%=intLeagueID%>">
	<input type="hidden" name="League" id="League" value="<%=Server.HTMLEncode(strLeagueName)%>">
	<table border="0" cellspacing="0" cellpadding="0" width="97%" align="center">
	<tr><td width="50%">
	<table border="0" cellspacing="0" cellpadding="0" width="97%" align="center" bgcolor="#444444">
	<tr><td>
	<table border="0" cellspacing="1" cellpadding="4" width="100%" align="center">
	<tr>
		<th colspan="2" bgcolor="#000000">Home Team Information</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right" width="100"><b>Conference:</b></td>
		<td bgcolor="<%=bgcone%>">&nbsp;
			<select name="selHomeConferenceID" id="selHomeConferenceID" onchange="ConferenceChange('Home');">
				<option value="">All Conferences</option>
				<%
				strSQL = "SELECT ConferenceName, c.LeagueConferenceID FROM tbl_league_conferences c WHERE c.LeagueID = '" & intLeagueID & "'"
				oRs.Open strSQL, oConn
				If Not(oRs.EOF AND oRs.BOF) Then
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & oRs.Fields("LeagueConferenceID").Value & """>" & Server.HTMLEncode(oRs.Fields("ConferenceName").Value) & "</option>" & vbCrlF
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
			<select name="selHomeDivisionID" id="selHomeDivisionID" onchange="DivisionChange('Home');">
				<option value="">All Divisions</option>
				<%
				strSQL = "SELECT LeagueDivisionID, DivisionName FROM tbl_league_divisions d WHERE d.LeagueID = '" & intLeagueID & "'"
				oRs.Open strSQL, oConn
				If Not(oRs.EOF AND oRs.BOF) Then
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & oRs.Fields("LeagueDivisionID").Value & """>" & Server.HTMLEncode(oRs.Fields("DivisionName").Value) & "</option>" & vbCrlF
						oRs.MoveNext
					Loop
				End If
				oRs.NextRecordSet
				%>
				</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right"><b>Team:</b></td>
		<td bgcolor="<%=bgcone%>">&nbsp;
			<select name="selHomeLeagueTeamID" id="selHomeLeagueTeamID">
				<option value="">Select a Team</option>
				<%
				strSQL = "SELECT t.TeamName, l.LeagueConferenceID, l.lnkLeagueTeamID, l.LeagueDivisionID FROM lnk_league_team l "
				strSQL = strSQL & " INNER JOIN tbl_teams t ON l.TeamID = t.TeamID WHERE LeagueDivisionID > 0 AND LeagueID = '" & intLeagueID & "'"
				strSQL = strSQL & " ORDER BY t.TeamName "
				oRs.Open strSQL, oConn
				If Not(oRs.EOF AND oRs.BOF) Then
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & oRs.Fields("lnkLeagueTeamID").Value & """>" & Server.HTMLEncode(oRs.Fields("TeamName").Value) & "</option>" & vbCrlF
						oRs.MoveNext
					Loop
				End If
				oRs.NextRecordSet
				%>
			</select>
		</td>
	</tr>
	</table></td></tr>
	</table>
	</td>
	<td width="50%">
	<table border="0" cellspacing="0" cellpadding="0" width="97%" align="center" bgcolor="#444444">
	<tr><td>
	<table border="0" cellspacing="1" cellpadding="4" width="100%" align="center">
	<tr>
		<th colspan="2" bgcolor="#000000">Visitor Team Information</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right" width="100"><b>Conference:</b></td>
		<td bgcolor="<%=bgcone%>">&nbsp;
			<select name="selVisitorConferenceID" id="selVisitorConferenceID" onchange="ConferenceChange('Visitor');">
				<option value="">All Conferences</option>
				<%
				strSQL = "SELECT ConferenceName, c.LeagueConferenceID FROM tbl_league_conferences c WHERE c.LeagueID = '" & intLeagueID & "'"
				oRs.Open strSQL, oConn
				If Not(oRs.EOF AND oRs.BOF) Then
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & oRs.Fields("LeagueConferenceID").Value & """>" & Server.HTMLEncode(oRs.Fields("ConferenceName").Value) & "</option>" & vbCrlF
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
			<select name="selVisitorDivisionID" id="selVisitorDivisionID" onchange="DivisionChange('Visitor');">
				<option value="">All Divisions</option>
				<%
				strSQL = "SELECT LeagueDivisionID, DivisionName FROM tbl_league_divisions d WHERE d.LeagueID = '" & intLeagueID & "'"
				oRs.Open strSQL, oConn
				If Not(oRs.EOF AND oRs.BOF) Then
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & oRs.Fields("LeagueDivisionID").Value & """>" & Server.HTMLEncode(oRs.Fields("DivisionName").Value) & "</option>" & vbCrlF
						oRs.MoveNext
					Loop
				End If
				oRs.NextRecordSet
				%>
				</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right"><b>Team:</b></td>
		<td bgcolor="<%=bgcone%>">&nbsp;
			<select name="selVisitorLeagueTeamID" id="selVisitorLeagueTeamID">
				<option value="">Select a Team</option>
				<%
				strSQL = "SELECT t.TeamName, l.LeagueConferenceID, l.lnkLeagueTeamID, l.LeagueDivisionID FROM lnk_league_team l "
				strSQL = strSQL & " INNER JOIN tbl_teams t ON l.TeamID = t.TeamID WHERE LeagueDivisionID > 0 AND LeagueID = '" & intLeagueID & "'"
				strSQL = strSQL & " ORDER BY t.TeamName "
				oRs.Open strSQL, oConn
				If Not(oRs.EOF AND oRs.BOF) Then
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & oRs.Fields("lnkLeagueTeamID").Value & """>" & Server.HTMLEncode(oRs.Fields("TeamName").Value) & "</option>" & vbCrlF
						oRs.MoveNext
					Loop
				End If
				oRs.NextRecordSet
				%>
			</select>
		</td>
	</tr>
	</table></td></tr>
	</table>
	</td>
	</tr>
	</table>
	<br />
	<table border="0" cellspacing="0" cellpadding="0" width="75%" bgcolor="#444444">
	<tr><td>
	<table border="0" cellspacing="1" cellpadding="4" width="100%">
	<tr>
		<th bgcolor="#000000" colspan="2">Match Information</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right" width="150">Match Date (01/18/2001):</td>
		<td bgcolor="<%=bgctwo%>">&nbsp;
			<input type="text" name="txtMatchDate" id="txtMatchDate" value="<%=FormatDateTime(dtmMatchDate, 2)%>" maxlength="10" />&nbsp;&nbsp;&nbsp;<a href="javascript:ChangeDate()">Display Other Matches</a></td>
	</tr>
	<% 
	strSQL = "SELECT m.MapName FROM tbl_maps m, lnk_league_maps l WHERE l.LeagueID='" & intLeagueID & "' AND l.MapID = m.MapID ORDER BY m.MapName "
	oRs.Open strSQL, oConn, 3, 3
	%>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right" width="150">Map 1:</td>
		<td bgcolor="<%=bgctwo%>">&nbsp;
			<select name="selMap1" id="selMap1">
				<option value="">-- no map --</option>
				<%
				If Not(oRs.EOF AND oRs.BOF) Then
					oRs.MoveFirst
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & Server.HTMLEncode(oRS.Fields("MapName").Value) & """>" & Server.HTMLEncode(oRS.Fields("MapName").Value) & "</option>" & vbCrLF
						oRs.MoveNext
					Loop
				End If
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right" width="150">Map 2:</td>
		<td bgcolor="<%=bgctwo%>">&nbsp;
			<select name="selMap2" id="selMap2">
				<option value="">-- no map --</option>
				<%
				If Not(oRs.EOF AND oRs.BOF) Then
					oRs.MoveFirst
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & Server.HTMLEncode(oRS.Fields("MapName").Value) & """>" & Server.HTMLEncode(oRS.Fields("MapName").Value) & "</option>" & vbCrLF
						oRs.MoveNext
					Loop
				End If
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right" width="150">Map 3:</td>
		<td bgcolor="<%=bgctwo%>">&nbsp;
			<select name="selMap3" id="selMap3">
				<option value="">-- no map --</option>
				<%
				If Not(oRs.EOF AND oRs.BOF) Then
					oRs.MoveFirst
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & Server.HTMLEncode(oRS.Fields("MapName").Value) & """>" & Server.HTMLEncode(oRS.Fields("MapName").Value) & "</option>" & vbCrLF
						oRs.MoveNext
					Loop
				End If
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right" width="150">Map 4:</td>
		<td bgcolor="<%=bgctwo%>">&nbsp;
			<select name="selMap4" id="selMap4">
				<option value="">-- no map --</option>
				<%
				If Not(oRs.EOF AND oRs.BOF) Then
					oRs.MoveFirst
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & Server.HTMLEncode(oRS.Fields("MapName").Value) & """>" & Server.HTMLEncode(oRS.Fields("MapName").Value) & "</option>" & vbCrLF
						oRs.MoveNext
					Loop
				End If
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right" width="150">Map 5:</td>
		<td bgcolor="<%=bgctwo%>">&nbsp;
			<select name="selMap5" id="selMap5">
				<option value="">-- no map --</option>
				<%
				If Not(oRs.EOF AND oRs.BOF) Then
					oRs.MoveFirst
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & Server.HTMLEncode(oRS.Fields("MapName").Value) & """>" & Server.HTMLEncode(oRS.Fields("MapName").Value) & "</option>" & vbCrLF
						oRs.MoveNext
					Loop
				End If
				%>
			</select>
		</td>
	</tr>
	<%
	' Close Map Recordset
	oRs.NextRecordSet
	%>
	<tr>
		<td colspan="2" align="center" bgcolor="#000000"><input type="submit" value="Save Match Information">
	</tr>
	</table>
	</td></tr></table>
	</form>
	<br />
	<br />
	<table border="0" cellspacing="0" cellpadding="0" width="75%" align="center">
	<tr><td>
	<table border="0" cellspacing="1" cellpadding="4" width="100%" align="center" bgcolor="#444444">
	<tr>
		<th colspan="5" bgcolor="#000000">Other matches on <%=FormatDateTime(dtmMatchDate, 2)%></th>
	</tr>
	<tr>
		<th bgcolor="#000000">Division</th>
		<th bgcolor="#000000">Home Team</th>
		<th bgcolor="#000000">Visitor Team</th>
		<th bgcolor="#000000">Delete</th>
		<th bgcolor="#000000">Edit Match</th>
	</tr>
	<%
	strSQL = "SELECT LeagueMatchID, HomeTeamLinkID, VisitorTeamLinkID, LeagueConferenceID, LeagueDivisionID "
	strSQL = strSQL & " FROM tbl_league_matches m"
	strSQL = strSQL & " WHERE DateDiff(d, MatchDate, '" & dtmMatchDate & "') = 0 AND LeagueID = '" & intLeagueID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND orS.BOF) Then
		Do While not(oRS.EOF)
			strHomeTeamName = ""
			strVisitorTeamName = ""
			strSQL = "SELECT DivisionName "
			strSQL = strsQL & " FROM tbl_league_divisions t "
			strSQL = strsQL & " WHERE LeagueDivisionID ='" & oRs.Fields("LeagueDivisionID").Value & "'"
			oRs2.Open strSQL, oConn
			If Not(oRs2.EOF AND oRs2.BOF) Then
				strDivisionName = oRs2.Fields("DivisionName").Value	
			Else
				strDivisionName = "Interdivision "
			End If
			oRS2.NextRecordSet

			strSQL = "SELECT ConferenceName "
			strSQL = strsQL & " FROM tbl_league_conferences "
			strSQL = strsQL & " WHERE LeagueConferenceID ='" & oRs.Fields("LeagueConferenceID").Value & "'"
			oRs2.Open strSQL, oConn
			If Not(oRs2.EOF AND oRs2.BOF) Then
				strConferenceName = oRs2.Fields("ConferenceName").Value	
			Else
				strConferenceName = "Interconference "
				strDivisionName = ""
			End If
			oRS2.NextRecordSet

			strSQL = "SELECT TeamName "
			strSQL = strsQL & " FROM tbl_teams t "
			strSQL = strsQL & " INNER JOIN lnk_league_team lnk"
			strSQL = strsQL & " ON lnk.TeamID = t.TeamID "
			strSQL = strsQL & " WHERE lnk.lnkLeagueTeamID ='" & oRs.Fields("HOmeTeamLinkID").Value & "'"
			oRs2.Open strSQL, oConn
			If Not(oRs2.EOF AND oRs2.BOF) Then
				strHomeTeamName = oRs2.Fields("TeamName").Value	
			End If
			oRS2.NextRecordSet

			strSQL = "SELECT TeamName "
			strSQL = strsQL & " FROM tbl_teams t "
			strSQL = strsQL & " INNER JOIN lnk_league_team lnk"
			strSQL = strsQL & " ON lnk.TeamID = t.TeamID "
			strSQL = strsQL & " WHERE lnk.lnkLeagueTeamID ='" & oRs.Fields("VisitorTeamLinkID").Value & "'"
			oRs2.Open strSQL, oConn
			If Not(oRs2.EOF AND oRs2.BOF) Then
				strVisitorTeamName = oRs2.Fields("TeamName").Value	
			End If
			oRS2.NextRecordSet
			%>
			<tr>
				<td bgcolor="<%=bgcone%>"><%=strConferenceName & " &raquo; " & strDivisionName%></td>
				<td bgcolor="<%=bgctwo%>"><%=strHomeTeamName%></td>
				<td bgcolor="<%=bgcone%>"><%=strVisitorTeamName%></td>
				<td bgcolor="<%=bgctwo%>" align="center"><a href="javascript:DeleteLeagueMatch(<%=oRS.Fields("LeagueMatchID").Value%>)">Delete</a></td>
				<td bgcolor="<%=bgctwo%>" align="center"><a href="LeagueEditMatch.asp?LeagueMatchID=<%=oRS.Fields("LeagueMatchID").Value%>">Edit</a></td>
				
			</tr>
			<%
			oRS.MoveNext
		Loop
	End If
	oRS.Close
	%>
	</tr>
	</table>
	</td></tr>
	</table>
	<%
	Call ContentEnd()
End If
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>