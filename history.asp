<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Team Match History"

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

Dim strTeamName
strTeamName = Request.QueryString("keydata")
If Len(strTeamName) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/default.asp"
End If	

Dim strResult, strEnemyName, bDefender, i, intMaps
Dim map1, map1usScore, Map1ThemScore, map1OT, map1FT
Dim map2, map2usScore, Map2ThemScore, map2OT, map2FT
Dim map3, map3usScore, Map3ThemScore, map3OT, map3FT

Dim intTeamID
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("") %>
Below is the complete history for <%=Server.HTMLEncode(strTeamName)%>, any questions <br />
regarding a particular match should be directed toward the appropriate admin, available <a href="/staff.asp">here</a>.
	<%
	bgc=bgctwo
	strSQL = "EXECUTE GetTeamHistory @TeamName = '" & CheckString(strTeamName) & "'"
	oRS.Open strSQL, oConn
	if not (ors.eof and ors.BOF) then
		%>
		
			<table class="cssBordered" width="100%">
			<tr bgcolor="#000000">
				<TH>Ladder</TH>
				<TH>Opponent</TH>
				<TH>Result</TH>
				<TH>Date</TH>
				<TH>Defender</TH>
				<TH>Demos</TH>
				<% if bSysAdmin Then %>
				<TH>Match Comms</TH>
				<% End If %>
			</tr>
		<%
		do while not ors.EOF
			bDefender = False
			If ors.Fields("WinnerName") = strTeamName Then
				strEnemyName = ors.Fields("LoserName").Value 
				strResult = "Win"
				If oRS.Fields("WinnerDefending").Value Then
					bDefender = True
				End If
			Else
				strEnemyName = ors.Fields("WinnerName").Value 
				strResult = "Loss"
				If oRS.Fields("WinnerDefending").Value Then
					bDefender = True
				End If
			End If
			intMaps = oRs.Fields("Maps").Value 
			
			If bDefender Then
				map1=ors.fields("MatchMap1").value
				map1usscore=ors.fields("MatchMap1DefenderScore").value
				map1themscore=ors.fields("MatchMap1AttackerScore").value
				map1ot=ors.fields("map1ot").value
				map1ft=ors.fields("map1forfeit").value
				map2=ors.fields("MatchMap2").value
				map2usscore=ors.fields("MatchMap2DefenderScore").value
				map2themscore=ors.fields("MatchMap2AttackerScore").value
				map2ot=ors.fields("map2ot").value
				map2ft=ors.fields("map2forfeit").value
				map3=ors.fields("MatchMap3").value
				map3usscore=ors.fields("MatchMap3DefenderScore").value
				map3themscore=ors.fields("MatchMap3AttackerScore").value
				map3ot=ors.fields("map3ot").value
				map3ft=ors.fields("map3forfeit").value
			else
				map1=ors.fields("MatchMap1").value
				map1themscore=ors.fields("MatchMap1DefenderScore").value
				map1usscore=ors.fields("MatchMap1AttackerScore").value
				map1ot=ors.fields("map1ot").value
				map1ft=ors.fields("map1forfeit").value
				map2=ors.fields("MatchMap2").value
				map2themscore=ors.fields("MatchMap2DefenderScore").value
				map2usscore=ors.fields("MatchMap2AttackerScore").value
				map2ot=ors.fields("map2ot").value
				map2ft=ors.fields("map2forfeit").value
				map3=ors.fields("MatchMap3").value
				map3themscore=ors.fields("MatchMap3DefenderScore").value
				map3usscore=ors.fields("MatchMap3AttackerScore").value
				map3ot=ors.fields("map3ot").value
				map3ft=ors.fields("map3forfeit").value
			end if
			%>
			<tr bgcolor=<%=bgctwo%>><td height=22>&nbsp;<%=Server.HTMLEncode(oRS.Fields("LadderName").Value)%></td>
			<td><a href="viewteam.asp?team=<%=server.urlencode(strEnemyName & "")%>"><%=Server.HTMLEncode(strEnemyName & "")%></a></td>
			<td ><%=strResult%></td><td><%=ors.Fields("MatchDate").Value%></td><td align=center>
			<% if bDefender then
				response.write Server.HTMLEncode(strteamName)
			   else
			   	response.write Server.HTMLEncode(strEnemyName & "")
			   end if
			   %>
			   </td>
			   <TD ALIGN=CENTER><A HREF="/demos/default.asp?historyid=<%=ors("HistoryID")%>"><%=ors("demos")%> demos</A></TD>
			   <% If bSysAdmin Then %>
			   <td align="center"><a href="viewmatchcomms.asp?matchid=<%=oRs.Fields("MatchID").Value%>">Match Comms</a></td>
			   <% End If %>
			</tr>
			<tr BGCOLOR="#000000">
			<td><img src="/images/spacer.gif" height="1"></td>
			<td align=left height=20 colspan=3 bgcolor=<%=bgcone%>>
			<% 
			If cBool(oRS.Fields("MatchForfeit").Value) Then
				Response.Write "&nbsp;Admin Forfeited match"
			Else
				If cint(intMaps) > 0 Then
					For i = 1 to cInt(intMaps )
							If i < 6 Then
							Response.Write "&nbsp;<b>" & Server.HTMLEncode(oRS.Fields("MatchMap" & i).Value & "") & ":</b> "
							if NOT (oRS.Fields("MatchMap" & i & "DefenderScore").Value > 0 OR oRS.Fields("MatchMap" & i & "AttackerScore").Value > 0 OR oRS.Fields("Map" & i & "OT").Value OR oRS.Fields("Map" & i & "Forfeit").Value) then
								Response.Write " not played"
							Else
								If bDefender Then
									Response.Write oRS.Fields("MatchMap" & i & "DefenderScore").Value & " - " & oRS.Fields("MatchMap" & i & "AttackerScore").Value 
								Else
									Response.Write oRS.Fields("MatchMap" & i & "AttackerScore").Value & " - " & oRS.Fields("MatchMap" & i & "DefenderScore").Value 
								End IF
								If oRS.Fields("Map" & i & "OT").Value Then
									Response.Write " in OT"
								End If
								If oRS.Fields("Map" & i & "ForFeit").Value THen
									Response.Write " by forfeit"
								End If
							End If
							Response.Write "<BR>"
							End If
					Next
				End If
			End If
			%>
			</td>
			<TD COLSPAN=3>&nbsp;</TD>
			</tr>
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.MoveNext
		loop
		%>
		</table>
		<%
	end if
	ors.Close 
	%>
	<br /><br />
<%
strSQL = "EXECUTE GetLeagueHistoryForTeam @TeamName = '" & CheckString(strTeamName) & "'"
oRs.Open strSQL, oConn
If oRs.State = 1 Then
	If Not(oRs.EOF AND oRs.BOF) Then
		%>
			<table class="cssBordered" width="100%">
		<tr bgcolor="#000000">
			<TH width="35%">League</TH>
			<TH>Home Team</TH>
			<TH>Visitor Team</TH>
			<TH>Date</TH>
			<% If bSysAdmin or IsAnyLeagueAdmin() Then %>
			<th>Comms</th>
			<th>Edit History</th>
			<% End If %>
		</tr>
		<%
		Do While Not (oRs.EOF)
			if bgc = bgcone then
				bgc = bgctwo
			else
				bgc = bgcone
			end if
			%>
			<tr>
				<td bgcolor="<%=bgc%>" rowspan="2"><%
				Response.Write "<a href=""viewleague.asp?league=" & Server.URLEncode(oRs.Fields("LeagueName").Value & "") & """>" & Server.HTMLEncode(oRs.Fields("LeagueName").Value & "") & " League</a> &raquo; <br />"
				If IsNull(oRs.Fields("ConferenceName").Value) Then
					Response.Write "&nbsp;&nbsp;&nbsp;Interconference"	
				Else 
					Response.Write "&nbsp;&nbsp;&nbsp;<a href=""viewleagueconference.asp?league=" & Server.URLEncode(oRs.Fields("LeagueName").Value & "") & "&conference=" & Server.URLEncode(oRs.Fields("ConferenceName").Value & "") & """>" & Server.HTMLEncode(oRs.Fields("ConferenceName").Value & "") & " Conference</a> &raquo; <br />"
					If IsNull(oRs.Fields("ConferenceName").Value) Then
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Interdivision"	
					Else 
						Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""viewleaguedivision.asp?league=" & Server.URLEncode(oRs.Fields("LeagueName").Value & "") & "&conference=" & Server.URLEncode(oRs.Fields("ConferenceName").Value & "") & "&division=" & Server.URLEncode(oRs.Fields("DivisionName").Value & "") & """>" & Server.HTMLEncode(oRs.Fields("DivisionName").Value & "") & " Division</a>"
					End If
				End If
				%></td>
				<td bgcolor="<%=bgc%>"><a href="viewteam.asp?team=<%=Server.URLEncode(oRs.Fields("HomeTeamName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("HomeTeamName").Value & "")%></a></td>
				<td bgcolor="<%=bgc%>"><a href="viewteam.asp?team=<%=Server.URLEncode(oRs.Fields("VisitorTeamName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("VisitorTeamName").Value & "")%></a></td>
				<td bgcolor="<%=bgc%>" align="center"><%=FormatDateTime(oRs.Fields("MatchDate").Value, 2)%></td>
			   <% If bSysAdmin Or IsAnyLeagueAdmin() Then %>
			   <td align="center" bgcolor="<%=bgc%>"><a href="viewleaguematchcomms.asp?matchid=<%=oRs.Fields("LeagueMatchID").Value%>">Match Comms</a></td>
			   <td align="center" bgcolor="<%=bgc%>"><a href="LeagueEditHistory.asp?historyID=<%=oRs.Fields("LeagueHistoryID").Value%>&f=<%=Server.URLENcode("history.asp?keydata=" & strTeamName)%>">Edit History</a></td>
			   <% End If %>
			</tr>
			<tr>
				
			   	<% If bSysAdmin or IsAnyLeagueAdmin() Then %>
				<td colspan="5"  bgcolor="<%=bgc%>" align="center">
				<% else %>
				<td colspan="3"  bgcolor="<%=bgc%>" align="center">
				<% End If %>
			<table width="100%">
				<%
				For i = 1 to 5
					If Len(oRs.Fields("Map" & i).value) > 0 Then
						%>
						<tr>
							<td bgcolor="<%=bgctwo%>" width="45%"><b><%=oRs.Fields("Map" & i).value%></b></td>
							<%
							If oRs.Fields("Map" & i & "VisitorScore").value = oRs.Fields("Map" & i & "HomeScore").value Then 
								%>
								<td align="center" bgcolor="<%=bgctwo%>"><%=oRs.Fields("Map" & i & "HomeScore").value%></td>
								<td align="center" bgcolor="<%=bgctwo%>"><%=oRs.Fields("Map" & i & "VisitorScore").value%></td>
								<%
							ElseIf oRs.Fields("Map" & i & "VisitorScore").value > oRs.Fields("Map" & i & "HomeScore").value Then 
								%>
								<td align="center" bgcolor="<%=bgctwo%>"><%=oRs.Fields("Map" & i & "HomeScore").value%></td>
								<td align="center" bgcolor="<%=bgctwo%>"><font color="#00cc00"><b><%=oRs.Fields("Map" & i & "VisitorScore").value%></b></td>
								<%
							Else
								%>
								<td align="center" bgcolor="<%=bgctwo%>"><font color="#00cc00"><b><%=oRs.Fields("Map" & i & "HomeScore").value%></b></td>
								<td align="center" bgcolor="<%=bgctwo%>"><%=oRs.Fields("Map" & i & "VisitorScore").value%></td>
								<%
							End If
							%>
						</tr>
						<%
'						Response.Write "<b>" & oRs.Fields("Map" & i).value & "</b> (" & oRs.Fields("Map" & i & "HomeScore").value & " - " & oRs.Fields("Map" & i & "VisitorScore").value & ") <br />"
					End If
				Next
				%>
				</table>
				</td>
			</tr>
			<%
			oRs.MoveNext
		Loop
		%>
		</table>
	<%
End If
oRs.NextRecordSet
End If
%>

<%
'' Scrim Ladder History

strSQL = "SELECT TeamID FROM tbl_teams WHERE TEamName = '" & CheckString(strTeamName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intTeamID = oRs.FIeldS("TeamID").Value
End If
oRs.NextRecordSet
dIM map4, map5, map4themscore, map5themscore, map4usscore, map5usscore, map4ot, map5ot
strSQL = "SELECT 'DefenderName' = dt.TeamName, 'AttackerName' = at.TeamName, "
strSQL = strSQL & " Map1, Map2, Map3, Map4, Map5, "
strSQL = strSQL & " Map1OT, Map2OT, Map3OT, Map4OT, Map5OT, "
strSQL = strSQL & " Map1DefenderScore, Map2DefenderScore, Map3DefenderScore, Map4DefenderScore, Map5DefenderScore, "
strSQL = strSQL & " Map1AttackerScore, Map2AttackerScore, Map3AttackerScore, Map4AttackerScore, Map5AttackerScore, "
strSQL = strSQL & " l.EloLadderName, DefenderRatingDiff, AttackerRatingDiff, "
strSQL = strSQL & " MatchDate, MatchWinnerDefending, EloHistoryID "
strSQL = strSQL & " FROM tbl_elo_history eh "
strSQL = strSQL & " INNER JOIN tbl_elo_ladders l ON eh.EloLadderID = l.EloLadderID "
strSQL = strSQL & " INNER JOIN lnk_elo_team det ON det.lnkEloTeamID = DefenderEloTeamID "
strSQL = strSQL & " INNER JOIN lnk_elo_team aet ON aet.lnkEloTeamID = AttackerEloTeamID "
strSQL = strSQL & " INNER JOIN tbl_teams at ON at.TeamID = aet.TeamID "
strSQL = strSQL & " INNER JOIN tbl_teams dt ON dt.TeamID = det.TeamID "
strSQL = strSQL & " WHERE dt.TeamID = '" & intTeamID & "' OR at.TeamID = '" & intTeamID & "' ORDER BY l.EloLadderName ASC, MatchDate DESC "
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	%>
	<br /><br />
			<table class="cssBordered" width="100%">
	<tr bgcolor="#000000">
		<TH>Ladder</TH>
		<TH>Opponent</TH>
		<TH>Result</TH>
		<TH>Date</TH>
		<TH>Defender</TH>
		<% If Session("uName") = "Triston" Then %>
		<th>History ID</th>
		<% End If %>
	</tr>
	<%
	Do While Not(oRs.EOF)
			If ors.Fields("DefenderName") = strTeamName Then
				strEnemyName = ors.Fields("AttackerName").Value 
				If oRS.Fields("MatchWinnerDefending").Value Then
					strResult = "Win"
				Else
					strResult = "Loss"
				End If
				bDefender = True
			Else
				strEnemyName = ors.Fields("DefenderName").Value 
				If oRS.Fields("MatchWinnerDefending").Value Then
					strResult = "Loss"
				else
					strResult = "Win"
				End If
				bDefender = false
			End If
			intMaps = 5
			
			If bDefender Then
				map1=ors.fields("Map1").value
				map1usscore=ors.fields("Map1DefenderScore").value
				map1themscore=ors.fields("Map1AttackerScore").value
				map1ot=ors.fields("map1ot").value
				map2=ors.fields("Map2").value
				map2usscore=ors.fields("Map2DefenderScore").value
				map2themscore=ors.fields("Map2AttackerScore").value
				map2ot=ors.fields("map2ot").value
				map3=ors.fields("Map3").value
				map3usscore=ors.fields("Map3DefenderScore").value
				map3themscore=ors.fields("Map3AttackerScore").value
				map3ot=ors.fields("map3ot").value

				map4=ors.fields("Map4").value
				map4usscore=ors.fields("Map4DefenderScore").value
				map4themscore=ors.fields("Map4AttackerScore").value
				map4ot=ors.fields("map4ot").value

				map5=ors.fields("Map5").value
				map5usscore=ors.fields("Map5DefenderScore").value
				map5themscore=ors.fields("Map5AttackerScore").value
				map5ot=ors.fields("map5ot").value
			else
				map1=ors.fields("Map1").value
				map1themscore=ors.fields("Map1DefenderScore").value
				map1usscore=ors.fields("Map1AttackerScore").value
				map1ot=ors.fields("map1ot").value
				map2=ors.fields("Map2").value
				map2themscore=ors.fields("Map2DefenderScore").value
				map2usscore=ors.fields("Map2AttackerScore").value
				map2ot=ors.fields("map2ot").value
				map3=ors.fields("Map3").value
				map3themscore=ors.fields("Map3DefenderScore").value
				map3usscore=ors.fields("Map3AttackerScore").value
				map3ot=ors.fields("map3ot").value

				map4=ors.fields("Map4").value
				map4themscore=ors.fields("Map4DefenderScore").value
				map4usscore=ors.fields("Map4AttackerScore").value
				map4ot=ors.fields("map4ot").value

				map5=ors.fields("Map5").value
				map5themscore=ors.fields("Map5DefenderScore").value
				map5usscore=ors.fields("Map5AttackerScore").value
				map5ot=ors.fields("map5ot").value

			end if
			%>
			<tr bgcolor=<%=bgctwo%>><td height=22>&nbsp;<%=Server.HTMLEncode(oRS.Fields("EloLadderName").Value)%></td>
			<td><a href="viewteam.asp?team=<%=server.urlencode(strEnemyName & "")%>"><%=Server.HTMLEncode(strEnemyName & "")%></a></td>
			<td ><%=strResult%></td>
			<td><%=ors.Fields("MatchDate").Value%></td><td align=center>
			<% if bDefender then
				response.write Server.HTMLEncode(strteamName)
			   else
			   	response.write Server.HTMLEncode(strEnemyName & "")
			   end if
			   %>
			   </td>
			  <% If Session("uName") = "Triston" Then %>
			  <td><%=oRs.Fields("EloHistoryID").Value%></td>
			  <% End If %>
			</tr>
			<tr BGCOLOR="#000000">
			<td><img src="/images/spacer.gif" height="1"></td>
			<td align=left height=20 colspan=3 bgcolor=<%=bgcone%>>
			<% 
				If cint(intMaps) > 0 Then
					For i = 1 to cInt(intMaps )
						If Len(oRs.Fields("Map" & i).Value) > 0 Then 
							Response.Write "&nbsp;<b>" & Server.HTMLEncode(oRS.Fields("Map" & i).Value & "") & ":</b> "
							if NOT (oRS.Fields("Map" & i & "DefenderScore").Value > 0 OR oRS.Fields("Map" & i & "AttackerScore").Value > 0 OR oRS.Fields("Map" & i & "OT").Value) then
								Response.Write " not played"
							Else
								If bDefender Then
									Response.Write oRS.Fields("Map" & i & "DefenderScore").Value & " - " & oRS.Fields("Map" & i & "AttackerScore").Value 
								Else
									Response.Write oRS.Fields("Map" & i & "AttackerScore").Value & " - " & oRS.Fields("Map" & i & "DefenderScore").Value 
								End IF
								If oRS.Fields("Map" & i & "OT").Value Then
									Response.Write " in OT"
								End If
							End If
							Response.Write "<BR>"
						End If
					Next
				End If
			%>
			</td>
			<TD COLSPAN=3>&nbsp;</TD>
			</tr>
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
		oRs.MoveNext
	Loop
	%>
	</table>
	<%	
End If
oRs.NextRecordSet

Call ContentEnd() 
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>
