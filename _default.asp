<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TeamWarfare: Community Based Gaming"

Dim strSQL, oConn, oRS, oRS2, oRS3
Dim bgcone, bgctwo, bgcheader, bgcblack, bgc

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")
Set oRS3 = Server.CreateObject("ADODB.RecordSet")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, blnLoggedIn
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
blnLoggedIn = Session("LOggedIn")

Dim intPendingCnt, intHistoryCnt, strCurrDate
Dim strWinnerTag, strWinnerName
Dim strLoserTag, strLoserName
Dim strLadderAbbr
Dim strDefenderName, strDefenderTag
Dim strAttackerName, strAttackerTag
Dim strMatchTime

Dim intNewsID, intArticles, intNewsCnt
intArticles = 6

Dim intPlayerLadderID
Dim intForfeits, strEnemyName, strResult, strPlayerName
Dim map, opponent, mDate, statusVerbage

Dim strMapArray(6), intMatchID
Dim strMatchDate
Dim i, strVerbiage
Dim strEnemyVerbiage, CurrentMap, CurrentMapNumber
Dim PPLLinkID

Dim intPlayerID
Dim mStatus, players
Dim intEnemyLinkID, xDate, aDate
Dim matchdate1, matchdate2, grammer
Dim intOptionID, blnOptionSame, intCounter, blnOptionShown
		
Dim strMapConfiguration
		
Dim strTeamName, strTeamTag, intTeamID
Dim intTLLinkID, intLadderID, strLadderName
Dim intMaps, blnIsAdmin, intGameID
Dim strGameName, strLadderRules, intForumID
Dim intRank, intWins 
Dim intLosses, strStatus
Dim intLadderAdminID, strLadderAdminName, strLadderAdminEmail
Dim intMinPlayer, intLadderLocked

Dim intDefenderVotes, intAttackerVotes, strLastRanterName, strLastRantTime, intRants

strPlayerName = Session("uName")
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<% 
'Response.Flush 
If blnLoggedIn Then
	strSQL = "EXECUTE PersonalizedHomePage '" & Session("PlayerID") & "'"
Else
	strSQL = "EXECUTE PersonalizedHomePage '0'"
End If
If Session("uName") = "Triston" Then
	'strSQL = "EXECUTE PersonalHomeReturnTeams_LeagueStyle '" & Session("PlayerID") & "'"
'	Response.Write strSQL
'	Response.End
End If
oRS.Open strSQL, oConn, 3, 3
%>

	<% Call Content2BoxStart("") %>
		<table border=0 cellspacing="0" cellpadding="0" width="97%" align="center" class="cssBordered">
			<tr>
				<th colspan="2" bgcolor="#000000">Recent News</th>
			</tr>		

	<%
	intNewsID = 0
	intNewsCnt = 1
	If Not(oRS.EOF and oRS.BOF) Then
		bgc = bgcone
		Do While Not(oRS.EOF)
			%>			
					<tr valign="top"> 
						<td bgcolor="<%=bgc%>"><%=oRs.Fields("GameAbbreviation").Value%>: <a href=#<%=ors.fields("NewsID").value%>><%=Server.HTMLEncode (ors.fields("NewsHeadLine").value)%></a></td>
						<td bgcolor="<%=bgc%>"> <%=FormatDateTime(oRS.Fields("NewsDate").Value, vbShortDate) %> </td>				   
					</tr>
			
			<%
			intNewsCnt = intNewsCnt + 1
			if bgc = bgcone then
				bgc = bgctwo
			else
				bgc = bgcone
			End If
			oRS.MoveNext 
		Loop
	Else
		Response.Write "<tr valign=""top""><td bgcolor=""" & bgcone & """ colspan=""2""><i>No recent news available for the games you participate in.</i></td></tr>"
	End If
	Set oRS = oRS.NextRecordSet
%>
	<tr>
		<td align="right" colspan="2"><a href="/xml/news.asp">xml news</a> / <a href="allnews.asp">view all news</a>&nbsp;&nbsp;</td>
	</tr>
	</table>
	<% Call Content2BoxMiddle() %>
		<table border=0 cellspacing="0" cellpadding="0" width="97%" class="cssBordered" align="center">
			<tr>
				<th bgcolor="#000000" colspan="2">Announcements &amp; Information</th>
			</tr>		
			<%
			intNewsID = 0
			intNewsCnt = 1
			If Not(oRS.EOF and oRS.BOF) Then
				bgc = bgcone
				Do While Not(oRS.EOF)
					%>			
							<tr valign="top"> 
								<td bgcolor="<%=bgc%>"><a href=#<%=ors.fields("NewsID").value%>><%=Server.HTMLEncode (ors.fields("NewsHeadLine").value)%></a></td>
								<td bgcolor="<%=bgc%>"> <%=FormatDateTime(oRS.Fields("NewsDate").Value, vbShortDate) %> </td>				   
							</tr>
					<%
					if bgc = bgcone then
						bgc = bgctwo
					else
						bgc = bgcone
					End If
					intNewsCnt = intNewsCnt + 1
					oRS.MoveNext 
				Loop
			Else
				Response.Write "<tr valign=""top""><td bgcolor=""" & bgcone & """ colspan=""2""><i>No recent announcements</i></td></tr>"
			End If
			Set oRS = oRS.NextRecordSet
			%>
	</table>
	<% Call Content2BoxEnd() %>
	<% 
If Session("LoggedIn") Then
	'' LOGGED IN
	If oRS.State = 1 Then
		'----------------------------------
		'' Tournaments
		'----------------------------------
		If Not(oRs.BOF And oRs.EOF) Then
			Call Content33BoxStart("Tournament Status") 
			Do While Not(oRs.EOF)
				strTeamName = oRS.Fields("TeamName").Value
				strTeamTag = oRS.Fields("TeamTag").Value
				intTeamID = oRS.Fields("TeamID").Value
				%>
				<table border="0" cellspacing="0" cellpadding="0" align="center" width="97%" class="cssBordered">
				<tr>
					<th colspan="2"><a href="tournament/default.asp?Tournament=<%=Server.URLEncode(oRs.Fields("TournamentName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("TournamentName").Value & "")%></a></th>
				</tr>
				<tr>
					<td bgcolor="<%=bgcone%>" align="right">Team:</td>
					<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(strTeamName & "")%>"><%=Server.HTMLEncode(strTeamName & "")%></a></td>
				</tr>				
				<tr>
					<td>&nbsp;</td>
					<td bgcolor="<%=bgctwo%>"><a href="tournament/default.asp?Tournament=<%=Server.URLEncode(oRs.Fields("TournamentName").Value & "")%>&page=brackets&div=<%=oRs.Fields("DivisionID").Value%>">Tournament Brackets</a></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td bgcolor="<%=bgcone%>"><a href="tournament/default.asp?Tournament=<%=Server.URLEncode(oRs.Fields("TournamentName").Value & "")%>&page=schedule">Tournament Schedule</a></td>
				</tr>
				<% If oRs.Fields("HasPrizes").Value = 1 Then %>
				<tr>
					<td>&nbsp;</td>
					<td bgcolor="<%=bgctwo%>"><a href="tournament/default.asp?Tournament=<%=Server.URLEncode(oRs.Fields("TournamentName").Value & "")%>&page=prizes">Tournament Prizes</a></td>
				</tr>
				<% End If %>
				<% If oRs.Fields("HasSponsors").Value = 1 Then %>
				<tr>
					<td>&nbsp;</td>
					<td bgcolor="<%=bgcone%>"><a href="tournament/default.asp?Tournament=<%=Server.URLEncode(oRs.Fields("TournamentName").Value & "")%>&page=sponsors">Tournament Sponsors</a></td>
				</tr>
				<% End If %>
				<tr>
					<td>&nbsp;</td>
					<td bgcolor="<%=bgctwo%>"><a href="forums/forumdisplay.asp?forumid=<%=Server.URLEncode(oRs.Fields("ForumID").Value & "")%>">Tournament Forum</a></td>
				</tr>
				</table>
				<%
				Call Content33BoxMiddle()
				%>
				<table border="0" cellspacing="0" cellpadding="0" align="center" width="97%" class="cssBordered">
				<tr>
					<th>Current Status</th>
				</tr>
					<%
					'Activity Code here
					Dim ServerName, ServerIP, MatchTime, TMLinkID
					Dim RoundsID, Team1, Team1ID, Team2ID, OpponentLinkID, Team1Name, Team2Name, LocationVerb, OpponentName, RoundNum, CurrentStatus
					TMLinkID = ors.Fields("TMLinkID").Value 
					strsql = "select *, Team1Name = (select TeamName from tbl_teams t, lnk_t_m lnk where TMLinkID = Team1ID AND lnk.teamid = t.teamid), " &_
								" Team2Name = (select TeamName from tbl_teams t, lnk_t_m lnk where TMLinkID = Team2ID AND lnk.teamid = t.teamid) " &_
								" from tbl_rounds where (Team1ID = '" & TMLinkID & "' or Team2ID = '" & TMLinkID & "') AND WinnerID = 0 order by Round desc"
					ors2.open strsql, oconn
					if not(ors2.eof and ors2.bof) then
						if ors2("Team1ID") = TMLinkID then
							RoundsID = ors2("RoundsID")
							Team1 = true
							Team1ID = TMLinkID
							Team2ID = ors2("Team2ID")
							OpponentLinkID = ors2("Team2ID")
							Team1Name = ors2("Team1Name")
							Team2Name = ors2("Team2Name")
							locationverb = "home"
							opponentname = Team2Name
						else
							RoundsID = ors2("RoundsID")
							Team1 = false
							Team1ID = ors2("Team1ID")
							Team2ID = TMLinkID
							OpponentLinkID = ors2("Team1ID")
							Team1Name = ors2("Team1Name")
							Team2Name = ors2("Team2Name")
							locationverb = "visitor"
							opponentname = Team1Name
						end if
						RoundNum = ors2("Round")
						ServerName = ors2.Fields("ServerName").Value
						ServerIP = ors2.Fields("ServerIP").Value
						MatchTime = ors2.Fields("MatchTime").Value
						if team1id = "0" or team2id = "0" then
							CurrentStatus = "Awaiting another team to be seeded into your bracket."
							OpponentName = "TBD"
						else
							CurrentStatus = "Challenging <a href=""/viewteam.asp?team=" & server.URLEncode(OpponentName)
							CurrentStatus = CurrentStatus & """>" & Server.HTMLEncode(OpponentName) & "</a> in round " & roundnum & "."
										
						end if
						%>
						<TR BGCOLOR="#000000"><TD>Current round: <%=RoundNum%>, <%=ucase(locationverb)%></TD></TR>
						<TR BGCOLOR="#000000"><TD>Current opponent: <a href="/viewteam.asp?team=<%=server.URLEncode(OpponentName)%>"><%=OpponentName%></a></TD></TR>
						<% If Not(IsNull(ServerName)) Then %>
						<tr>
							<td bgcolor="#000000"><b>Server Name:</b> <%=ServerName%></td>
						</tr>
						<tr>
							<td bgcolor="#000000"><b>Server IP:</b> <%=ServerIP%></td>
						</tr>
						<tr>
							<td bgcolor="#000000"><b>Match Time:</b> <%=FormatDateTime(MatchTime, 0)%></td>
						</tr>
						<% 
						If oRs.Fields("IsAdmin").Value = 1 Then 
							%>
							<tr>
								<td>Last Comm:
							<%
							strSQL = "SELECT TOP 1 CommDate, CommTime FROM tbl_round_comm WHERE RoundsID = '" & RoundsID & "' ORDER BY CommID DESC"
							oRs3.Open strSQL, oConn
							If Not(oRs3.EOF AND oRs3.BOF) Then
								Response.Write "" & oRs3.Fields("CommDate").Value & " " & formatdatetime(oRs3.Fields("CommTime").Value, 3) & "</td></tr>"
							Else
								Response.Write "No comms yet</td></tr>"
							End If
							oRs3.NextRecordSet
						End If 
					End If
					oRs2.NextRecordSet
					%>
					<% If oRs.Fields("IsAdmin").Value = 1 Then %>
					<tr>
						<td align="center" ><a href="TeamTournamentAdmin.asp?tournament=<%=server.urlencode(ors.Fields("TournamentName").Value)%>&team=<%=server.urlencode(strTeamName)%>">Tournament Admin Panel</a></td>
					</tr>
					<% End If %>
					</table>
					<%
				End If
				oRs.MoveNext
				If Not (oRs.EOF) Then
					Call Content33BoxEnd() 
					Call Content33BoxStart("")
				End If
			Loop
			Call Content33BoxEnd() 
		End If
	End If
	'----------------------------------
	'' END Tournaments
	'----------------------------------

	Set oRs = oRs.NextRecordSet
	If oRS.State = 1 Then
		'----------------------------------
		'' Leagues
		'----------------------------------
		If Not(oRs.BOF And oRs.EOF) Then
			Call Content33BoxStart("League Competition Status") 
			Do While Not(oRs.EOF)
				strTeamName = oRS.Fields("TeamName").Value
				strTeamTag = oRS.Fields("TeamTag").Value
				intTeamID = oRS.Fields("TeamID").Value
				%>
				<table border="0" cellspacing="0" cellpadding="0" align="center" width="97%" class="cssBordered">
				<tr>
					<th colspan="2"><a href="viewleague.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("LeagueName").Value & "")%> League</a></th>
				</tr>
				<tr>
					<td bgcolor="<%=bgcone%>" align="right">Team:</td>
					<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(strTeamName & "")%>"><%=Server.HTMLEncode(strTeamName & "")%></a></td>
				</tr>
				<tr>
					<td bgcolor="<%=bgctwo%>" align="right">Points:</td>
					<td bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(ors.Fields("LeaguePoints").Value & "")%> Points</td>
				</tr>
				<tr>
					<td bgcolor="<%=bgcone%>" align="right">W/L/Pct:</td>
					<td  bgcolor="<%=bgcone%>"><%=oRs.Fields("Wins").Value%>/<%=oRs.Fields("Losses").Value%>/<%=FormatNumber(cInt(oRs.Fields("WinPct").Value) / 10000, 3, 0)%></td>
				</tr>
				<% If Not(IsNull(oRs.Fields("DivisionName").Value)) Then %>
				<tr>
					<td>&nbsp;</td>
					<td bgcolor="<%=bgctwo%>"><a href="viewleagueconference.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value & "")%>&conference=<%=Server.URLEncode(oRs.Fields("ConferenceName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("ConferenceName").Value & "")%> Conference</a></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td bgcolor="<%=bgcone%>"><a href="viewleaguedivision.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value & "")%>&conference=<%=Server.URLEncode(oRs.Fields("ConferenceName").Value & "")%>&division=<%=Server.URLEncode(oRs.Fields("DivisionName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("DivisionName").Value & "")%> Division</a></td>
				</tr>
				<% End If %>
				<tr>
					<td>&nbsp;</td>
					<td bgcolor="<%=bgctwo%>"><a href="forums/forumdisplay.asp?forumid=<%=Server.URLEncode(oRs.Fields("GameForumID").Value & "")%>">Game Forum</a></td>
				</tr>
				</table>
				<%
				Call Content33BoxMiddle()
				%>
				<table border="0" cellspacing="0" cellpadding="0" align="center" width="97%" class="cssBordered">
				<tr>
					<th colspan="7">Pending Matches</th>
				</tr>
					<%
					strSQL = "EXECUTE LeagueTeamMatches @LeagueTeamID = '" & oRs.Fields("lnkLeagueTeamID").Value & "'"
					oRs2.Open strSQL, oConn
					If (oRs2.State = 1) Then
						If Not(oRs2.EOF AND oRs2.BOF) Then
							%>
							<tr>
								<th bgcolor="#000000">Date</th>
								<th bgcolor="#000000">Opponent</th>
								<th bgcolor="#000000">Status</th>
								<th bgcolor="#000000">Maps</th>
								<% If oRs.Fields("IsAdmin").Value Then %>
									<th bgcolor="#000000">Comms</th>
									<th bgcolor="#000000">Last Comm</th>
									<th bgcolor="#000000">Details</th>
								<% End If %>
							</tr>
							<%
							Do While Not (oRs2.EOF)
								%>
								<tr>
									<td valign="top" align="center" bgcolor="<%=bgcone%>"><%=FormatDateTime(oRs2.FIelds("MatchDate").Value, 2)%></td>
									<td valign="top"  bgcolor="<%=bgctwo%>"><a href="viewteam.asp?team=<%=Server.URLEncode(oRs2.Fields("OpponentName").Value)%>"><%=Server.HTMLEncode(oRs2.Fields("OpponentName").Value)%></a></td>
									<% If cInt(oRs2.Fields("HomeTeamLinkID").Value) = oRs.Fields("lnkLeagueTeamID").Value Then %>
									<td valign="top" align="center" bgcolor="<%=bgcone%>">Home</td>
									<% Else %>
									<td valign="top" align="center" bgcolor="<%=bgcone%>">Visitor</td>
									<% End If %>
									<td valign="top"  bgcolor="<%=bgctwo%>" align="center">
										<%
										If Len(oRs2.Fields("Map1").Value) > 0 Then
											Response.Write oRs2.Fields("Map1").Value
										End If
										If Len(oRs2.Fields("Map2").Value) > 0 Then
											Response.Write "<br /> " & oRs2.Fields("Map2").Value
										End If
										If Len(oRs2.Fields("Map3").Value) > 0 Then
											Response.Write "<br /> " & oRs2.Fields("Map3").Value
										End If
										If Len(oRs2.Fields("Map4").Value) > 0 Then
											Response.Write "<br /> " & oRs2.Fields("Map4").Value
										End If
										If Len(oRs2.Fields("Map5").Value) > 0 Then
											Response.Write "<br /> " & oRs2.Fields("Map5").Value
										End If
										%></td>
	
								<% If oRs.Fields("IsAdmin").Value Then %>
									<td valign="top"  align="center" bgcolor="<%=bgctwo%>"><%=oRs2.FIelds("CommsCount").Value%></td>
									<td valign="top"  align="center" bgcolor="<%=bgcone%>"><%
										If Not(IsNull(oRs2.Fields("LastCommDate").Value)) Then
											Response.Write oRs2.Fields("LastCommDate").Value & "<br />" & Server.HTMLEncode(oRs2.Fields("LastCommAuthor").Value)
										Else
											Response.Write "none"
										End If
										%></td>
									<td valign="top" align="center" bgcolor="<%=bgctwo%>"><a href="teamleagueadmin.asp?league=<%=Server.URLEncode(oRs.Fields("LeagueName").Value)%>&team=<%=Server.URLEncode(strTeamName)%>&matchid=<%=oRs2.Fields("LeagueMatchID")%>">Manage</a></td>
								<% End If %>
								</tr>
								<%
								oRs2.MoveNext
							Loop
						Else
							%>
							<tr>
								<td colspan="7" bgcolor="<%=bgctwo%>">No pending matches scheduled.</td>
							</tr>
							<%
						End If
						oRs2.NextRecordSet
					End If
					%>					
				<% If oRs.Fields("IsAdmin").Value = 1 Then %>
					<tr>
						<td align="center" colspan="7"><a href="TeamLeagueAdmin.asp?league=<%=server.urlencode(ors.Fields("LeagueName").Value)%>&team=<%=server.urlencode(strTeamName)%>">League Admin Panel</a></td>
					</tr>
					<% End If %>
					</table>
					<%
				
				oRs.MoveNext
				If Not (oRs.EOF) Then
					Call Content33BoxEnd() 
					Call Content33BoxStart("")
				End If
			Loop
			Call Content33BoxEnd() 
		End If
	End If
	Set oRs = oRs.NextRecordSet
	
	If oRS.State = 1 Then
		If Not(oRS.EOF AND oRS.BOF) Then
			Call Content33BoxStart("Ladder Competition Status") 
			Do While Not (oRS.EOF)
				strTeamName = oRS.Fields("TeamName").Value
				strTeamTag = oRS.Fields("TeamTag").Value
				intTeamID = oRS.Fields("TeamID").Value
				intTLLinkID = oRS.Fields("TLLinkID").Value
				intLadderID = oRS.Fields("LadderID").Value
				strLadderName = oRS.Fields("LadderName").Value
				intMaps = oRS.Fields("Maps").Value
				blnIsAdmin = cBool(oRS.Fields("IsAdmin").Value)
				intGameID = oRS.Fields("GameID").Value
				strGameName = oRS.Fields("GameName").Value
				strLadderRules = oRS.Fields("LadderRules").Value
			'	intForumID = oRS.Fields("ForumID").Value
				
				intRank = oRS.Fields("Rank").Value
				intWins = oRS.Fields("Wins").Value
				intLosses = oRS.Fields("Losses").Value
				strStatus = oRS.Fields("Status").Value
				intLadderAdminID = oRS.Fields("LadderAdminID").Value
				strLadderAdminName = oRS.Fields("LadderAdminName").Value
				strLadderAdminEmail = oRS.Fields("LadderAdminEmail").Value
				intMinPlayer = oRS.Fields("MinPlayer").Value
				intLadderLocked = oRS.Fields("LadderLocked").Value
				strMapConfiguration = oRS.Fields("MapConfiguration").Value
				
				If IsNull(oRS.Fields("LadderForumID").Value) Then
					intForumID = oRS.Fields("GameForumID").Value
				Else
					intForumID = oRS.Fields("LadderForumID").Value
				End If
	%>
		<table border="0" cellspacing="0" cellpadding="0" align="center" width="97%">
		<tr>
			<td bgcolor="#444444">
				<table border="0" cellspacing="1" cellpadding="2" width="100%">
				<tr>
					<th bgcolor="#000000" colspan="2"><a href="viewladder.asp?ladder=<%=Server.URLEncode(strLadderName)%>"><%=Server.HTMLEncode(strLadderName & "")%></a> #<%=intRank & " (" & intWins & "/" & intLosses & ")"%></th>
				</tr>
				<tr>
					<td bgcolor="<%=bgcone%>">Team:</td>
					<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(strTeamName)%>"><%=Server.HTMLEncode(strTeamName & "")%></a></td>
				</tr>
				<tr>
					<td width="15%" bgcolor="#000000">&nbsp;</td>
					<td bgcolor="<%=bgcone%>"><a href="history.asp?keydata=<%=Server.URLEncode(strTeamName)%>">Match History</a></td>
				</tr>
				<tr>
					<td width="15%" bgcolor="#000000">&nbsp;</td>
					<td bgcolor="<%=bgctwo%>"><a href="rules.asp?set=<%=Server.URLEncode(strLadderRules)%>">View Ladder Rules</a></td>
				</tr>
				<tr>
					<td width="15%" bgcolor="#000000">&nbsp;</td>
					<td bgcolor="<%=bgcone%>"><a href="viewladderdetails.asp?ladder=<%=Server.URLEncode(strLadderName)%>">View Ladder Information</a></td>
				</tr>
				<tr>
					<td width="15%" bgcolor="#000000">&nbsp;</td>
					<td bgcolor="<%=bgctwo%>"><a href="forums/forumdisplay.asp?forumid=<%=Server.URLEncode(intForumID)%>">Game Forum</a></td>
				</tr>
				<% If Not(IsNull(intLadderAdminID)) Then %>
				<tr>
					<td width="15%" bgcolor="#000000">&nbsp;</td>
					<td bgcolor="<%=bgcone%>"><a href="mailto:<%=strLadderAdminEmail%>">Email Admin (<%=strLadderAdminName%>)</a></td>
				</tr>
				<% End If %>
				</table>
			</td>
		</tr>
		</table>
	<% Call Content33BoxMiddle() %>
		<% If blnIsAdmin Then %>
		<center><a href="teamladderadmin.asp?team=<%=Server.URLEncode(strTeamName)%>&ladder=<%=Server.URLEncode(strLadderName)%>">Admin Panel</a></center><br />
		<%
		if strStatus="Attacking" then
						strSQL = "SELECT * FROM vMatches WHERE MatchAttackerID = " & intTLLinkID & " and matchladderid=" & intLadderID
						oRS3.Open strSQL, oConn
						If NOT(oRS3.EOF AND oRS3.BOF) Then
							intEnemyLinkID = ors3.Fields("MatchDefenderID").Value
							mDate = ors3.Fields("MatchDate").Value 
							xDate = ors3.Fields("MatchChallengeDate").Value 
							aDate = ors3.Fields("MatchAcceptanceDate").Value 
							strEnemyName = ors3.Fields("DefenderName").Value
							strMapArray(1) = ors3.Fields("MatchMap1ID").Value 
							strMapArray(2) = ors3.Fields("MatchMap2ID").Value 
							strMapArray(3) = ors3.Fields("MatchMap3ID").Value 
							strMapArray(4) = ors3.Fields("MatchMap4ID").Value 
							strMapArray(5) = ors3.Fields("MatchMap5ID").Value 
							matchdate1 = ors3.Fields("MatchSelDate1").Value
							matchdate2= ors3.Fields("MatchSelDate2").Value 
							intMatchID = ors3.Fields("MatchID").Value 
							intDefenderVotes = oRs3.Fields("DefenderVotes").Value
							intAttackerVotes = oRs3.Fields("AttackerVotes").Value
							strLastRanterName = oRs3.Fields("LastRanterName").Value
							strLastRantTime = oRs3.Fields("LastRantTime").Value
							intRants = oRs3.Fields("Rants").Value
						End If
						oRS3.NextRecordset
						Response.Write "<center><a href=""viewmatch.asp?matchId=" & intMatchID & "&Ladder=" & strLadderName & """>visit rant board</a></center><br />"
						Response.Write "<center><b>Current Status:</b></font><font size=2 color=#c0c0c0> Attacking <a href=viewteam.asp?team=" & server.urlencode(strEnemyName) & ">"&Server.HTMLEncode(strEnemyName)&"</a> <br>Maps: "
						For i = 1 to intMaps
							If i > 1 Then
								Response.Write ", "
							End If
							Response.Write Server.HTMLEncode(strMapArray(i))
						Next
						response.write "<BR>Match Date: " & mdate & "</center><br><br><center>"
						if matchdate1="TBD" then
							Response.Write "<center><b>Awaiting match acceptance from " & strEnemyName & " (Challenged on " & xDate & ")</b></center>"
						elseif mDate = "TBD" then
							if right(ucase(strEnemyName),1)="S" then
								grammer=" have "
							else
								grammer=" has "
							end if
							%>
							<table align=center width=97% border=0 cellspacing=0 CELLPADDING=0 BGCOLOR="#444444">
							<TR><TD>
								<table align=center width=100% border=0 cellspacing=1 CELLPADDING=4>
								<form name=frmAccept action=saveitem.asp method=post>
								<TR BGCOLOR="#000000"><TH><%="Challenge accepted by " & strEnemyName & " on " &  aDate%></TH></TR>
								<TR BGCOLOR="<%=bgctwo%>"><TD><%=Server.HTMLEncode(strEnemyName) & grammer%> selected the match dates listed below. Confirm your choice below</TD></TR>
								<tr BGCOLOR="<%=bgcone%>"><td align=center>Chosen Dates: <select name=matchdate class=brightred><option selected><%=matchdate1%><option><%=matchdate2%></select></td></tr>
								<tr BGCOLOR="<%=bgcone%>"><td align=center>Approved for Shoutcasting: <input type=checkbox name=scApproved value=true checked></td></tr>
								<%
								For i = 1 to Len(strMapConfiguration)
									If mid(strMapConfiguration, i, 1) = "A" Then
										'' Allow them to choose a map	
										%>
										<TR>
											<TD ALIGN=CENTER BGCOLOR=<%=bgcone%>>Choose Map <%=i%>: <SELECT Name=Map<%=i%> CLASS=bright>
											<%
											strSQL = "EXEC GetMapList '" & intMatchID & "', " & i
											oRS2.Open strsql, oconn
											if not (oRS2.EOF and oRS2.BOF) then
												do while not oRS2.EOF
													Response.Write "<option VALUE=""" & oRS2.Fields("MapName").Value & """>" & oRS2.Fields("MapName").Value & "</OPTION>" & vbCrLf
													oRS2.MoveNext 
												loop
											End If
											oRS2.NextRecordSet 
											%>
											</TD>
										</TR>
										<%
									End If
								Next
								%>				
								<tr BGCOLOR="<%=bgcone%>"><td align=center>
									<INPUT TYPE=HIDDEN NAME=MC VALUE="<%=strMapConfiguration%>">
									<input type=hidden name=matchid value=<%=intMatchID%>>
									<input type=hidden name=SaveType value=AcceptMatchDate>
									<input type=hidden name=Ladder value="<%=Server.HTMLEncode(strLadderName)%>">
									<input type=hidden name=Team value="<%=Server.HTMLEncode(strTeamName)%>">
									<input type=submit name=submit1 value="Confirm Date and Time" class=bright>
								</td></tr>
								</form>
							</table>
							</TD></TR>
							</tABLE>
							<%
					else
						strSQL = "SELECT * FROM vLadderOptions "
						strSQL = strSQL & " WHERE SelectedBy <> 'R' AND "
						strSQL = strSQL & " LadderID = '" & intLadderID & "' AND "
						strSQL = strSQL & " OptionID NOT IN (SELECT mo.OptionID FROM lnk_match_options mo WHERE mo.MatchID = '" & intMatchID & "')"
						'Response.Write strSQL
						oRs2.Open strSQL, oConn
						If Not(oRS2.EOF AND oRS2.BOF) Then
							blnOptionShown = False
							%>
							<FORM NAME=frm_map_options ACTION="/ladder/option_saveitem.asp" METHOD="POST">
							<INPUT TYPE=HIDDEN NAME="MatchID" VALUE="<%=intMatchID%>">
							<INPUT TYPE=HIDDEN NAME="SaveType" VALUE="SaveMapOptions">
							<INPUT TYPE=HIDDEN NAME="LadderName" VALUE="<%=Server.HTMLEncode(strLadderName & "")%>">
							<INPUT TYPE=HIDDEN NAME="TeamName" VALUE="<%=Server.HTMLEncode(strTeamName & "")%>">
								
							<TABLE ALIGN=CENTER WIDTH=97% BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
							<TR><TD>
							<TABLE BORDER=0 CELLSPACING=1 WIDTH=100% CELLPADDING=2>
							<TR BGCOLOR="#000000">
								<TH COLSPAN=2>Select Match Options</TH>
							</TR>
							<%
							intCounter = 0
							Do While Not(oRS2.EOF) And intCounter < 100
								intCounter = intCounter + 1
								intOptionID = oRS2.Fields("OptionID").Value 
								If bgc = bgcone Then
									bgc = bgctwo
								Else
									bgc = bgcone
								End If
								Response.Write "<TR BGCOLOR=""" & bgc & """>"
								If oRs2.Fields("MapNumber").Value <> 0 Then
									CurrentMap = strMapArray(cint(oRs2.Fields("MapNumber").Value))
								Else 
									CurrentMap = "the match"
								End If
								Select Case(oRS2.Fields("SelectedBy").Value)
									Case "A", "C"
										blnOptionSame = True
										Response.Write "<TD ALIGN=RIGHT>Choose your " & lCase(oRS2.Fields("OptionName").Value) & " for " & CurrentMap & ": </TD>"
										Response.Write "<TD><SELECT NAME=MO_" & oRS2.Fields("OptionID").Value & ">"
										While Not(oRS2.EOF) AND blnOptionSame
											Response.Write "<OPTION VALUE=""" & oRS2.Fields("OptionValueID").Value & """>" & oRS2.Fields("ValueName").Value & "</OPTION>" & vbCrLf
											oRS2.MoveNext
											If Not(oRS2.EOF) Then
												If oRS2.Fields("OptionID").Value = intOptionID Then
													blnOptionSame = True
												Else
													blnOptionSame = False
												End If
											End If
										Wend
										Response.Write "</SELECT></TD>"
										blnOptionShown = True
									Case "D"
										blnOptionSame = True
										Response.Write "<TD COLSPAN=2>" & strEnemyName & " will choose " & lcase(oRs2.Fields("OptionName").Value) & " for " & CurrentMap & "</TD>"
										While Not(oRS2.EOF) AND blnOptionSame
											oRS2.MoveNext
											If Not(oRS2.EOF) Then
												If oRS2.Fields("OptionID").Value = intOptionID Then
													blnOptionSame = True
												Else
													blnOptionSame = False
												End If
											End If
										Wend
									Case Else
										Response.Write "Error"
										oRS2.MoveNext 						
								End Select
								Response.Write "</TR>"
				'					oRS2.MoveNext
							Loop
							If blnOptionShown Then
								Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=2 ALIGN=CENTER><INPUT TYPE=SUBMIT VALUE=""Confirm Match Options""></TD></TR>"
							End If
							%>
							</TABLE>
							</TD></TR>
							</TABLE>
							</FORM>
							<%
						End If
						oRS2.NextRecordSet 
						' End Map Selection Options
				
						' Show Current Selected Options
						strSQL = "SELECT * FROM vMatchOptions "
						strSQL = strSQL & " WHERE MatchID = '" & intMatchID & "'" 
						'Response.Write strSQL
						oRs2.Open strSQL, oConn
						If Not(oRS2.EOF AND oRS2.BOF) Then
							blnOptionShown = False
							%>
							<TABLE ALIGN=CENTER WIDTH=97% BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
							<TR><TD>
							<TABLE BORDER=0 CELLSPACING=1 WIDTH=100% CELLPADDING=2>
							<TR BGCOLOR="#000000">
								<TH COLSPAN=2>Current Match Options</TH>
							</TR>
							<%
							Do While Not(oRS2.EOF)
								If bgc = bgcone Then
									bgc = bgctwo
								Else
									bgc = bgcone
								End If
								Response.Write "<TR BGCOLOR=""" & bgc & """>"
								CurrentMap = ""
								If oRs2.Fields("MapNumber").Value <> 0 Then
									CurrentMap = " on " & strMapArray(cInt(oRs2.Fields("MapNumber").Value))
								Else 
									CurrentMap = ""
								End If
									
								Select Case(oRS2.Fields("SelectedBy").Value)
									Case "A", "C"
										Response.Write "<TD>You choose " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
									Case "D"
										If oRS2.Fields("SideChoice").Value = "Y" Then
											Response.Write "<TD>You have " & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
										Else
											Response.Write "<TD>" & strEnemyName & " selected " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
										End If
									Case "R"
										If oRS2.Fields("SideChoice").Value = "Y" Then
											Response.Write "<TD>" & strTeamName & strVerbiage & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & "</TD>"
										Else
											Response.Write "<TD>TWL choose " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
										End If
									Case Else
										Response.Write "Error"
								End Select
								Response.Write "</TR>"
								oRS2.MoveNext
							Loop
							%>
							</TABLE>
							</TD></TR>
							</TABLE>
							<%
						End If
						oRS2.NextRecordSet 
						
						' End Map Selection Options
						Response.Write "<br><table width=60% align=center ><tr bgcolor="&bgcone&" height=35 valign=center><td align=center>[ <a href=MatchReportLoss.asp?matchid=" & intMatchID & "&teamid=" & intTeamID & ">Report Loss</a> ]&nbsp;&nbsp;&nbsp;&nbsp; "
						Response.Write " [ <a href=DisputeMatch.asp?matchid=" & intMatchID & "&DisputeTeamID=" & intTLLinkID & "&ladder=" & Server.URLEncode(strLadderName) & ">Dispute Match Results</a> ]</td></tr></table>"

'						Response.Write "<br><table width=45% align=center ><tr bgcolor="&bgcone&" height=35 valign=center><td align=center><a href=MatchReportLoss.asp?matchid=" & intMatchID & "&teamid=" & intTeamID & ">Report Loss</a></td></tr></table>"
					end if	
					strSQL = " SELECT COUNT(CommID) FROM tbl_comms WHERE MatchID = '" & intMatchID & "'"
					oRS2.Open strSQL, oConn
					If Not(oRS2.Eof and ORS2.BOF) Then
						Response.Write "<br /><br /><a href=""teamladderadmin.asp?team=" & Server.URLEncode(strTeamName) & "&ladder=" & Server.URLEncode(strLadderName) & "#matchcomms"">Match Comms</a> (" & oRS2.Fields(0).Value & ")"
					End If
					oRS2.NextRecordSet
	elseif strStatus="Defending" then
			strSQL = "select * FROM vMatches where MatchDefenderID = " & ors.fields("tllinkid").value & " and matchladderid=" & oRS.Fields("ladderid").value
			ors3.Open strSQL, oconn
			if not (ors3.EOF and ors3.BOF) then
				intEnemyLinkID = ors3.Fields("MatchAttackerID").Value
				mDate = ors3.Fields("MatchDate").Value 
				xDate = ors3.Fields("MatchChallengeDate").Value 
				aDate = ors3.Fields("MatchAcceptanceDate").Value 
				strEnemyName = ors3.Fields("Attackername").Value
				strMapArray(1) = ors3.Fields("MatchMap1ID").Value 
				strMapArray(2) = ors3.Fields("MatchMap2ID").Value 
				strMapArray(3) = ors3.Fields("MatchMap3ID").Value 
				strMapArray(4) = ors3.Fields("MatchMap4ID").Value 
				strMapArray(5) = ors3.Fields("MatchMap5ID").Value 
				matchdate1 = ors3.Fields("MatchSelDate1").Value
				matchdate2= ors3.Fields("MatchSelDate2").Value 
				intMatchID = ors3.Fields("MatchID").Value 
				intDefenderVotes = oRs3.Fields("DefenderVotes").Value
				intAttackerVotes = oRs3.Fields("AttackerVotes").Value
				strLastRanterName = oRs3.Fields("LastRanterName").Value
				strLastRantTime = oRs3.Fields("LastRantTime").Value
				intRants = oRs3.Fields("Rants").Value
			end if
			Response.Write "<center><a href=""viewmatch.asp?matchId=" & intMatchID & "&Ladder=" & strLadderName & """>visit rant board</a></center><br />"
	
			Response.Write "<center><font size=2><b>Current Status:</b></font><font size=2 color=#c0c0c0> Defending vs <a href=viewteam.asp?team=" & server.urlencode(strEnemyName) & ">"&Server.HTMLEncode(strEnemyName)&"</a><br>"
			Response.Write "Maps: "
			For i = 1 to intMaps
				If i > 1 Then
					Response.Write ", "
				end if
				Response.Write strMapArray(i)
			Next
			response.write " <br>Match Date: " & mdate & "</font></center><br><br><center>"
			if  (matchdate1 = "TBD") then
				Response.Write "<center><font size=2><a href=acceptmatch.asp?team=" & server.urlencode(strTeamName) & "&ladder=" & server.urlencode(strLadderName) & "&matchid=" & intMatchID & "&enemy=" & server.urlencode(strEnemyName) & ">Accept the Challenge from " & Server.HTMLEncode(strEnemyName) & "</a></font><center><center><font size=1>(You were challenged on " & xDate & ")</font></center>"
			else if mDate="TBD" then
				Response.Write "<center><font size=2>You have chosen <b>" & matchdate1 & "</b> and <b>" & matchdate2 & "</b> for match dates.</center>"
				Response.Write "<center>Awaiting acceptance from " & Server.HTMLEncode(strEnemyName) & "</font><center><center><font size=1>(You were challenged on " & xDate & ")</font></center>"
			else
			
				strSQL = "SELECT * FROM vLadderOptions "
				strSQL = strSQL & " WHERE SelectedBy <> 'R' AND "
				strSQL = strSQL & " LadderID = '" & intLadderID & "' AND "
				strSQL = strSQL & " OptionID NOT IN (SELECT mo.OptionID FROM lnk_match_options mo WHERE mo.MatchID = '" & intMatchID & "')"
				'Response.Write strSQL
				oRs2.Open strSQL, oConn
				If Not(oRS2.EOF AND oRS2.BOF) Then
					blnOptionShown = False
					%>
					<FORM NAME=frm_map_options ACTION="/ladder/option_saveitem.asp" METHOD="POST">
					<INPUT TYPE=HIDDEN NAME="MatchID" VALUE="<%=intMatchID%>">
					<INPUT TYPE=HIDDEN NAME="SaveType" VALUE="SaveMapOptions">
					<INPUT TYPE=HIDDEN NAME="LadderName" VALUE="<%=Server.HTMLEncode(strLadderName & "")%>">
					<INPUT TYPE=HIDDEN NAME="TeamName" VALUE="<%=Server.HTMLEncode(strTeamName & "")%>">
					
					<TABLE ALIGN=CENTER WIDTH=97% BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
					<TR><TD>
					<TABLE BORDER=0 CELLSPACING=1 WIDTH=100% CELLPADDING=2>
					<TR BGCOLOR="#000000">
						<TH COLSPAN=2>Select Match Options</TH>
					</TR>
					<%
					intCounter = 0
					Do While Not(oRS2.EOF) And intCounter < 100
						intCounter = intCounter + 1
						intOptionID = oRS2.Fields("OptionID").Value 
						If bgc = bgcone Then
							bgc = bgctwo
						Else
							bgc = bgcone
						End If
						Response.Write "<TR BGCOLOR=""" & bgc & """>"
						CurrentMap = ""
						If oRs2.Fields("MapNumber").Value <> 0 Then
							CurrentMap = strMapArray(cInt(oRs2.Fields("MapNumber").Value))
						Else 
							CurrentMap = "the match"
						End If
						
						Select Case(oRS2.Fields("SelectedBy").Value)
							Case "D"
								blnOptionSame = True
								Response.Write "<TD ALIGN=RIGHT>Choose your " & lCase(oRS2.Fields("OptionName").Value) & " for " & CurrentMap & ": </TD>"
								Response.Write "<TD><SELECT NAME=MO_" & oRS2.Fields("OptionID").Value & ">"
								While Not(oRS2.EOF) AND blnOptionSame
									Response.Write "<OPTION VALUE=""" & oRS2.Fields("OptionValueID").Value & """>" & oRS2.Fields("ValueName").Value & "</OPTION>" & vbCrLf
									oRS2.MoveNext
									If Not(oRS2.EOF) Then
										If oRS2.Fields("OptionID").Value = intOptionID Then
											blnOptionSame = True
										Else
											blnOptionSame = False
										End If
									End If
								Wend
								Response.Write "</SELECT></TD>"
								blnOptionShown = True
							Case "A", "C"
								blnOptionSame = True
								Response.Write "<TD COLSPAN=2>" & strEnemyName & " will choose " & lcase(oRs2.Fields("OptionName").Value) & " for " & CurrentMap & "</TD>"
								While Not(oRS2.EOF) AND blnOptionSame
									oRS2.MoveNext
									If Not(oRS2.EOF) Then
										If oRS2.Fields("OptionID").Value = intOptionID Then
											blnOptionSame = True
										Else
											blnOptionSame = False
										End If
									End If
								Wend
							Case Else
								Response.Write "Error"
								oRS2.MoveNext 						
						End Select
						Response.Write "</TR>"
	'					oRS2.MoveNext
					Loop
					If blnOptionShown Then
						Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=2 ALIGN=CENTER><INPUT TYPE=SUBMIT VALUE=""Confirm Match Options""></TD></TR>"
					End If
					%>
					</TABLE>
					</TD></TR>
					</TABLE>
					</FORM>
					<%
				End If
				oRS2.NextRecordset 
				' End Map Selection Options
	
				' Show Current Selected Options
	
				strSQL = "SELECT * FROM vMatchOptions "
				strSQL = strSQL & " WHERE MatchID = '" & intMatchID & "'" 
				'Response.Write strSQL
				oRs2.Open strSQL, oConn
				If Not(oRS2.EOF AND oRS2.BOF) Then
					blnOptionShown = False
					%>
					<TABLE ALIGN=CENTER WIDTH=97% BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
					<TR><TD>
					<TABLE BORDER=0 CELLSPACING=1 WIDTH=100% CELLPADDING=2>
					<TR BGCOLOR="#000000">
						<TH COLSPAN=2>Current Match Options</TH>
					</TR>
					<%
					Do While Not(oRS2.EOF)
						If bgc = bgcone Then
							bgc = bgctwo
						Else
							bgc = bgcone
						End If
						Response.Write "<TR BGCOLOR=""" & bgc & """>"
						CurrentMap = ""
						If oRs2.Fields("MapNumber").Value <> 0 Then
							CurrentMap = " on " & strMapArray(cInt(oRs2.Fields("MapNumber").Value))
						Else 
							CurrentMap = ""
						End If
						
						Select Case(oRS2.Fields("SelectedBy").Value)
							Case "D"
								Response.Write "<TD>You choose " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
							Case "A", "C"
								If oRS2.Fields("SideChoice").Value = "Y" Then
									Response.Write "<TD>You have " & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
								Else
									Response.Write "<TD>" & strEnemyName & " selected " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
								End If
							Case "R"
								If oRS2.Fields("SideChoice").Value = "Y" Then
									Response.Write "<TD>" & strTeamName & strVerbiage & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & "</TD>" 
								Else
									Response.Write "<TD>TWL choose " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
								End If
							Case Else
								Response.Write "Error"
						End Select
						Response.Write "</TR>"
						oRS2.MoveNext
					Loop
					%>
					</TABLE>
					</TD></TR>
					</TABLE>
					<%
				End If
				oRS2.NextRecordset 
				' End Map Selection Options
				
				Response.Write "<br><table width=60% align=center ><tr bgcolor="&bgcone&" height=35 valign=center><td align=center>[ <a href=MatchReportLoss.asp?matchid=" & intMatchID & "&teamid=" & intTeamID & ">Report Loss</a> ]&nbsp;&nbsp;&nbsp;&nbsp; "
				Response.Write " [ <a href=DisputeMatch.asp?matchid=" & intMatchID & "&DisputeTeamID=" & intTLLinkID & "&ladder=" & Server.URLEncode(strLadderName) & ">Dispute Match Results</a> ]</td></tr></table>"
'				Response.Write "<br><table width=45% align=center ><tr bgcolor="&bgcone&" height=35 valign=center><td align=center><a href=MatchReportLoss.asp?matchid=" & intMatchID & "&teamid=" & intTeamID & ">Report Loss</a></td></tr></table>"
			end if
		end if
		oRS3.NextRecordSet
		strSQL = " SELECT COUNT(CommID) FROM tbl_comms WHERE MatchID = '" & intMatchID & "'"
		oRS2.Open strSQL, oConn
		If Not(oRS2.Eof and ORS2.BOF) Then
				Response.Write "<br /><br /><a href=""teamladderadmin.asp?team=" & Server.URLEncode(strTeamName) & "&ladder=" & Server.URLEncode(strLadderName) & "#matchcomms"">Match Comms</a> (" & oRS2.Fields(0).Value & ")"
		End If
		oRS2.NextRecordSet
	elseif (strStatus="Available" or left(strStatus,6)="Immune") then
		intLadderLocked = oRS.Fields("LadderLocked").Value
		intMinPlayer = oRS.Fields("MinPlayer").Value

		strsql= "select count(TPLLinkID) from lnk_T_P_L where TLLinkID ='" & intTLLinkID & "'"
		ors3.open strsql, oconn
		if not (ors3.eof and ors3.bof) then
			players = ors3.fields(0).value
		end if
		ors3.close
		If IsNull(intMinPlayer) then
			intMinPlayer = 0
		End If
		if (intLadderLocked = 0) and players >= intminplayer then
			Response.Write "<div align=center><a href=""teamladderadmin.asp?ladder=" & Server.URLEncode(oRS.Fields("Laddername").Value) & "&team=" & Server.URLEncode(oRS.Fields("Teamname").Value) & """>Click here to initiate a challenge</a></div>"
		else
			if intLadderLocked = 1 then 
				Response.Write "<div align=center><b><i>This ladder is not open for challenging at this time.</i></b></div>"
			end if
			if Players <= intminplayer then
				Response.Write "<div align=center><b><i>Your roster is smaller than " & intminplayer & " people, you cannot challenge another team at this time.</i></b></div>"
			end if
			
		end if
	End If
'''-----------------
' End of Admin stuff
'''-----------------
Else 
				strStatus = ors.Fields("Status").Value
				strTeamName = oRS.Fields("TeamName").Value
					Select Case Left(uCase(strStatus), 6)
						Case "DEFEND", "ATTACK"
							If  Left(uCase(strStatus), 6)  = "DEFEND" Then
								strsql = "select m.MatchAttackerID, m.MatchID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID, MatchMap4ID, MatchMap5ID, t.teamname, t.teamtag, DefenderVotes, AttackerVotes, Rants, LastRantTime, LastRanterName "
								strsql = strsql & " from tbl_Matches m, tbl_teams t, lnk_t_l lnk "
								strsql = strsql & " where m.matchdefenderID = " & oRS.Fields("TLLinkID").Value  
								strsql = strsql & " AND t.teamid = lnk.teamid "
								strsql = strsql & " AND lnk.tllinkid = m.MatchAttackerID "
								oRS2.Open strSQL, oconn
								if not (oRS2.EOF and oRS2.BOF) then
									intMatchID = oRS2.Fields("MatchID").Value 
									strEnemyName = oRS2.Fields("TeamName").Value
									strMatchDate = oRS2.Fields("MatchDate").Value
									strMapArray(1) = oRS2.fields("MatchMap1ID").value
									strMapArray(2) = oRS2.fields("MatchMap2ID").value
									strMapArray(3) = oRS2.fields("MatchMap3ID").value
									strMapArray(4) = oRS2.fields("MatchMap4ID").value
									strMapArray(5) = oRS2.fields("MatchMap5ID").value
									intDefenderVotes = oRs2.Fields("DefenderVotes").Value
									intAttackerVotes = oRs2.Fields("AttackerVotes").Value
									strLastRanterName = oRs2.Fields("LastRanterName").Value
									strLastRantTime = oRs2.Fields("LastRantTime").Value
									intRants = oRs2.Fields("Rants").Value
								end if
								Response.Write "&nbsp;&nbsp;&nbsp;<b>Defending:</b>&nbsp;<a href=viewteam.asp?team=" & server.urlencode(strEnemyName) & ">" & Server.HTMLEncode(strEnemyName)&"</a>"
								Response.Write "<br><br>"
								Response.Write "<table border=0 cellspacing=0 cellpadding=0 width=""90%"" align=""center"">"
								Response.Write "<tr><td colspan=2><a href=""viewmatch.asp?matchId=" & intMatchID & "&Ladder=" & strLadderName & """>visit rant board</a></td></tr>"
								Response.Write "<tr><td colspan=2>&nbsp;</td></tr>"
								response.write "<tr><td colspan=2><b>Match Date:</b><br>" & strMatchDate & "</td></TR>"
								ors2.NextRecordset 

								If Right(uCase(strTeamName), 1) = "S" Then
									strVerbiage = " have "
								Else
									strVerbiage = " has "
								End IF
								
								If Right(uCase(strEnemyName), 1) = "S" Then
									strEnemyVerbiage = " have "
								Else
									strEnemyVerbiage = " has "
								End IF
								Response.Write "<tr><td>&nbsp;</td></tr>"
								' Show Current Selected Options
								strSQL = "SELECT * FROM vMatchOptions "
								strSQL = strSQL & " WHERE MatchID = '" & intMatchID & "'" 
								'Response.Write strSQL
								oRs2.Open strSQL, oConn
								If Not(oRS2.EOF AND oRS2.BOF) Then
									%>
									<tr><td colspan=2><B>Maps and Match Options:</B></td></tr>
									<tr><td colspan=2><img src="/images/spacer.gif" height="3"></td></tr>
									<%
									CurrentMap = ""
									CurrentMapNumber = -1
									Do While Not(oRS2.EOF)
										If CurrentMapNumber <> oRS2.Fields("MapNumber").Value Then
											CurrentMapNumber = oRS2.Fields("MapNumber").Value
											If CurrentMapNumber <> 0 Then
												CurrentMap = strMapArray(CurrentMapNumber)
											End If
										End If
										Response.Write "<tr><td width=""30%""><b>" & CurrentMap & "</b></td><td>"
										Select Case(oRS2.Fields("SelectedBy").Value)
											Case "D"
												If oRS2.Fields("SideChoice").Value = "Y" Then
													Response.Write strTeamName & strVerbiage & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & "<BR>"
												Else
													Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value & "<BR>"
												End If
											Case "A", "C"
												If oRS2.Fields("SideChoice").Value = "Y" Then
													Response.Write strTeamName & strVerbiage & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & "<BR>"
												Else
													Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value & "<BR>"
												End If
											Case "R"
												If oRS2.Fields("SideChoice").Value = "Y" Then
													Response.Write strTeamName & strVerbiage & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & "<BR>"
												Else
													Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value & "<BR>"
												End If
											Case Else
												Response.Write "Error<BR>"
										End Select
										Response.Write "</td></tr>"
										oRS2.MoveNext
									Loop
								Else
									Response.Write "<tr><td><B>Maps:</b> <br>"
									For i = 1 to intMaps
										Response.Write strMapArray(i)
										If i <> intMaps Then
											Response.Write ", "
										End If
									Next
								End If
								oRS2.NextRecordset 
								Response.write "</TD></TR>"
								Response.Write "</table>"
							Else
								strsql = "select m.MatchDefenderID, m.MatchID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID, MatchMap4ID, MatchMap5ID, t.teamname, t.teamtag, DefenderVotes, Rants, AttackerVotes, LastRantTime, LastRanterName "
								strsql = strsql & " from tbl_Matches m, tbl_teams t, lnk_t_l lnk "
								strsql = strsql & " where m.MatchAttackerID = " & oRS.Fields("TLLinkID").Value 
								strsql = strsql & " AND t.teamid = lnk.teamid "
								strsql = strsql & " AND lnk.tllinkid = m.matchdefenderID "
								oRS2.Open strSQL, oconn
								if not (oRS2.EOF and oRS2.BOF) then
									intMatchID = oRS2.Fields("MatchID").Value 
									strEnemyName = oRS2.Fields("TeamName").Value
									strMatchDate = oRS2.Fields("MatchDate").Value
									strMapArray(1) = oRS2.fields("MatchMap1ID").value
									strMapArray(2) = oRS2.fields("MatchMap2ID").value
									strMapArray(3) = oRS2.fields("MatchMap3ID").value
									strMapArray(4) = oRS2.fields("MatchMap4ID").value
									strMapArray(5) = oRS2.fields("MatchMap5ID").value
									intDefenderVotes = oRs2.Fields("DefenderVotes").Value
									intAttackerVotes = oRs2.Fields("AttackerVotes").Value
									strLastRanterName = oRs2.Fields("LastRanterName").Value
									strLastRantTime = oRs2.Fields("LastRantTime").Value
									intRants = oRs2.Fields("Rants").Value
								end if
								Response.Write "&nbsp;&nbsp;&nbsp;<b>Attacking:</b>&nbsp;<a href=viewteam.asp?team=" & server.urlencode(strEnemyName) & ">" & Server.HTMLEncode(strEnemyName)&"</a>"
								Response.Write "<br><br>"
								Response.Write "<table border=0 cellspacing=0 cellpadding=0 width=""90%"" align=""center"">"
								Response.Write "<tr><td colspan=2><a href=""viewmatch.asp?matchId=" & intMatchID & "&Ladder=" & strLadderName & """>visit rant board</a></td></tr>"
								Response.Write "<tr><td colspan=2>&nbsp;</td></tr>"
								response.write "<tr><td colspan=2><b>Match Date:</b><br>" & strMatchDate & "</td></TR>"
								Response.Write "<tr><td>&nbsp;</td></tr>"
								ors2.NextRecordset 

								If Right(uCase(strTeamName), 1) = "S" Then
									strVerbiage = " have "
								Else
									strVerbiage = " has "
								End IF
								
								If Right(uCase(strEnemyName), 1) = "S" Then
									strEnemyVerbiage = " have "
								Else
									strEnemyVerbiage = " has "
								End IF
																
								' Show Current Selected Options
								strSQL = "SELECT * FROM vMatchOptions "
								strSQL = strSQL & " WHERE MatchID = '" & intMatchID & "'" 
								'Response.Write strSQL
								oRs2.Open strSQL, oConn
								If Not(oRS2.EOF AND oRS2.BOF) Then
									%>
									<tr><td colspan=2><B>Maps and Match Options:</B></td></tr>
									<tr><td colspan=2><img src="/images/spacer.gif" height="3"></td></tr>
									<%
									CurrentMap = ""
									CurrentMapNumber = -1
									Do While Not(oRS2.EOF)
										If CurrentMapNumber <> oRS2.Fields("MapNumber").Value Then
											CurrentMapNumber = oRS2.Fields("MapNumber").Value
											If CurrentMapNumber <> 0 Then
												CurrentMap = strMapArray(CurrentMapNumber)
											End If
										End If
										Response.Write "<tr><td width=""30%""><b>" & CurrentMap & "</b></td><td>"
										Select Case(oRS2.Fields("SelectedBy").Value)
											Case "A", "C"
												If oRS2.Fields("SideChoice").Value = "Y" Then
													Response.Write strTeamName & strVerbiage & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & "<BR>"
												Else
													Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value & "<BR>"
												End If
											Case "D"
												If oRS2.Fields("SideChoice").Value = "Y" Then
													Response.Write strTeamName & strVerbiage & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & "<BR>"
												Else
													Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value & "<BR>"
												End If
											Case "R"
												If oRS2.Fields("SideChoice").Value = "Y" Then
													Response.Write strTeamName & strVerbiage & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & "<BR>"
												Else
													Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value & "<BR>"
												End If
											Case Else
												Response.Write "Error<BR>"
										End Select
										Response.Write "</td></tr>"
										oRS2.MoveNext
									Loop
								Else
									Response.Write "<tr><td><B>Maps:</b> <br>"
									For i = 1 to intMaps
										Response.Write strMapArray(i)
										If i <> intMaps Then
											Response.Write ", "
										End If
									Next
									Response.Write "</td></tr>"
								End If
								oRS2.NextRecordset 
								Response.Write "</table>"
							End If
						Case "IMMUNE", "DEFEAT", "RESTIN"
							Response.Write strStatus
						Case Else
							Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Available to challenge or be challenged"
					End Select

	'----------------------
	' End not an admin section
	'----------------------
	End If 
				oRS.MoveNext
				If Not(oRS.EOF) Then
					Call Content33BoxEnd()
					Call Content33BoxStart("") 
				End If
			Loop
		Call Content33BoxEnd()
		End If
	End If
	Set oRS = oRS.NextRecordSet		
	'Response.Flush

	If oRS.State = 1 Then
		'----------------------------------
		'' Power Ladders
		'----------------------------------
		If Not(oRs.BOF And oRs.EOF) Then
			Call Content33BoxStart("Power Ladder Competition Status") 
			Do While Not(oRs.EOF)
				strTeamName = oRS.Fields("TeamName").Value
				strTeamTag = oRS.Fields("TeamTag").Value
				intTeamID = oRS.Fields("TeamID").Value
				%>
				<table border="0" cellspacing="0" cellpadding="0" align="center" width="97%" class="cssBordered">
				<tr>
					<th colspan="2"><a href="viewscrimladder.asp?ladder=<%=Server.URLEncode(oRs.Fields("EloLadderName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("EloLadderName").Value & "")%> Ladder</a></th>
				</tr>
				<tr>
					<td bgcolor="<%=bgcone%>" align="right">Team:</td>
					<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(strTeamName & "")%>"><%=Server.HTMLEncode(strTeamName & "")%></a></td>
				</tr>
				<tr>
					<td bgcolor="<%=bgctwo%>" align="right">Rating:</td>
					<td bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(ors.Fields("Rating").Value & "")%></td>
				</tr>
				<tr>
					<td bgcolor="<%=bgcone%>" align="right">W/L:</td>
					<td  bgcolor="<%=bgcone%>"><%=oRs.Fields("Wins").Value%>/<%=oRs.Fields("Losses").Value%></td>
				</tr>
				<tr>
					<td width="15%" bgcolor="#000000">&nbsp;</td>
					<td bgcolor="<%=bgctwo%>"><a href="rules.asp?set=<%=Server.URLEncode(oRs.Fields("EloRulesName").Value & "")%>">View Ladder Rules</a></td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td bgcolor="<%=bgctwo%>"><a href="forums/forumdisplay.asp?forumid=<%=Server.URLEncode(oRs.Fields("GameForumID").Value & "")%>">Game Forum</a></td>
				</tr>
				</table>
				<% Call Content33BoxMiddle()%>
				<table border="0" cellspacing="0" cellpadding="0" width="97%" class="cssBordered" align="center">
				<tr>
					<th colspan="7">Pending Matches</th>
				</tr>
				<%
				Dim strOpponentName, strOpponentTag, intOpponentRanking
				strSQL = "SELECT EloMatchID, DefenderEloTeamID, AttackerEloTeamID, ChallengeDate, MatchDate, EloLadderID, Map1, Map2, Map3, Map4, Map5, LastComm = (SELECT TOP 1 CommDate FROM tbl_elo_comms ec WHERE ec.EloMatchID = m.EloMatchID ORDER BY EloCommID DESC) "
				strSQL = strSQL & " FROM tbl_elo_matches m WHERE (DefenderEloTeamID = '" & oRs.Fields("lnkEloTeamID").Value & "' OR AttackerEloTeamID = '" & oRs.Fields("lnkEloTeamID").Value & "') AND MatchActive = 1 ORDER BY ChallengeDate ASC"
				oRs3.Open strSQL, oConn
				If Not(oRs3.EOF AND oRs3.BOF) Then
					Do While Not(oRs3.EOF)
						If bgc = bgcone Then
							bgc = bgctwo
						Else
							bgc = bgcone
						End If
						
						If oRs3.Fields("DefenderEloTeamID").Value = oRs.Fields("lnkEloTeamID").Value Then
							strSQL = "SELECT TeamName, TeamTag, Rating FROM tbl_teams t INNER JOIN lnk_elo_team et ON et.TeamID = t.TeamID WHERE et.lnkEloTeamID = '" & oRs3.Fields("AttackerEloTeamID").Value & "'"
						Else
							strSQL = "SELECT TeamName, TeamTag, Rating FROM tbl_teams t INNER JOIN lnk_elo_team et ON et.TeamID = t.TeamID WHERE et.lnkEloTeamID = '" & oRs3.Fields("DefenderEloTeamID").Value & "'"
						End If
						oRs2.Open strSQL, oConn
						If Not(oRs2.EOF AND oRs2.BOF) Then
							strOpponentName =  oRs2.Fields("TeamName").Value
							strOpponentTag = oRs2.Fields("TeamTag").Value
							intOpponentRanking = oRs2.Fields("Rating").Value
						End If
						oRs2.NextRecordSet						
						%>
						<tr>
							<td bgcolor="<%=bgc%>"><a href="/viewteam.asp?team=<%=Server.URLEncode(strOpponentName & "")%>"><%=Server.HTMLEncode(strOpponentName & "")%> (<%=intOpponentRanking%>)</a></td>
							<td bgcolor="<%=bgc%>" align="center"><%=FormatDateTime(oRs3.Fields("ChallengeDate").Value, 2)%></td>
							<td bgcolor="<%=bgc%>"><%
								If IsDate(oRs3.FieldS("MatchDate").Value) Then
									Response.Write FormatDateTime(oRs3.Fields("MatchDate").Value, 0)
								Else
									Response.Write "Unscheduled"
								End If
								%></td>
							<td bgcolor="<%=bgc%>"><%
								If Not(IsNull(oRs3.Fields("Map1").Value) OR Len(oRs3.Fields("Map1").Value) = 0) Then
									Response.Write oRs3.Fields("Map1").Value
								End If
								If Not(IsNull(oRs3.Fields("Map2").Value) OR Len(oRs3.Fields("Map2").Value) = 0) Then
									Response.Write ", " & oRs3.Fields("Map2").Value
								End If
								If Not(IsNull(oRs3.Fields("Map3").Value) OR Len(oRs3.Fields("Map3").Value) = 0) Then
									Response.Write ", " & oRs3.Fields("Map3").Value
								End If
								If Not(IsNull(oRs3.Fields("Map4").Value) OR Len(oRs3.Fields("Map4").Value) = 0) Then
									Response.Write ", " & oRs3.Fields("Map4").Value
								End If
								If Not(IsNull(oRs3.Fields("Map5").Value) OR Len(oRs3.Fields("Map5").Value) = 0) Then
									Response.Write ", " & oRs3.Fields("Map5").Value
								End If
								%></td>
							<% If oRs.FieldS("IsAdmin").Value Then %>
								<td bgcolor="<%=bgc%>" align="center"><%
									If IsDate(oRs3.FieldS("LastComm").Value) Then
										Response.Write FormatDateTime(oRs3.Fields("LastComm").Value, 0)
									Else
										Response.Write " never "
									End If
									%></td>
									<td bgcolor="<%=bgc%>" align="center"><a href="teamscrimladderadmin.asp?ladder=<%=Server.URLEncode(oRs.Fields("EloLadderName").Value & "")%>&team=<%=Server.URLEncode(strTeamName)%>&matchid=<%=oRs3.Fields("EloMatchID").Value%>">manage</a></td>
							<% End If %>
						</tr>
						<%
						oRs3.MoveNext
					Loop
				Else 
					%>
					<tr>
						<td colspan="7"><i>No pending matches</i></td>
					</tr>
					<%
				End If
				oRs3.NextRecordSet
				%>
				<% If oRs.Fields("IsAdmin").Value = 1 Then %>
					<tr>
						<td align="center" colspan="7"><a href="teamscrimladderadmin.asp?ladder=<%=Server.URLEncode(oRs.Fields("EloLadderName").Value & "")%>&team=<%=Server.URLEncode(strTeamName)%>">Power Ladder Admin Panel</a></td>
					</tr>
					<% End If %>
				</table>
				<%
				oRs.MoveNext
				If Not (oRs.EOF) Then
					Call Content33BoxEnd() 
					Call Content33BoxStart("")
				End If
			Loop
			Call Content33BoxEnd() 
		End If
		'----------------------------------
		'' End Power Ladders
		'----------------------------------
	End If
	Set oRs = oRs.NextRecordSet

	If oRS.State = 1 Then
		'----------------
		' Player Ladders
		'----------------
		If Not(oRs.Eof and oRs.boF) Then
			'' PLayer Ladders
			Call ContentStart("Player Competition Status")
			intPlayerID=session("playerid")
			%>
			<table align=center border=0 cellspacing="0" cellpadding="0" class="cssBordered" width="97%">
			<tr bgcolor="#000000">
				<th width=150>Ladder Name</th>
				<th width=50>Rank</th>
				<th width=75>Record</th>
				<th width=300>Status</th>
				<th>&nbsp;</th>
				<th>&nbsp;</th>
			</tr>
			<%
			bgc = bgcone
			Do While Not(oRs.Eof)
				intPlayerLadderID = oRs("PlayerLadderID")
				strLadderName = oRs("PlayerLadderName")
				intRank = oRs("Rank")
				intLosses = oRs("Losses")
				intforfeits = oRs("forfeits")
				intwins = oRs("wins")
				strStatus = oRs("status")
				PPLLinkID = oRs("PPLLinkID")
						
				Select Case(uCase(strStatus))
					Case "ATTACKING"
						strSQL = "SELECT p.PlayerHandle, m.MatchMap1ID, m.MatchDate "
						strSQL = strSQL & " FROM lnk_p_pl lnk, tbl_players p, tbl_PlayerMatches m "
						strSQL = strSQL & " WHERE lnk.PlayerID = p.PlayerID AND m.MatchDefenderID = lnk.PPLLinkID "
						strSQL = strSQL & " AND m.MatchAttackerID = " & PPLLinkID
						oRS2.open strsql, oconn
						If not(oRS2.eof and oRS2.bof) then
							map = oRS2("matchMap1ID")
							opponent = oRS2("PlayerHandle")
							mDate = oRS2("MatchDate")
							statusVerbage = strStatus & " vs. <a href=viewplayer.asp?player=" & server.urlencode(opponent) & ">" & opponent & "</A> (" & map & ")<BR>" & mDate
						Else
							statusVerbage = " Data Error "
						End If
						oRS2.NextRecordset
					Case "DEFENDING"
						strSQL = "SELECT p.PlayerHandle, m.MatchMap1ID, m.MatchDate "
						strSQL = strSQL & " FROM lnk_p_pl lnk, tbl_players p, tbl_PlayerMatches m "
						strSQL = strSQL & " WHERE lnk.PlayerID = p.PlayerID AND m.MatchAttackerID = lnk.PPLLinkID "
						strSQL = strSQL & " AND m.MatchDefenderID = " & PPLLinkID
						oRS2.open strsql, oconn
						If not(oRS2.eof and oRS2.bof) then
							map = oRS2("matchMap1ID")
							opponent = oRS2("PlayerHandle")
							mDate = oRS2("MatchDate")
							statusVerbage = strStatus & " vs. <a href=viewplayer.asp?player=" & server.urlencode(opponent) & ">" & opponent & "</A> (" & map & ")<BR>" & mDate
						Else
							statusVerbage = " Data Error "
						End if
						oRS2.NextRecordset 
					Case Else
						statusVerbage = strStatus
				End Select
				%>
				<TR>
					<TD BGCOLOR=<%=bgc%>><A href="viewplayerladder.asp?ladder=<%=server.URLEncode(strLadderName)%>"><%=strLadderName%></A></TD>
					<TD ALIGN=CENTER BGCOLOR=<%=bgc%>><%=intRank%></TD>
					<TD ALIGN=CENTER BGCOLOR=<%=bgc%>><%=intWins & "/" & intLosses & " (" & intForFeits & ") "%></TD>
					<TD ALIGN=CENTER BGCOLOR=<%=bgc%>><%=statusVerbage%></TD>
					<TD ALIGN=CENTER BGCOLOR=<%=bgc%>><a href=playerladderadmin.asp?player=<%=server.URLEncode(strPlayerName)%>&ladder=<%=server.URLEncode(strLadderName)%>>Admin</A></TD>
					<TD ALIGN=CENTER BGCOLOR=<%=bgc%>"><a href="javascript:popup('playerquitLadder.asp?playerid=<%=intPlayerID%>&ladder=<%=server.urlencode(strLadderName)%>&url=viewplayer.asp?player=<%=server.urlencode(strPlayerName)%>', 'quitladder', 150, 300, 'no');">Quit</A></TD>
				</TR>
				<%
				If bgc = bgcone then
					bgc = bgctwo
				Else
					bgc = bgcone
				End If
				oRs.MoveNext
			Loop
			%>
			</table>
			<%
			Call ContentEnd()
		End If
	End If
	Set oRs = oRs.NextRecordSet
	'Response.Flush 

End If
'' LOGGED IN

	intNewsID = 0
	intNewsCnt = 1
	If oRs.State = 1 Then
		If Not(oRS.EOF and oRS.BOF) Then
			Do While Not(oRS.EOF)
				If intNewsCnt = 1 Then
					strCurrDate = FormatDateTime(oRS.Fields("NewsDate").Value, vbShortDate)
					Call ContentNewsStart(weekdayname(weekday(strCurrDate)) & ", " & monthname(month(strCurrDate)) & " " & day(strCurrDate))
				ElseIf (strCurrDate <> FormatDateTime(oRS.Fields("NewsDate").Value, vbShortDate)) then
					Call ContentNewsEnd()
					strCurrDate = FormatDateTime(oRS.Fields("NewsDate").Value, vbShortDate)
					Call ContentNewsStart(weekdayname(weekday(strCurrDate)) & ", " & monthname(month(strCurrDate)) & " " & day(strCurrDate))
				Else
					Response.Write "<tr><td><hr></td></tr>"
				End If
				%>
				<tr><td>
					<a name="<%=ors.Fields("NewsId").Value %>" />
					<table width="100%" align=center border=0 cellpadding="0">
					   <tr valign="top"> 
					    <td class="newsheader"><%
					    If Len(oRs.Fields("GameAbbreviation").Value) > 0 Then
					    	Response.Write Server.HTMLEncode (oRs.Fields("GameAbbreviation").Value) & ": "
					    End If 
					    Response.Write Server.HTMLEncode (ors.fields("NewsHeadLine").value)
					    %></td>
					    <td ALIGN=RIGHT>Posted by <a href="viewplayer.asp?player=<%=Server.URLEncode(ors.fields("NewsAuthor").value)%>"><%=Server.HTMLEncode(ors.fields("NewsAuthor").value)%><br>
					        </a><%=formatdatetime(oRS.Fields("NewsDate").Value, 4)%>
					    </td>
					  </tr>
					  <TR>
						<TD COLSPAN=2 style="padding: 0px, 20px, 0px, 20px;"><%=ors.fields("NewsContent").value%></td>
					  </tr>
					</table>
				</td></tr>
				<%
				intNewsCnt = intNewsCnt + 1
				oRS.MoveNext 
			Loop
			Call ContentNewsEnd()
		End If
	End If
	oRS.Close
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oRS = Nothing
Set oConn = Nothing
%>