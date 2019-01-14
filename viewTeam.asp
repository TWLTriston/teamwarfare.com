<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("team"), """", "&quot;") 

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin, bTournamentAdmin 
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim MapName
Dim strTeamName, intFounderID, intTeamID, strURL, strFounderName
Dim bShownAdmin, intMaps
Dim intURLID, intNumDone
Dim bOnALeague, bOnATournament, bOnTeam
Dim intRosterLimit, strLadderName, intMatchID, intLadderID
Dim strStatus, strMatchDate, strMap1, strMap2, strMap3
Dim strDateJoined, strAdmin, intRosterCount
Dim strSide1, strSide2, strSide3
Dim TLLinkID, strEnemyName, strResult
'Tournament var
Dim RoundsID, Team1, Team1ID, Team2ID, OpponentLinkId
Dim Team1Name, Team2Name, LocationVerb, OpponentName
Dim RoundNum, Tournament, TMLinkID, CurrentStatus
Dim LinkID, strVerbiage, strEnemyVerbiage, CurrentMap, CurrentMapNumber
Dim strMapArray(6), i
bShownAdmin = False

strTeamName = Request.QueryString("Team")
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
strSQL = "SELECT t.*, p.PlayerHandle "
strSQL = strSQL & " FROM tbl_Teams t, tbl_players p "
strSQL = strSQL & " WHERE  t.TeamFounderID *= p.PlayerID "
strSQL = strSQL & " AND TeamName='" & CheckString(strTeamName) & "'"
oRS.Open strSQL, oConn
If oRS.EOF AND oRS.BOF Then 
	Call ContentStart("Invalid Team Name")
	%>
	<FONT COLOR="red">Invalid team name specified, check linking URL.</FONT>
	<%
	Call ContentEnd()
Else
	intTeamID = oRs.Fields("TeamID").Value 
	strFounderName = oRS.Fields("PlayerHandle").Value
	strTeamName = oRs.Fields("TeamName").Value
	Call Content66BoxStart(Server.HTMLEncode(strTeamName) & " Team Profile")
		strURL = oRS.Fields("TeamURL").Value
		If ucase(left(strURL,4))<> "HTTP" AND Len(strURL) > 0 Then
			strURL = "http://" & strURL
		End If
		%>
		<table cellspacing=0 cellpadding=0 border=0 width=97% class="cssBordered" align="center">
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2><%=Server.HTMLEncode(strTeamName)%></TH>
		</TR>
		<tr bgcolor=<%=bgctwo%>>
			<td align=right WIDTH=20%>Tags:</td><td><%=Server.HTMLEncode(ors.Fields("TeamTag").Value)%>  </td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>URL:</td>
			<td><a href="<%=Server.HTMLEncode(strURL & "") %>" target="_blank"><%=Server.HTMLEncode(strURL & "")%></a>&nbsp;</td></tr>
		<tr bgcolor=<%=bgctwo%>><td align=right>Email:</td>
								<td bgcolor="<%=bgc%>" valign="top"><%=Replace(Replace("" & oRs.Fields("TeamEmail").Value, "@", " at "), ".", " dot ")%></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right NOWRAP>IRC Channel:</td><td><% = Server.HTMLEncode(oRs.Fields("TeamIRC").value & "")%>&nbsp;</td></tr>
		<tr bgcolor=<%=bgctwo%>><td align=right>IRC Server:</td><td><% = Server.HTMLEncode(oRS.Fields("TeamIRCServer").Value  & "")%>&nbsp;</td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Status:</td><td> <% 
		If ors.Fields("TeamActive").Value= "1" then
			Response.Write "Active"
		Else
			Response.Write "Inactive"
		End If
		%></TD></TR>
		<tr bgcolor=<%=bgctwo%>>
			<td align=right>Founder:</td>
			<td><a href=viewplayer.asp?Player=<% = Server.URLEncode(oRS.Fields("PlayerHandle").Value & "")%>><% = Server.HTMLEncode(oRS.Fields("PlayerHandle").Value  & "")%></a></td>
		</tr>
		<tr bgcolor=<%=bgcone%>>
			<td align=right valign=top>Description:</td>
			<td><% = ForumEncode2(oRS.Fields("TeamDesc").Value) %></td>
		</tr>
		<TR BGCOLOR="#000000" VALIGN="MIDDLE">
			<TD COLSPAN=2 ALIGN="CENTER">
				<a href="history.asp?Keydata=<% = server.URLEncode(strTeamName & "")%>">Match History</a> / <A HREF="/xml/viewteam_v2.asp?team=<%=Server.UrlEncode(strTeamName)%>">Team Profile XML Version</A>
			</td>
		</tr>
		</table>
	<% Call Content66BoxMiddle() %>
		<table width=250 align=center valign=center border=0 cellspacing="0" cellpadding= "0">
			<tr><td align=center>
			<% 
			If Len(Trim(oRS.Fields("TeamLogoURL").Value)) > 0  then
				Response.Write "<img src='" & oRS.Fields("TeamLogoURL").Value & "' width=200 height=200>"
			else
				Response.Write "<img src='images/nopicture.jpg' width=200 height=200>"
			end if
			%>
			</td></tr>
		</table>
	<%
	oRS.Close
	Call Content66BoxEnd() 
	
	'-------------------
	' Begin Admin Tools
	'-------------------
	Sub ShowAdminPanel ()
		Call ContentStart("")
		%>
			<table border="0" cellspacing="0" cellpadding="0" align="center" class="cssBordered">
			<form name="frmAdminStuff" id="frmAdminStuff">
			<tr>
				<th colspan="3">Team Administration</th>
			</tr>
		<%
	End Sub
	bgc = bgcone
	If Session("LoggedIn") Then
		bTeamFounder = IsTeamFounder(strTeamName)
		If bTeamFounder Or bSysAdmin Then
			bShownadmin = True
			Call ShowAdminPanel()
			%>
			<tr>
				<td>&nbsp;</td>
				<td><a href="join.asp?team=<%=Server.URLEncode(strTeamName)%>">Join New Competition</a></td>
				<td align="center"><a href="addteam.asp?isedit=true&team=<%=Server.URLEncode(strTeamName)%>">Edit Team Information</a></td>
			</tr>
			<%
		End If
		
		' League Lookup
		strSQL = "SELECT l.LeagueName, lnk.lnkLeagueTeamID, l.LeagueID "
		strSQL = strSQL & " FROM tbl_leagues l, lnk_league_team lnk "
		strSQL = strSQL & " WHERE l.LeagueID = lnk.LeagueID "
		strSQL = strSQL & " AND lnk.TeamID = '" & intTeamID & "'" 
		strSQL = strSQL & " AND lnk.Active = 1 "
		strSQL = strSQL & " AND l.LeagueActive = 1 "
		strSQL = strSQL & " ORDER BY l.LeagueName "
		oRS.Open strSQL, oConn
		bgc = bgctwo
		if not (ors.EOF and ors.BOF) then
			intURLID = 0
			intNumDone = 0
			Do While Not oRS.EOF
				If bSysAdmin OR bTeamFounder OR IsLeagueAdminByID(oRs.Fields("LeagueID").Value) OR IsLeagueTeamCaptainByID(intTeamID, oRs.Fields("LeagueID").Value) Then
					intURLID = intURLID + 1
					If Not(bShownadmin) Then 
						bShownadmin = True
						Call ShowAdminPanel()
					End If
					%>
					<tr>
						<td bgcolor="<%=bgc%>"><%=Server.HTMLEncode(ors.Fields("LeagueName").Value)%> League</td>
						<td bgcolor="<%=bgc%>" align="center"><a href="TeamLeagueAdmin.asp?league=<%=server.urlencode(ors.Fields("LeagueName").Value)%>&team=<%=server.urlencode(strTeamName)%>">Admin Panel</td>
						<td bgcolor="<%=bgc%>">&nbsp;</td>
					</tr>
					<%
					If bgc = bgcone Then
						bgc = bgctwo
					Else
						bgc = bgcone
					End If
				End If
				oRS.MoveNext
			Loop
		End If
		oRS.NextRecordset 
		
		' start ladder lookup
		strSQL = "SELECT l.LadderName, lnk.TLLinkID, l.LadderID "
		strSQL = strSQL & " FROM tbl_ladders l, lnk_t_l lnk "
		strSQL = strSQL & " WHERE l.LadderID = lnk.LadderID "
		strSQL = strSQL & " AND lnk.TeamID = '" & intTeamID & "'" 
		strSQL = strSQL & " AND lnk.IsActive = 1 "
		strSQL = strSQL & " AND l.LadderActive = 1 "
		strSQL = strSQL & " ORDER BY l.LadderName "
		oRS.Open strSQL, oConn
		bgc = bgcone
		if not (ors.EOF and ors.BOF) then
			intNumDone = 0
			Do While Not oRS.EOF
				bTeamCaptain=IsTeamCaptainByID(intTeamID, ors.fields("LadderID").value)
				bLadderAdmin=IsLadderAdminByID(ors.fields("LadderID").value)
				if bTeamCaptain or bTeamFounder or bLadderAdmin or bSysAdmin then
					intURLID = intURLID + 1
					If Not(bShownadmin) Then 
						bShownadmin = True
						Call ShowAdminPanel()
					End If
					%>
					<script>
						quitladderurl<%=intURLID%> = "quitLadder.asp?team=<%=server.URLEncode(strTeamName)%>&ladder=<%=server.urlencode(ors.Fields("LadderName").Value)%>&url="+this.location.href;
					</script>
					<tr>
						<td bgcolor="<%=bgc%>"><%=Server.HTMLEncode(ors.Fields("LadderName").Value)%> Ladder</td>
						<td align="center" bgcolor="<%=bgc%>"><a href="TeamLadderAdmin.asp?ladder=<%=server.urlencode(ors.Fields("LadderName").Value)%>&team=<%=server.urlencode(strTeamName)%>">Admin Panel</a></td>
						<td bgcolor="<%=bgc%>"><a href="javascript:popup(quitladderurl<%=intURLID%>, 'quitladder', 150, 300, 'no');">Remove Team from Ladder</a>
					</tr>
					</td>
					<%
					If bgc = bgcone Then
						bgc = bgctwo
					Else
						bgc = bgcone
					End If
				End If
				oRS.MoveNext
			Loop
		End If
		oRS.NextRecordset 
		

		' start scrim aldder lookup
		strSQL = "SELECT l.EloLadderName, lnk.lnkEloTeamID, l.EloLadderID "
		strSQL = strSQL & " FROM tbl_elo_ladders l, lnk_elo_team lnk "
		strSQL = strSQL & " WHERE l.EloLadderID = lnk.EloLadderID "
		strSQL = strSQL & " AND lnk.TeamID = '" & intTeamID & "'" 
		strSQL = strSQL & " AND lnk.Active = 1 "
		strSQL = strSQL & " AND l.EloActive = 1 "
		strSQL = strSQL & " ORDER BY l.EloLadderName "
		oRS.Open strSQL, oConn
		bgc = bgcone
		if not (ors.EOF and ors.BOF) then
			intNumDone = 0
			Do While Not oRS.EOF
				bTeamCaptain=IsEloTeamCaptainByID(intTeamID, ors.fields("EloLadderID").value)
				bLadderAdmin=IsEloLadderAdminByID(ors.fields("EloLadderID").value)
				if bTeamCaptain or bTeamFounder or bLadderAdmin or bSysAdmin then
					intURLID = intURLID + 1
					If Not(bShownadmin) Then 
						bShownadmin = True
						Call ShowAdminPanel()
					End If
					%>
					<script>
						quitladderurl<%=intURLID%> = "/scrim/quitLadder.asp?team=<%=server.URLEncode(strTeamName)%>&ladder=<%=server.urlencode(ors.Fields("EloLadderName").Value)%>&url="+this.location.href;
					</script>
					<tr>
						<td bgcolor="<%=bgc%>"><%=Server.HTMLEncode(ors.Fields("EloLadderName").Value)%></td>
						<td align="center" bgcolor="<%=bgc%>"><a href="TeamScrimLadderAdmin.asp?ladder=<%=server.urlencode(ors.Fields("EloLadderName").Value)%>&team=<%=server.urlencode(strTeamName)%>">Admin Panel</a></td>
						<td bgcolor="<%=bgc%>"><a href="javascript:popup(quitladderurl<%=intURLID%>, 'quitladder', 150, 300, 'no');">Remove Team from Ladder</a></td>
					</td>
					<%
					If bgc = bgcone Then
						bgc = bgctwo
					Else
						bgc = bgcone
					End If
				End If
				oRS.MoveNext
			Loop
		End If
		oRS.NextRecordset 
		
		'start tournament look up code
		strsql = "select tbl_tournaments.tournamentname, lnk_t_m.tmlinkid, tbl_tournaments.tournamentid from lnk_t_m " &_
					"inner join tbl_tournaments on lnk_t_m.tournamentid = tbl_tournaments.tournamentid " &_
					"where teamid='" & intTeamID & "' AND tbl_tournaments.Active = 1 AND lnk_t_m.Active = 1 "
		ors.Open strsql, oconn
		bOnATournament = false
		if not (ors.EOF and ors.BOF) then
			bOnATournament = True
			do while not ors.EOF
				bTeamCaptain = IsTournamentTeamCaptainByID(intTeamID, ors.fields("TournamentID").value)
				bTournamentAdmin = IsTournamentAdminByID(ors.fields("TournamentID").value)
				if bTeamCaptain or bTeamFounder or bTournamentAdmin or bSysAdmin then
					intURLID = intURLID + 1
					If Not(bShownadmin) Then 
						bShownadmin = True
						Call ShowAdminPanel()
					End If
					%>
					<tr>
						<td bgcolor="<%=bgc%>"><%=Server.HTMLEncode(ors.Fields("TournamentName").Value)%> Tournament</td>
						<td align="center" bgcolor="<%=bgc%>"><a href="TeamTournamentAdmin.asp?tournament=<%=server.urlencode(ors.Fields("TournamentName").Value)%>&team=<%=server.urlencode(strTeamName)%>">Admin Panel</a></td>
						<td bgcolor="<%=bgc%>">&nbsp;</td>
					</tr>
					<%
					if bgc = bgctwo then
						bgc = bgcone
					else
						bgc=bgctwo
					end if
				end if
				ors.movenext
			loop			
		end if
		oRS.NextRecordset 
		if bShownAdmin then
			Response.Write "</form></table>"
			Call ContentEnd()
		End If
	end if
	
	'----------------------------
	' Active Roster Information
	'----------------------------
	' Call Content2BoxStart("Active Roster and Competition Information")
	Dim blnFirst
	blnFirst = False
	'----------------------------
	' Leagues
	'----------------------------
	Dim strLeagueName, intLeagueID, intConferenceID, strConferenceName, intDivisionID, strDivisionName
	Dim intLeagueTeamID, intLeagueWins, intLeagueLosses
	Dim intLeagueDraws, intLeaguePoints, intLeagueWinPct, intLeagueRank, intLeagueRoundsOne, intLeagueRoundsLost
	strSQL = "SELECT lnk.lnkLeagueTeamID, LeagueName, RosterLock, c.LeagueID, c.LeagueConferenceID, ConferenceName, lnk.LeagueDivisionID, "
	strSQL = strSQL & " Wins, Losses, Draws, LeaguePoints, WinPct, Rank, RoundsWon, RoundsLost "
	strSQL = strSQL & " FROM lnk_league_team lnk "
	strSQL = strSQL & " INNER JOIN tbl_leagues l "
	strSQL = strSQL & " ON l.LeagueID = lnk.LeagueID "
	strSQL = strSQL & " INNER JOIN tbl_league_conferences c "
	strSQL = strSQL & " ON c.LeagueConferenceID = lnk.LeagueConferenceID "
	strSQL = strSQL & " WHERE lnk.TeamID = '" & intTeamID & "' AND Active = 1 AND LeagueActive = 1 "
	strSQL = strSQL & " ORDER BY LeagueName ASC"
	'Response.Write strSQL
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		Do While Not(oRs.EOF)
			If Not(blnFirst) Then
				Call Content2BoxStart("")
			End If
			blnFirst = False
			
			strLeagueName = oRs.FieldS("LeagueName").Value
			intLeagueID = oRs.FieldS("LeagueID").Value
			intConferenceID = oRs.FieldS("LeagueConferenceID").Value
			strConferenceName = oRs.FieldS("ConferenceName").Value
			intDivisionID = oRs.FieldS("LeagueDivisionID").Value
			intLeagueTeamID = oRs.FieldS("lnkLeagueTeamID").Value
			intLeagueWins = oRs.FieldS("Wins").Value
			intLeagueLosses = oRs.FieldS("Losses").Value
			intLeagueDraws = oRs.FieldS("Draws").Value
			intLeaguePoints = oRs.FieldS("LeaguePoints").Value
			intLeagueWinPct = oRs.FieldS("WinPct").Value
			intLeagueRank = oRs.FieldS("Rank").Value
			intLeagueRoundsOne = oRs.FieldS("RoundsWon").Value
			intLeagueRoundsLost = oRs.FieldS("RoundsLost").Value
			%>
				<table border=0 cellspacing=0 cellpadding=0 bgcolor="#444444" width=97% class="cssBordered" align="center">
			<%
			If Cint(intDivisionID) = 0 Then
				' Pending joining... show that...
				%>
				<tr>
					<td bgcolor="<%=bgcone%>" colspan="2"><a href="viewleague.asp?league=<%=Server.URLEncode(strLeagueName)%>"><%=strLeagueName%> League</a> &raquo; <a href="viewleagueconference.asp?league=<%=Server.URLEncode(strLeagueName)%>&conference=<%=Server.URLEncode(strConferenceName)%>"><%=strConferenceName%> Conference</a></td>
				</tr>
				<tr>
					<td bgcolor="#000000" colspan="2"><i>Pending assignment to a division.</i></td>
				</tr>
				<%
			Else
				strSQL = "SELECT DivisionName FROM tbl_league_divisions WHERE LeagueDivisionID = '" & intDivisionID & "'"
				oRs2.Open strSQL, oConn
				If Not(oRs2.EOF AND oRs2.BOF) Then
					strDivisionName = oRs2.Fields("DivisionName").Value
				End If
				oRs2.NextRecordSet
				%>
				<tr>
					<td bgcolor="<%=bgcone%>" colspan="3">
						<a href="viewleague.asp?league=<%=Server.URLEncode(strLeagueName)%>"><%=strLeagueName%> League</a> 
						&raquo; 
						<a href="viewleagueconference.asp?league=<%=Server.URLEncode(strLeagueName)%>&conference=<%=Server.URLEncode(strConferenceName)%>"><%=strConferenceName%> Conference</a> 
						&raquo; 
						<a href="viewleaguedivision.asp?league=<%=Server.URLEncode(strLeagueName)%>&conference=<%=Server.URLEncode(strConferenceName)%>&division=<%=Server.URLEncode(strDivisionName)%>"><%=Server.HTMLEncode(strDivisionName)%> Division
					</td>
				</tr>
				<tr>
					<td colspan="3" bgcolor="#000000"><b>Pending Matches:</b></td>
				</tr>
				<%
				strSQL = "EXECUTE LeagueTeamMatches @LeagueTeamID = '" & intLeagueTeamID & "'"
				oRs2.Open strSQL, oConn
				If (oRs2.State = 1) Then
					If Not(oRs2.EOF AND oRs2.BOF) Then
						%>
						<tr>
							<th bgcolor="#000000" width="65%">Opponent</td>
							<th bgcolor="#000000">Date</td>
							<th bgcolor="#000000">&nbsp;</td>
						</tr>
						<%
						Do While Not (oRs2.EOF)
							%>
							<tr>
								<td bgcolor="<%=bgctwo%>"><b><%
									If oRs2.Fields("LeagueDivisionID").Value <> 0 Then
										Response.Write "Division"
									ElseIf oRs2.Fields("LeagueConferenceID").Value <> 0 Then
										Response.Write "Conference"
									Else
										Response.Write "League"
									End If
									%></b> &raquo; <a href="viewteam.asp?team=<%=Server.URLEncode(oRS2.Fields("OpponentName").Value & "")%>"><%=Server.HTMLEncode(oRs2.Fields("OpponentName").Value & "")%></a></td>
								<td align="center" bgcolor="<%=bgctwo%>"><%=FormatDateTime(oRs2.FIelds("MatchDate").Value, 2)%></td>
								<td align="center" bgcolor="<%=bgctwo%>" nowrap="nowrap"><a href="viewLeagueMatch.asp?League=<%=Server.URLEncode(strLeagueName & "")%>&LeagueMatchID=<%=oRs2.Fields("LeagueMatchID").Value%>">Rants (<%=oRs2.Fields("Rants").Value%>)</a></td>
							</tr>								
							<tr>
								<td colspan="3" bgcolor="<%=bgcone%>" align="center">
									Map(s): <%
									If Len(oRs2.Fields("Map1").Value) > 0 Then
										Response.Write oRs2.Fields("Map1").Value
									End If
									If Len(oRs2.Fields("Map2").Value) > 0 Then
										Response.Write ", " & oRs2.Fields("Map2").Value
									End If
									If Len(oRs2.Fields("Map3").Value) > 0 Then
										Response.Write ", " & oRs2.Fields("Map3").Value
									End If
									If Len(oRs2.Fields("Map4").Value) > 0 Then
										Response.Write ", " & oRs2.Fields("Map4").Value
									End If
									If Len(oRs2.Fields("Map5").Value) > 0 Then
										Response.Write ", " & oRs2.Fields("Map5").Value
									End If
									%></td>
							</tr>
							<tr>
								<td colspan="3" bgcolor="#000000"><img src="images/spacer.gif" height="5" border="0" width="0" /></td>
							</tr>
							<%

							oRs2.MoveNext
						Loop
					Else
						%>
						<tr>
							<td colspan="3" bgcolor="<%=bgctwo%>">No pending matches scheduled.</td>
						</tr>
						<%
					End If
				End If
				oRS2.NextRecordSet
			End If
			%>
			</table>
			<% Call Content2BoxMiddle() %>
				<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 class="cssBordered" WIDTH=97% align=center>
						<tr>
							<th bgcolor="#000000" colspan="3">
								<a href="viewLeagueRoster.asp?Team=<%=Server.URLEncode(strTeamName)%>&League=<%=Server.URLEncode(strLeagueName)%>">View Detailed Roster</a></th>
						</tr>								
				<TR BGCOLOR="#000000">
					<TH WIDTH=130>Player</TH>
					<TH WIDTH=120>Status</TH>
					<TH WIDTH=120>Join Date</TH>
				</TR>
			<%
			strSQL="select PlayerHandle, Suspension, lnk_league_team_player.JoinDate, lnk_league_team_player.IsAdmin "
			strSQL = strSQL & " from tbl_Players inner join lnk_league_team_player on "
			strSQL = strSQL & "lnk_league_team_player.PlayerID=tbl_players.playerid where "
			strSQL = strSQL & " lnk_league_team_player.lnkLeagueTeamID=" & intLeagueTeamID & " ORDER BY PlayerHandle"
			ors2.Open strSQL, oconn
			bgc = bgcone
			bOnTeam = False
			intRosterCount = 0
			if not (ors2.eof and ors2.BOF) then
				do while not ors2.EOF
					intRosterCount = intRosterCount + 1
					if len(ors2.Fields("JoinDate").Value) < 8 then
						strDateJoined="-"
					else
						strDateJoined = formatdatetime(ors2.Fields("JoinDate").Value,2)
					end if
					if ors2.Fields("PlayerHandle").Value = Session("uName") then
						bOnTeam = True
					end if
					if ors2.Fields("IsAdmin").Value=1 then
						strAdmin = "Team Captain"
					else
						strAdmin = "&nbsp;"						
					end if
					if Trim(ors2.Fields("PlayerHandle").Value) = Trim(strFounderName) then
						strAdmin = "Team Founder"
					end if
					If (ors2.Fields("Suspension").Value = 1) Then
						strAdmin = "<b><font color=""#ff0000"">SUSPENDED</font></b>"
					End If
					Response.Write "<tr height=18 bgcolor=" & bgc & ">"
					Response.Write "<td><a href=viewplayer.asp?Player=" & server.urlencode(ors2.Fields("PlayerHandle").Value) & ">" & Server.HTMLEncode(ors2.Fields("PlayerHandle").Value) & "</a></td>"
					Response.Write "<td ALIGN=CENTER>" & strAdmin & "</td>"
					Response.Write "<td align=right>" & strDateJoined & "</td></tr>" & vbCrLf
					oRS2.MoveNext 
					if bgc=bgcone then
						bgc=bgctwo
					else
						bgc=bgcone
					end if
				loop
			end if
			if Session("uName") = "" or Session("uName") = strFounderName then
				Response.Write " "
			elseif bOnTeam Then
				%>
				<form name="frmQuitTeam<%=intURLID%>" id="frmQuitTeam<%=intURLID%>">
				<tr BGCOLOR="#000000"><td align=center colspan=3>
				<script>
				quiturl<%=intURLID%> = "quitTeamOnLeague.asp?teamID=" + <%=intTeamID%> + "&leagueid=" + <%=intLeagueID%> + "&type=quit&url="+this.location.href;
				</script>
				<input type="button" value="Quit Team" class="bright" onclick="javascript:popup(quiturl<%=intURLID%>, 'quit', 150, 300, 'no')" style='width:150'>
				</td></tr>
				</form>
			<%
			elseIf cStr(oRs.fields("RosterLock").Value & "") = "1" Then
				%>
				<tr BGCOLOR="#000000"><td align=center colspan=3><font color="#ff0000"><b>Rosters are locked for this league</b></font></td></tr>
				<%
			Else
			%>
				<form name="frmJoinTeam<%=intURLID%>" id="frmJoinTeam<%=intURLID%>">
				<tr BGCOLOR="#000000"><td align=center colspan=3>
				<script>
					joinurlleague<%=intURLID%> = "joinTeamOnLeague.asp?teamID=" + <%=intTeamID%> + "&leagueid=" + <%=intLeagueID%> + "&type=join&url="+this.location.href;
				</script>
				<input type="button" value="Join" class="bright" onclick="javascript:popup(joinurlleague<%=intURLID%>, 'jointeam', 150, 300, 'no')" style='width:150'><br>
				</td></tr>
				</form>
				<%
			End If
			ors2.Close 
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			intURLID = intURLID + 1	
			%>
			</table>
			<% Call Content2BoxEnd() %>
			<%
			oRs.MoveNext
		Loop
	End If
	oRs.NextRecordSet
	'----------------------------
	' End Leagues
	'----------------------------
	
	
	'----------------------------
	' Ladders
	'----------------------------
	strSQL="select LadderName, tbl_ladders.LadderID, tbl_ladders.Maps, lnk_T_L.TLLinkID, lnk_T_L.Rank, lnk_T_L.wins, lnk_T_L.losses, lnk_T_L.isactive, lnk_t_l.status, tbl_ladders.RosterLimit from tbl_Ladders inner join lnk_T_L on tbl_ladders.ladderID=lnk_T_L.ladderID where lnk_T_L.teamID=" & intTeamID & " AND IsActive = 1 AND LadderActive=1 order by laddername"
	ors.Open strSQL, oconn
	bOnTeam = False
	bgc=bgctwo
	if not (ors.EOF and ors.BOF) then
		do while not ors.EOF
			intMaps = oRs.FIeldS("Maps").Value
			intRosterLimit = ors("RosterLimit")
			intLadderID = oRS.Fields("LadderID").Value 

			if ors.Fields("IsActive").Value = 1 then
				intMatchID = ""
				If Not(blnFirst) Then
					Call Content2BoxStart("")
				End If
				blnFirst = False
				
				%>
						<table border=0 cellspacing=0 cellpadding=0 class="cssbordered" width="97%" align=center>
						<tr BGCOLOR="#000000">
							<TH width=225>Ladder</TH>
							<TH width=75 align=center>Rung</TH>
							<TH width=75 align=center>Record</TH>
						</tr>
					<%
					strLadderName = ors.Fields("LadderName").Value 
					Response.Write "<tr BGCOLOR=" & bgcone & "><td align=left>&nbsp;<a href=viewladder.asp?ladder=" & server.urlencode(strLadderName) & ">" & Server.HTMLEncode(strLadderName) & "</a></td>"
					Response.Write "<td valign=top align=center>" & ors.Fields("Rank").Value & "</td>"
					Response.Write "<td valign=top align=center>" & ors.Fields("Wins").Value & "/" & ors.Fields("Losses").Value & "</td></tr>"
					strStatus = ors.Fields("Status").Value
					Select Case Left(uCase(strStatus), 6)
						Case "DEFEND", "ATTACK"
							If  Left(uCase(strStatus), 6)  = "DEFEND" Then
								strsql = "select m.MatchAttackerID, m.MatchID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID, MatchMap4ID, MatchMap5ID, t.teamname, t.teamtag "
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
								end if
								Response.Write "<tr BGCOLOR=""#000000""><td COLSPAN=3><b>Defending:</b>&nbsp;<a href=viewteam.asp?team=" & server.urlencode(strEnemyName) & ">" & Server.HTMLEncode(strEnemyName)&"</a>"
								Response.Write "<br><br><b>Maps:</b><br>"
								For i = 1 To oRs.Fields("Maps").Value
									If i > 1 Then
										Response.Write ", "
									End If
									Response.Write Server.HTMLEncode(strMapArray(i))
								Next
								response.write "<br><br><b>Match Date:</b><br>" & strMatchDate
								ors2.NextRecordset 

								If Right(uCase(strTeamName), 1) = "s" Then
									strVerbiage = " have "
								Else
									strVerbiage = " has "
								End IF
								
								If Right(uCase(strEnemyName), 1) = "s" Then
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
									<BR><BR><B>Match Options:</B><BR>
									<%
									CurrentMap = ""
									CurrentMapNumber = -1
									Do While Not(oRS2.EOF)
									'Response.Write oRS2.Fields("MapNumber").Value
										If cInt(oRS2.Fields("MapNumber").Value) <= cInt(intMaps) Then
											If CurrentMapNumber <> oRS2.Fields("MapNumber").Value Then
												CurrentMapNumber = oRS2.Fields("MapNumber").Value
												If CurrentMapNumber <> 0 Then
													CurrentMap = strMapArray(CurrentMapNumber)
													Response.Write "&nbsp;&nbsp;<B>" & CurrentMap & "</B><BR>"
												End If
											End If
											Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
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
										End If
										oRS2.MoveNext
									Loop
								End If
								oRS2.NextRecordset 
								Response.write "</TD></TR>"
							Else
								strsql = "select m.MatchDefenderID, m.MatchID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID, MatchMap4ID, MatchMap5ID, t.teamname, t.teamtag "
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
								end if
								Response.Write "<tr BGCOLOR=""#000000""><td COLSPAN=3><b>Attacking:</b>&nbsp;<a href=viewteam.asp?team=" & server.urlencode(strEnemyName) & ">" & Server.HTMLEncode(strEnemyName)&"</a>"
								Response.Write "<br><br><b>Maps:</b><br>"
								For i = 1 To oRs.Fields("Maps").Value
									If i > 1 Then
										Response.Write ", "
									End If
									Response.Write Server.HTMLEncode(strMapArray(i))
								Next
								response.write "<br><br><b>Match Date:</b><br>" & strMatchDate
								ors2.NextRecordset 
								

								If Right(uCase(strTeamName), 1) = "s" Then
									strVerbiage = " have "
								Else
									strVerbiage = " has "
								End IF
								
								If Right(uCase(strEnemyName), 1) = "s" Then
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
									<BR><BR><B>Match Options:</B><BR>
									<%
									CurrentMap = ""
									CurrentMapNumber = -1
									Do While Not(oRS2.EOF)
										If cInt(oRS2.Fields("MapNumber").Value) <= cInt(intMaps) Then
											If CurrentMapNumber <> oRS2.Fields("MapNumber").Value Then
												CurrentMapNumber = oRS2.Fields("MapNumber").Value
												If CurrentMapNumber <> 0 Then
													CurrentMap = strMapArray(CurrentMapNumber)
													Response.Write "&nbsp;&nbsp;<B>" & CurrentMap & "</B><BR>"
												End If
											End If
											Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;"
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
										End If
										oRS2.MoveNext
									Loop
								End If
								oRS2.NextRecordset 
							End If
							Response.write "</TD></TR>"
						Case "IMMUNE", "DEFEAT", "RESTIN"
							Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=3>" & strStatus & "</TD></TR>"
						Case Else
							Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=3>Available to challenge or be challenged</TD></TR>"
					End Select
					%>
					</table>
					<% Call Content2BoxMiddle() %>
						<table border=0 cellspacing=0 cellpadding=0 class="cssbordered" width="97%" align=center>
						<tr>
							<th bgcolor="#000000" colspan="3">
								<a href="viewLadderRoster.asp?Team=<%=Server.URLEncode(strTeamName)%>&Ladder=<%=Server.URLEncode(strLadderName)%>">View Detailed Roster</a></th>
						</tr>								
						<TR BGCOLOR="#000000">
							<TH WIDTH=130>Player</TH>
							<TH WIDTH=120>Status</TH>
							<TH WIDTH=120>Join Date</TH>
						</TR>
						<%
						strSQL="select PlayerHandle, Suspension, lnk_T_P_L.DateJoined, lnk_T_P_L.IsAdmin "
						strSQL = strSQL & " from tbl_Players inner join lnk_T_P_L on "
						strSQL = strSQL & "lnk_T_P_L.PlayerID=tbl_players.playerid where "
						strSQL = strSQL & " lnk_T_P_L.TLLinkID=" & ors.Fields("TLLinkID").Value & " order by PlayerHandle"
						ors2.Open strSQL, oconn
						bgc = bgcone
						bOnTeam = False
						intRosterCount = 0
						if not (ors2.eof and ors2.BOF) then
							do while not ors2.EOF
								intRosterCount = intRosterCount + 1
								if len(ors2.Fields("DateJoined").Value) < 8 then
									strDateJoined="-"
								else
									strDateJoined = formatdatetime(ors2.Fields("DateJoined").Value,2)
								end if
								if ors2.Fields("PlayerHandle").Value = Session("uName") then
									bOnTeam = True
								end if
								if ors2.Fields("IsAdmin").Value=1 then
									strAdmin = "Team Captain"
								else
									strAdmin = "&nbsp;"						
								end if
								if Trim(ors2.Fields("PlayerHandle").Value) = Trim(strFounderName) then
									strAdmin = "Team Founder"
								end if
								If (ors2.Fields("Suspension").Value = 1) Then
									strAdmin = "<b><font color=""#ff0000"">SUSPENDED</font></b>"
								End If
								Response.Write "<tr height=18 bgcolor=" & bgc & ">"
								Response.Write "<td><a href=viewplayer.asp?Player=" & server.urlencode(ors2.Fields("PlayerHandle").Value) & ">" & Server.HTMLEncode(ors2.Fields("PlayerHandle").Value) & "</a></td>"
								Response.Write "<td ALIGN=CENTER>" & strAdmin & "</td>"
								Response.Write "<td align=right>" & strDateJoined & "</td></tr>" & vbCrLf
								oRS2.MoveNext 
								if bgc=bgcone then
									bgc=bgctwo
								else
									bgc=bgcone
								end if
							loop
						end if
						if Session("uName") = "" or Session("uName") = strFounderName then
							Response.Write " "
						elseif bOnTeam Then
							%>
							<form name="frmQuitTeam<%=intURLID%>" id="frmQuitTeam<%=intURLID%>">
							<tr BGCOLOR="#000000"><td align=center colspan=3>
							<script>
							quiturl<%=intURLID%> = "quitTeamOnLadder.asp?teamID=" + <%=intTeamID%> + "&ladderid=" + <%=intLadderID%> + "&type=quit&url="+this.location.href;
							</script>
							<input type="button" value="Quit Team" class="bright" onclick="javascript:popup(quiturl<%=intURLID%>, 'quit', 150, 300, 'no')" style='width:150'>
							</td></tr>
							</form>
						<%
						elseIf intRosterCount < intRosterLimit OR (intRosterLimit = 0) Then
						%>
							<form name="frmJoinTeam<%=intURLID%>" id="frmJoinTeam<%=intURLID%>">
							<tr BGCOLOR="#000000"><td align=center colspan=3>
							<script>
								joinurlladder<%=intURLID%> = "joinTeamOnLadder.asp?teamID=" + <%=intTeamID%> + "&ladderid=" + <%=intLadderID%> + "&type=join&url="+this.location.href;
							</script>
							<input type="button" value="Join" class="bright" onclick="javascript:popup(joinurlladder<%=intURLID%>, 'jointeam', 150, 300, 'no')" style='width:150'><br>
							</td></tr>
							</form>
							<%
						Else
							Response.write "<TR BGCOLOR=""#000000""><TD ALIGN=CENTER COLSPAN=3><B><FONT color=red>Roster limit reached for this ladder.</FONT></B></TD></TR>"
						End If
						ors2.Close 
						if bgc=bgcone then
							bgc=bgctwo
						else
							bgc=bgcone
						end if
						intURLID = intURLID + 1	
					end if
					%>
					</table>
				<% Call Content2BoxEnd() %>
			<%					
		ors.MoveNext
		loop
	END IF
	ors.NextRecordset
	'------------------
	' End Ladders
	'------------------
	
	'----------------------------
	' Scrim Ladders
	'----------------------------
	strSQL = "SELECT EloLadderName, l.EloLadderID, lnk.lnkEloTeamID, lnk.Rating, lnk.WIns, lnk.Losses "
	strSQL = strSQL & " FROM lnk_elo_team lnk INNER JOIN tbl_elo_ladders l ON l.EloLadderID = lnk.EloLadderID "
	strSQL = strSQL & " WHERE lnk.TeamID = '" & intTeamID & "' AND Active = 1 AND EloActive = 1 ORDER BY EloLadderName ASC "
	ors.Open strSQL, oconn
	bOnTeam = False
	bgc=bgctwo
	if not (ors.EOF and ors.BOF) then
		do while not ors.EOF
			If Not(blnFirst) Then
				Call Content2BoxStart("")
			End If
			blnFirst = False
			
			intLadderID = oRS.Fields("EloLadderID").Value 
			%>
					<table border=0 cellspacing=0 cellpadding=0 class="cssbordered" width="97%" align=center>
					<tr BGCOLOR="#000000">
						<TH width=225>Ladder</TH>
						<TH width=75 align=center>Rating</TH>
						<TH width=75 align=center>Record</TH>
					</tr>
			<%
			strLadderName = ors.Fields("EloLadderName").Value 
			Response.Write "<tr BGCOLOR=" & bgcone & "><td align=left>&nbsp;<a href=viewscrimladder.asp?ladder=" & server.urlencode(strLadderName) & ">" & Server.HTMLEncode(strLadderName) & "</a></td>"
			Response.Write "<td valign=top align=center>" & ors.Fields("Rating").Value & "</td>"
			Response.Write "<td valign=top align=center>" & ors.Fields("Wins").Value & "/" & ors.Fields("Losses").Value & "</td></tr>"
			%>	
					<tr>
						<td colspan="3" bgcolor="#000000">
							<%
							strSQL = "SELECT TeamName, et.TeamID, Mode = 'A' FROM tbl_elo_matches em "
							strSQL = strSQL & "	INNER JOIN lnk_elo_team et ON em.DefenderEloTeamID = et.lnkEloTEamID "
							strSQL = strSQL & "	INNER JOIN tbl_teams t ON t.TeamID = et.TeamID "
							strSQL = strSQL & "	WHERE em.AttackerEloTeamID = '" & oRs.FIelds("lnkEloTeamID").Value & "' AND MatchActive = 1"
							strSQL = strSQL & "UNION ALL "
							strSQL = strSQL & "SELECT TeamName, et.TeamID, 'D' FROM tbl_elo_matches em "
							strSQL = strSQL & "	INNER JOIN lnk_elo_team et ON em.AttackerEloTeamID = et.lnkEloTEamID "
							strSQL = strSQL & "	INNER JOIN tbl_teams t ON t.TeamID = et.TeamID "
							strSQL = strSQL & "	WHERE em.DefenderEloTeamID =  '" & oRs.FIelds("lnkEloTeamID").Value & "' AND MatchActive = 1"
 							oRs2.Open strSQL, oConn
 							'Response.Write strSQL
 							If Not(oRs2.EOF AND oRs2.BOF) Then
 								%>
 								<br /><b>Pending Matches:</b><br />
 								<%
 								Do While Not(oRs2.EOF)
	 								If oRs2.Fields("Mode").Value = "A" Then
	 									%>
	 									&nbsp;Attacking <a href="viewteam.asp?team=<%=Server.URLEncode(oRs2.Fields("TeamName").Value & "")%>"><%=Server.HTMLEncode(oRs2.Fields("TeamName").Value & "")%></a><br />
	 									<%
	 								Else
										%>
	 									&nbsp;Defending against <a href="viewteam.asp?team=<%=Server.URLEncode(oRs2.Fields("TeamName").Value & "")%>"><%=Server.HTMLEncode(oRs2.Fields("TeamName").Value & "")%></a><br />
	 									<%
	 	 							End If
	 	 							oRs2.MoveNext
	 	 						Loop
 							Else
 								%>
 								<b>No pending matches</b>
 								<%
 							End If	
 							oRs2.nextRecordSet
 							%><br />
						</td>
					</tr>
					</TABLE>
			<% Call Content2BoxMiddle() %>
				<table border=0 cellspacing=0 cellpadding=0 class="cssbordered" width="97%" align=center>
			<TR BGCOLOR="#000000">
				<TH WIDTH=130>Player</TH>
				<TH WIDTH=120>Status</TH>
				<TH WIDTH=120>Join Date</TH>
			</TR>
			<%
				strSQL="select PlayerHandle, Suspension, lnk.JoinDate, lnk.IsAdmin "
				strSQL = strSQL & " from tbl_Players p INNER JOIN lnk_elo_team_player lnk ON lnk.PlayerID = p.PlayerID"
				strSQL = strSQL & " WHERE lnk.lnkEloTeamID = '" & oRs.Fields("lnkEloTeamID").Value & "' ORDER BY PlayerHandle"
				ors2.Open strSQL, oconn
				bgc = bgcone
				bOnTeam = False
				intRosterCount = 0
				if not (ors2.eof and ors2.BOF) then
					do while not ors2.EOF
						intRosterCount = intRosterCount + 1
						if len(ors2.Fields("JoinDate").Value) < 8 then
							strDateJoined="-"
						else
							strDateJoined = formatdatetime(ors2.Fields("JoinDate").Value,2)
						end if
						if ors2.Fields("PlayerHandle").Value = Session("uName") then
							bOnTeam = True
						end if
						if ors2.Fields("IsAdmin").Value=1 then
							strAdmin = "Team Captain"
						else
							strAdmin = "&nbsp;"						
						end if
						if Trim(ors2.Fields("PlayerHandle").Value) = Trim(strFounderName) then
							strAdmin = "Team Founder"
						end if
						If (ors2.Fields("Suspension").Value = 1) Then
							strAdmin = "<b><font color=""#ff0000"">SUSPENDED</font></b>"
						End If
						Response.Write "<tr height=18 bgcolor=" & bgc & ">"
						Response.Write "<td><a href=viewplayer.asp?Player=" & server.urlencode(ors2.Fields("PlayerHandle").Value) & ">" & Server.HTMLEncode(ors2.Fields("PlayerHandle").Value) & "</a></td>"
						Response.Write "<td ALIGN=CENTER>" & strAdmin & "</td>"
						Response.Write "<td align=right>" & strDateJoined & "</td></tr>" & vbCrLf
						oRS2.MoveNext 
						if bgc=bgcone then
							bgc=bgctwo
						else
							bgc=bgcone
						end if
					loop
				end if
				if Session("uName") = "" or Session("uName") = strFounderName then
					Response.Write " "
				elseif bOnTeam Then
					%>
					<form name="frmQuitTeam<%=intURLID%>" id="frmQuitTeam<%=intURLID%>">
					<tr BGCOLOR="#000000"><td align=center colspan=3>
					<script>
					quiturl<%=intURLID%> = "quitTeamOnScrimLadder.asp?teamID=" + <%=intTeamID%> + "&ladderid=" + <%=intLadderID%> + "&type=quit&url="+this.location.href;
					</script>
					<input type="button" value="Quit Team" class="bright" onclick="javascript:popup(quiturl<%=intURLID%>, 'quit', 150, 300, 'no')" style='width:150'>
					</td></tr>
					</form>
				<%
				Else
				%>
					<form name="frmJoinTeam<%=intURLID%>" id="frmJoinTeam<%=intURLID%>">
					<tr BGCOLOR="#000000"><td align=center colspan=3>
					<script>
						joinurlladder<%=intURLID%> = "joinTeamOnScrimLadder.asp?teamID=" + <%=intTeamID%> + "&ladderid=" + <%=intLadderID%> + "&type=join&url="+this.location.href;
					</script>
					<input type="button" value="Join" class="bright" onclick="javascript:popup(joinurlladder<%=intURLID%>, 'jointeam', 150, 300, 'no')" style='width:150'><br>
					</td></tr>
					</form>
					<%
				End If
				ors2.Close 
				if bgc=bgcone then
					bgc=bgctwo
				else
					bgc=bgcone
				end if
				intURLID = intURLID + 1	
			%>
			</table>
			<% Call Content2BoxEnd() %>
			<%					
		ors.MoveNext
		loop
	END IF
	ors.NextRecordset
	'--------------------
	' End Scrim Ladders
	'--------------------
	
	'--------------------
	' Start Tournaments
	'--------------------
	strSQL = "select TournamentName, tbl_tournaments.TournamentID, lnk_T_M.TMLinkID, tbl_tdivisions.DivisionID, DivisionName, tbl_tournaments.RosterLock from tbl_tdivisions, tbl_tournaments "
	strSQL = strSQL & "inner join lnk_T_M on tbl_tournaments.TournamentID=lnk_T_M.tournamentID where "
	strSQL = strSQL & " tbl_tdivisions.tournamentID = tbl_tournaments.TournamentID AND lnk_t_M.divisionid = tbl_tdivisions.divisionid AND lnk_T_M.teamID=" & intTeamID
	strSQL = strSQL & " AND tbl_tournaments.Active = 1  AND lnk_t_m.Active = 1 order by TournamentName"
	ors.Open strSQL, oconn
	bonteam=false
	bgc=bgctwo
	if not (ors.EOF and ors.BOF) then
	intURLID = 0
	do while not ors.EOF
			If Not(blnFirst) Then
				Call Content2BoxStart("")
			End If
			blnFirst = False
			%>
			<table border=0 cellspacing=0 cellpadding=0 class="cssbordered" width="97%" align=center>
		<tr BGCOLOR="#000000">
			<TH>Tournament</TH>
		</tr>
	<%		Tournament=ors.Fields(0).Value 
			Response.Write "<tr BGCOLOR=" & bgcone & "><td align=left>&nbsp;<a href=/tournament/default.asp?page=brackets&tournament=" & server.urlencode(ors.Fields(0).Value) & "&div=" & ors.Fields("DivisionID") & ">"&Server.HTMLEncode(ors.Fields(0).Value)& " - " & Server.HTMLEncode(Ors.fields("DivisionName").value) & "</a></td>"
	'Activity Code here
			Dim ServerName, ServerIP, MatchTime
			TMLinkID = ors.Fields("TMLinkID").Value 
			strsql = "select *, Team1Name = (select TeamName from tbl_teams t, lnk_t_m lnk where TMLinkID = Team1ID AND lnk.teamid = t.teamid), " &_
						" Team2Name = (select TeamName from tbl_teams t, lnk_t_m lnk where TMLinkID = Team2ID AND lnk.teamid = t.teamid) " &_
						" from tbl_rounds where (Team1ID = '" & TMLinkID & "' or Team2ID = '" & TMLinkID & "') AND WinnerID = 0 order by Round desc"
			ors2.open strsql, oconn
			if not(ors2.eof and ors2.bof) then
				if ors2("Team1ID") = TMLinkID then
					RoundsID = ors2("RoundsID")
					Team1 = true
					Team1ID = linkID
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
					Team2ID = linkID
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
				End If 
			End If
			ors2.nextrecordset
			response.write "</table>"
			Call Content2BoxMiddle()
			%>
						<table border=0 cellspacing=0 cellpadding=0 class="cssbordered" width="97%" align=center>
						<TR BGCOLOR="#000000">
							<TH>Player</TH>
							<TH>Status</TH>
							<TH>Join Date</TH>
						</TR>
			<%
			strSQL="select PlayerHandle, lnk_T_M_P.DateJoined, lnk_T_M_P.IsAdmin "
			strSQL = strSQL & " from tbl_Players inner join lnk_T_M_P "
			strSQL = strSQL & " on lnk_T_M_P.PlayerID=tbl_players.playerid where lnk_T_M_P.TMLinkID=" & ors.Fields(2).Value & " order by PlayerHandle"
			ors2.Open strSQL, oconn
			bonteam = false
			bgc=bgcone
			if not (ors2.eof and ors2.BOF) then
				do while not ors2.EOF
					if len(ors2.Fields(1).Value) < 8 then
						strDateJoined="-"
					else
						strDateJoined= formatdatetime(ors2.Fields(1).Value,2)
					end if
					if ors2.Fields(0).Value = Session("uName") then
						bonteam=true
					end if
					if ors2.Fields(2).Value=1 then
						strAdmin="Team Captain"
					else
						strAdmin="&nbsp;"						
					end if
					if ors2.Fields(0).Value = strFounderName then
						strAdmin="Team Founder"
					end if
					Response.Write "<tr height=18 bgcolor=" & bgc & "><td><a href=viewplayer.asp?Player=" & server.urlencode(ors2.Fields(0).Value) & ">" & Server.HTMLEncode(ors2.Fields(0).Value) & "</a></td><td width=125 align=center>" & strAdmin & "</td><td width=125 align=right>"&strDateJoined & "</td></tr>" & VBCRLF
					ors2.Movenext
					if bgc=bgcone then
						bgc=bgctwo
					else
						bgc=bgcone
					end if
				loop
			end if
			if Not(Session("LoggedIn")) or Session("uName") = strFounderName then
				Response.Write " "
			elseif oRs.FIelds("RosterLock").Value = "1" Then
				Response.Write "<tr><td colspan=""3"" bgcolor=""#000000"" align=""center""><b><font color=""#ff0000"">Rosters are locked for this tournament</font></b></td></tr>"
			elseif bonteam  then
				%>
				<form name="frmQuitTeam<%=intURLID%>" id="frmQuitTeam<%=intURLID%>">
				<tr BGCOLOR="#000000"><td align=center colspan=3>
				<script>
				quiturl<%=intURLID%> = "quitTeamOnTournament.asp?teamID=" + <%=intTeamID%> + "&tournamentid=" + <%=ors.Fields(1).Value%> + "&type=quit&url="+this.location.href;
				</script>
				<input type="button" value="Quit Team" class="bright" onclick="javascript:popup(quiturl<%=intURLID%>, 'quit', 150, 300, 'no')" style='width:150' id=button2 name=button2>
				</td></tr>
				</form>
			<%
			else
			%>
				<form name="frmJoinTeam<%=intURLID%>" id="frmJoinTeam<%=intURLID%>">
				<tr BGCOLOR="#000000"><td align=center colspan=3>
				<script>
					joinurltourny<%=intURLID%> = "joinTeamOnTournament.asp?teamID=" + <%=intTeamID%> + "&tournamentid=" + <%=ors.Fields(1).Value%> + "&type=join&url="+this.location.href;
				</script>
				<input type="button" value="Join" class="bright" onclick="javascript:popup(joinurltourny<%=intURLID%>, 'jointeam', 150, 300, 'no')" style='width:150' id=button3 name=button3><br>
				</td></tr>
				</form>
				<%
			end if
			ors2.Close 
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			intURLID=intURLID+1	
			%>
				</table>
			<%
			Call Content2BoxEnd()
		ors.MoveNext
		loop
	end if
	ors.Close 
	
	If False Then
		'Display History
		Call ContentStart("Recent History")
		%>
		<table width=760 border="0" cellspacing="0" cellpadding="0" class="cssBordered">
		<tr BGCOLOR="#000000">
			<TH>Ladder</TH>
			<TH width=300>Opponent</TH>
			<TH width=75>Result</TH>
			<TH WIDTH=75>Date</TH>
		</tr>
		<%
		bgc=bgctwo
		strsql="select TLLinkID from lnk_T_L where teamid=" & intTeamID & " and isactive=1"
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			do while not ors.eof
			TLLinkID=ors.Fields(0).Value
			strSQL="select TOP 2 * from vHistory where (matchwinnerid=" & TLLinkID & " or matchloserid=" & TLLinkID & ") and matchforfeit=0 order by matchdate desc"
			ors2.Open strSQL, oconn
			if not (ors2.eof and ors2.BOF) then
				do while not ors2.EOF
					If ors2.Fields("MatchWinnerID") = TLLinkID Then
						strEnemyName = oRS2.Fields("LoserName").Value 
						strResult = "Win"
					Else
						strEnemyName = oRS2.Fields("WinnerName").Value 
						strResult = "Loss"
					End If
					%>
					<tr bgcolor=<%=bgc%>><td>&nbsp;<a href=viewladder.asp?ladder=<%=server.urlencode(oRS2.Fields("LadderName").Value )%>><%=Server.HTMLEncode(oRS2.Fields("LadderName").Value)%></a></td>
					<td><a href=viewteam.asp?team=<%=server.urlencode(strEnemyName)%>><%=Server.HTMLEncode(strEnemyName)%></a></td>
					<td align="center"><%=strResult%></td>
					<td align=right><%=ors2.Fields("MatchDate").Value%>&nbsp;</td></tr>
					<%
					if bgc=bgcone then
						bgc=bgctwo
					else
						bgc=bgcone
					end if
					ors2.MoveNext
				loop
			end if
			ors2.NextRecordSet
			ors.movenext
			loop
		end if
		ors.Close 
		%>
		<tr bgcolor=<%=bgc%>><td colspan=4 align=center height=20><a href=history.asp?Keydata=<% = server.URLEncode(strTeamName)%>>Complete History</a></td></tr>
		</table>
		<%
		Call ContentEnd()
	End If
	If Not(blnFirst) Then
	'	Call Content2BoxEnd()
	End If
end if
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>