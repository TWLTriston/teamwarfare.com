<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("player"), """", "&quot;") 

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLadderAdmin, bLoggedIn
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
bLoggedIn = Session("LoggedIn")

Dim strPlayerName, intPlayerID
strPlayerName = Request("Player")

Dim strTeamName
Dim TMLinkID, DivID, TournamentName
Dim bBarDone
Dim bCanAdmin
Dim strLadderName, intRank, intLosses, intPlayerLadderID
Dim intForfeits, intWins, strStatus, strEnemyName, strResult
Dim linkID, map, opponent, mDate, statusVerbage, PlayerStatus
bBarDone = False
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<script language="javascript" type="text/javascript">
<!--
	function fConfirmRemoveIdentifier(intIdentifierID) {
		if (confirm("Are you certain you want to remove this item?")) {
			location.href = "SaveItem.asp?SaveType=AntiSmurfDel&Player=<%=Server.URLEncode(strPlayerName)%>&Identifier=" + intIdentifierID;
		}
	}
//-->
</script>

<%
If Request.QueryString("pop") = 1 Then
%>
	<SCRIPT LANGUAGE="JavaScript">
		<!--
			popup("/login.asp?url=' + this.location.href + '", "login", 175, 300, "no");
		//-->
	</SCRIPT>
<% 
End If
Call ContentStart("Member Data") 
strSQL = "select * from tbl_Players where PlayerHandle='" & CheckString(strPlayerName) & "'"
oRs.Open strSQL, oConn
If oRS.EOF and oRS.BOF Then
	%>
	<CENTER><b><font color=red>Member not found... please check your URL.</font></b></center>
	<%
	Call ContentEnd()
Else
	%>
		<table class="cssBordered" width="50%" align="center">
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2><%=Server.HTMLEncode(strPlayerName)%></TH>
		</TR>
		<%
		intPlayerID = ors.Fields("PlayerID").Value 
		strPlayerName = oRs.Fields("PlayerHandle").Value
		PlayerStatus = "Active"
		if (oRs.Fields("Suspension") = 1) Then
		  PlayerStatus = "SUSPENDED"
		End if
		if (oRs.Fields("PlayerCanActivate") = 0) Then
		  PlayerStatus = "BANNED"
		End If
		if (oRs.Fields("PlayerHandle") = "snoop[sg]") Then
		  PlayerStatus = "No Banned, Yo."
		End If
		%>
		<tr bgcolor=<%=bgctwo%>><td align=right>Email:</td>
		<%
		If (ors.fields("PlayerHideEmail").value <> "1" OR bSysAdmin OR bAnyLadderAdmin) AND bLoggedIn then
			If (bAnyLadderAdmin OR bSysAdmin) Then%>
			<td valign="top"><%=oRs.Fields("PlayerEmail")%></td>
			<%Else%>
			<td valign="top"><%=Replace(Replace("" & oRs.Fields("PlayerEmail").Value, "@", " at "), ".", " dot ")%></td>
			<%End If
		else
			%>
			<td><font color=red>not available</td></tr>
			<%
		end if
		%>
		<tr bgcolor=<%=bgcone%>><td align=right>Xfire:</td><td><%=Server.HTMLEncode(ors.Fields("PlayerICQ").Value)%></td></tr>
		<tr bgcolor=<%=bgctwo%>><td align=right>Date Joined:</td><td><%=Formatdatetime(ors.Fields(2).Value, 2)%></TD></TR>
		<% If Not (PlayerStatus = "Active") Then %>
		<tr bgcolor=<%=bgcone%>><td align=right>Status:</td><td><font color=red><%=Server.HTMLEncode(PlayerStatus)%></font></td></tr>
		<% Else %>
		<tr bgcolor=<%=bgcone%>><td align=right>Status:</td><td><%=Server.HTMLEncode(PlayerStatus)%></td></tr>
		<% End If %>
		<%	
		bgc=bgcone
		if session("uName") = strPlayerName or bSysAdmin then
			%>
			<tr bgcolor=<%=bgcone%>><td align=center colspan=2><input class=bright type=button value="Edit Member" onclick="window.location.href='addplayer.asp?isedit=true&amp;playername=<%=server.urlencode(strPlayerName)%>'" /></td></tr>
			<tr bgcolor=<%=bgcone%>><td align=center colspan=2><input class=bright type=button value="Join A Ladder" onclick="window.location.href='playerjoin.asp?player=<%=server.urlencode(strPlayerName)%>'" id=button1 name=button1 /></td></tr>
			<tr bgcolor=<%=bgcblack%>><td align=center colspan=2><a href="/request/ReqNameChange.asp?player=<%=server.urlencode(strPlayerName)%>">Request Name Change</a></td></tr>
			<%
		End IF
		If bSysAdmin Then
			%>
			<tr bgcolor=<%=bgcblack%>><td align=center colspan=2><a href="/tracker.asp?frm_name=<%=server.urlencode(strPlayerName)%>">Track IP</a></td></tr>
			<%
		End If
		oRS.NextRecordset 
		
		strSQL="Select Count(LadderID) FROM vPlayerTeams WHERE PlayerID='" & intPlayerID & "'"
		oRS.Open strSQL, oconn
		If Not(oRS.EOF and oRS.BOF) then
			If oRS.Fields(0).Value = 0 Then
				response.write "<tr bgcolor=#000000><td colspan=2 align=center><b>No active teams found for selected member.</b></td></tr>"
			End If
		End If
		oRS.NextRecordset 
		%>
		</table>
	<br />
	<%
	Dim iColSpan
	iColSpan = 3
	If bSysAdmin then
		iColSpan = 5
	End If
	%>
		<table class="cssBordered" width="50%" align="center">
	<tr>
		<th colspan="<%=iColSpan%>" bgcolor="#000000">Anti-Smurf In Game Identifiers</th>
	</tr>
	<tr>
		<th bgcolor="#000000">Type</th>
		<th bgcolor="#000000">ID</th>
		<th bgcolor="#000000">Date Added</th>
		<% If bSysAdmin then %>
		<th bgcolor="#000000">Delete</th>
	<% End If
		If bSysAdmin or bAnyLadderAdmin Then %>
		<th bgcolor="#000000">Status</th>
		<% End If %>
		
	</tr>
	<%
	' Game Unique IDs (WONID Hack)
	strSQL = "SELECT lnkPlayerIdentifierID, IdentifierValue, IdentifierName, DateAdded, l.IdentifierActive FROM lnk_player_identifier l INNER JOIN tbl_identifiers i ON i.IdentifierID = l.IdentifierID WHERE PlayerID = '" & intPlayerID & "' AND l.IdentifierActive = 1  AND i.IdentifierActive = 1 ORDER BY IdentifierName ASC"
	If bSysAdmin or bAnyLadderAdmin then
	strSQL = "SELECT lnkPlayerIdentifierID, IdentifierValue, IdentifierName, DateAdded, l.IdentifierActive FROM lnk_player_identifier l INNER JOIN tbl_identifiers i ON i.IdentifierID = l.IdentifierID WHERE PlayerID = '" & intPlayerID & "' ORDER BY IdentifierName ASC"
	End If 
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		Do While Not(oRs.EOF)
			%>
			<tr>
				<td bgcolor="<%=bgc%>"><b><%=oRs.Fields("IdentifierName").Value%></b></td>
				<td bgcolor="<%=bgc%>" align="right"><%=oRs.Fields("IdentifierValue").Value%></td>
				<td bgcolor="<%=bgc%>" align="right"><%=FormatDateTime(oRs.Fields("DateAdded").Value, 2)%></td>
				<% If bSysAdmin then 
				     If oRs.Fields("IdentifierActive").Value Then%>
				<td align="center" bgcolor="<%=bgc%>"><a href="javascript:fConfirmRemoveIdentifier(<%=oRs.Fields("lnkPlayerIdentifierID").Value%>);">delete</a></td>
				<%   Else %>
				<td align="center" bgcolor="<%=bgc%>"></td>
				<%   End If
					End If
					If bSysAdmin or bAnyLadderAdmin Then			
				       If oRs.Fields("IdentifierActive").Value Then%>
				<td bgcolor="<%=bgc%>" align="right"><b><font color=#FFD142>Active</font></b></td>
				<%     Else %>
				<td bgcolor="<%=bgc%>" align="right">Deleted</td>
				<%     End If
				   End If %>
			</tr>
			<%
			If bgc = bgctwo Then
				Bgc = Bgcone
			Else 
				Bgc = Bgctwo
			End If
			oRs.MoveNext
		Loop
		%>
		<%
	Else
		%>
		<tr>
			<td colspan="<%=iColSpan%>" bgcolor="<%=bgcone%>">No registered information found for user.</td>
		</tr>
		<%
	End If
	oRs.NextRecordSet
	If session("uName") = strPlayerName or bSysAdmin then
		%>
		<tr>
			<td colspan="<%=iColSpan%>" bgcolor="#000000" align="center"><a href="IdentifierAdd.asp?player=<%=Server.URLEncode(strPlayerName)%>">add another id</a></td>
		</tr>
		<%
	End If
	%>
	</table>
	
	<%
	Call ContentEnd()
	strSQL = "Select TeamName, TeamFounderid from tbl_Teams where TeamFounderid=" & intPlayerID & " and TeamActive = 1 order by TeamName"
	oRS.Open strsql, oconn
	If Not (oRS.eof and oRS.bof) then
		Call ContentStart("Teams Founded by " & Server.HTMLEncode(strPlayerName))
		%>
		<table class="cssBordered" width="50%" align="center">
		<%
		bgc=bgctwo
		do while not oRS.eof
			%>
			<tr bgcolor=<%=bgc%>><td align=center><a href=viewteam.asp?team=<%=server.urlencode(oRS.fields("TeamName").value)%>><%=Server.HTMLEncode(oRS.fields("TeamName").value)%></a></td></tr>
			<%
			if bgc=bgctwo then
				bgc=bgcone
			else 
				bgc=bgctwo
			end if
			oRS.movenext
		loop
		%>
		</table>
		<%
		Call ContentEnd()
	end if
	oRs.NextRecordset 
	strSQL="Select * FROM vPlayerTeams WHERE PlayerID='" & intPlayerID & "' AND Rank > 0"
	oRS.Open strSQL, oconn	
	if not (oRS.EOF and oRS.BOF) then
		bBarDone = True
		Call ContentStart("Member Teams")
		%>
		<table class="cssBordered" width="50%" align="center">
		<%
		do while not oRS.eof 
			%>	
			<tr bgcolor=<%=bgcone%>>
			<td align=left colspan=2><a href=viewladder.asp?ladder=<%=server.urlencode(oRS.Fields("LadderName").Value )%>><b><%=Server.HTMLEncode(oRS.Fields("LadderName").Value )%></b></a>: <a href=viewTeam.asp?team=<%=server.urlencode(oRS.Fields("TeamName").Value )%>><%=Server.HTMLEncode(oRS.Fields("TeamName").Value )%></a></td>
			</tr>
			<tr bgcolor=<%=bgctwo%>><td valign=center colspan=2><b>&nbsp;&nbsp;Rank:</b>&nbsp;<%=oRS.Fields("Rank").Value%>&nbsp;<b>Record:</b>&nbsp;<%=oRS.Fields("Wins").Value%>/<%=oRS.Fields("Losses").Value%></td></tr>
			<%
			oRS.MoveNext 
		loop
	end if
	oRS.NextRecordset 
			
	strSQL = "SELECT TeamName, et.TeamID, EloLadderName, Rating, Wins, Losses "
	strSQL = strSQL & " FROM lnk_elo_team_player etp INNER JOIN lnk_elo_team et ON etp.lnkEloTeamID = et.lnkEloTeamID "
	strSQL = strSQL & " INNER JOIN tbl_elo_ladders el ON el.EloLadderID = et.EloLadderID "
	strSQL = strSQL & " INNER JOIN tbl_teams t ON t.TeamID = et.TeamID "
	strSQL = strSQL & " WHERE PlayerID='" & intPlayerID & "' AND et.Active = 1 AND el.EloActive = 1 ORDER BY EloLadderName ASC"
	oRS.Open strSQL, oconn	
	if not (oRS.EOF and oRS.BOF) then
		If Not(bBarDone) Then
			bBarDone = True
			Call ContentStart("Member Teams")
			%>
		<table class="cssBordered" width="50%" align="center">
		<%
		End If
		do while not oRS.eof 
			%>	
			<tr bgcolor=<%=bgcone%>>
			<td align=left colspan=2><a href=viewscrimladder.asp?ladder=<%=server.urlencode(oRS.Fields("EloLadderName").Value )%>><b><%=Server.HTMLEncode(oRS.Fields("EloLadderName").Value )%></b></a>: <a href=viewTeam.asp?team=<%=server.urlencode(oRS.Fields("TeamName").Value )%>><%=Server.HTMLEncode(oRS.Fields("TeamName").Value )%></a></td>
			</tr>
			<tr bgcolor=<%=bgctwo%>><td valign=center colspan=2><b>&nbsp;&nbsp;Rating:</b>&nbsp;<%=oRS.Fields("Rating").Value%>&nbsp;<b>Record:</b>&nbsp;<%=oRS.Fields("Wins").Value%>/<%=oRS.Fields("Losses").Value%></td></tr>
			<%
			oRS.MoveNext 
		loop
	end if
	oRS.NextRecordset 
			
	strSQL="Select t.TournamentID, t.TournamentName, t.Active, TeamID, lnk.TMLinkID "
	strSQL=strSQL & "from tbl_tournaments t, lnk_t_m lnk, lnk_t_m_p plnk"
	strSQL = strSQL & " where t.tournamentid = lnk.tournamentid "
	strSQL = strSQL & " AND plnk.TMLinkID=lnk.TMLinkID "
	strSQL = strSQL & " AND plnk.PlayerID=" & intPlayerID
 	strSQL = strSQL & " AND t.Active = 1"
	oRS.Open strSQL, oconn	
	if not (oRS.EOF and oRS.BOF) then
		TMLinkID = oRS.Fields(2).value
		if not(bBarDone) then
			Call ContentStart("Member Teams")
			%>
		<table class="cssBordered" width="50%" align="center">
			<%
			bBarDone= true
		end if
		do while not oRs.eof 
			TournamentName = oRs.Fields("TournamentName").Value
			strSQL = "select DivisionID from tbl_rounds where Team1ID=" & TMlinkID & " or Team2ID =" & TMLinkID
			oRS2.Open strSQL, oconn
			if not (oRS2.EOF and oRS2.BOF) then
				DivID = oRS2.Fields(0).Value
			end if
			oRS2.NextRecordset 
			
			strSQL="select TeamName from tbl_teams where TeamID=" & ors.Fields("TeamID").Value 
			oRS2.Open strSQL, oconn
			if not (oRS2.EOF and oRS2.BOF) then
				strTeamName=oRS2.Fields("TeamName").Value
			end if
			oRS2.NextRecordset 					
			%>	
			<tr bgcolor=<%=bgcone%> valign=top>
			<td align=left colspan=2><a href=/tournament/default.asp?page=brackets&tournament=<%=server.urlencode(TournamentName)%>&div=<%=divid%>><b><%=Server.HTMLEncode(TournamentName)%></b></a>: <a href=viewTeam.asp?team=<%=server.urlencode(strTeamName)%>><%=Server.HTMLEncode(strTeamName)%></a></td>
			</tr>
			<%
			ors.MoveNext 
		loop
	end if
	ors.NextRecordset 

	strSQL="EXECUTE PersonalHomeReturnLeagueTeams '" & intPlayerID & "'"
	oRS.Open strSQL, oconn	
	if not (oRS.EOF and oRS.BOF) then
		if not(bBarDone) then
			Call ContentStart("Member Teams")
			%>
		<table class="cssBordered" width="50%" align="center">
			<%
			bBarDone= true
		end if
		do while not oRs.eof 
			%>	
			<tr bgcolor=<%=bgcone%> valign=top>
			<td align=left colspan=2>
				<b><a href="/viewleague.asp?league=<%=server.urlencode(oRs.Fields("LeagueName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("LeagueName").Value)%> League</a></b>
				<br />
				&nbsp;&nbsp;<a href="viewteam.asp?team=<%=Server.URLEncode(oRs.Fields("TeamName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("TeamName").Value & "")%></a> <br />
				<b>Pct:</b> <%=FormatNumber(oRs.Fields("WinPct").Value / 10000, 3, 0)%> <b>Wins:</b> <%=oRs.Fields("Wins").Value%> <b>Losses:</b> <%=oRs.Fields("Losses").Value%>
				
			</tr>
			<%
			ors.MoveNext 
		loop
	end if
	ors.NextRecordset 
	
	if bBardone then
		%>
		</table>
		<%
		Call ContentEnd()
	end if
	'----------------
	' Player Ladders
	'----------------
	strSQL = "SELECT l.PlayerLadderName, lnk.Rank, lnk.Losses, lnk.PPLLinkID, L.PlayerLadderID, "
	strSQL = strSQL & " lnk.forfeits, lnk.wins, lnk.status "
	strSQL = strSQL & " FROM lnk_p_pl lnk, tbl_playerLadders l "
	strSQL = strSQL & " WHERE lnk.PlayerID = '" & intPlayerID & "' AND lnk.IsActive = 1 "
	strSQL = strSQL & " AND lnk.PlayerLadderID = l.PlayerLadderID "
	strSQL = strSQL & " AND l.active = 1 "
	strSQL = strSQL & " ORDER BY l.PlayerLadderName "
	oRS.Open strSQL, oConn
	If Not(oRS.Eof and oRS.boF) Then
		Call ContentStart("Member Ladders")
		%>
		<table class="cssBordered" width="100%">
		<TR BGCOLOR="#000000">
			<TH WIDTH=150>Ladder Name</TH>
			<TH WIDTH=50>Rank</TH>
			<TH WIDTH=75>Record</TH>
			<TH WIDTH=300>Status</TH>
			<% 
			If Session("uName") = strPlayerName Or bSysAdmin Then
				Response.Write "<TH>&nbsp;</TH>"
				Response.Write "<TH>&nbsp;</TH>"
				bCanAdmin = True
			End If
			%> 
		</TR>
		<%
		bgc = bgcone
		Do While Not(oRS.Eof)
			intPlayerLadderID = oRS("PlayerLadderID")
			strLadderName = oRS("PlayerLadderName")
			intRank = oRS("Rank")
			intLosses = oRS("Losses")
			intforfeits = oRS("forfeits")
			intwins = oRS("wins")
			strStatus = oRS("status")
			LinkID = oRS("PPLLinkID")
					
			Select Case(uCase(strStatus))
				case "ATTACKING"
					strSQL = "SELECT p.PlayerHandle, m.MatchMap1ID, m.MatchDate "
					strSQL = strSQL & " FROM lnk_p_pl lnk, tbl_players p, tbl_PlayerMatches m "
					strSQL = strSQL & " WHERE lnk.PlayerID = p.PlayerID AND m.MatchDefenderID = lnk.PPLLinkID "
					strSQL = strSQL & " AND m.MatchAttackerID = " & linkID
					oRS2.open strsql, oconn
					If not(oRS2.eof and oRS2.bof) then
						map = oRS2("matchMap1ID")
						opponent = oRS2("PlayerHandle")
						mDate = oRS2("MatchDate")
						statusVerbage = strStatus & " vs. <a href=viewplayer.asp?player=" & server.urlencode(opponent) & ">" & Server.HTMLEncode(opponent & "") & "</A> (" & map & ")<BR>" & mDate
					Else
						statusVerbage = " Data Error "
					End if
					oRS2.NextRecordset
				Case "DEFENDING"
					strSQL = "SELECT p.PlayerHandle, m.MatchMap1ID, m.MatchDate "
					strSQL = strSQL & " FROM lnk_p_pl lnk, tbl_players p, tbl_PlayerMatches m "
					strSQL = strSQL & " WHERE lnk.PlayerID = p.PlayerID AND m.MatchAttackerID = lnk.PPLLinkID "
					strSQL = strSQL & " AND m.MatchDefenderID = " & linkID
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
				<% 
					If bCanAdmin OR IsPlayerLadderAdminByID(intPlayerLadderID) Then 
						Response.Write "<TD ALIGN=CENTER BGCOLOR=" & bgc & "><a href=playerladderadmin.asp?player=" & server.URLEncode(strPlayerName) & "&ladder=" & server.URLEncode(strLadderName) & ">Admin</A></TD>"
						Response.write "<TD ALIGN=CENTER BGCOLOR=" & bgc & ">"%>
						<a href="javascript:popup('playerquitLadder.asp?playerid=<%=intPlayerID%>&ladder='+escape('<%=replace(strLadderName & "", "'", "\'")%>')+'&url=<%=Server.URLEncode("viewplayer.asp?player=" & server.urlencode(strPlayerName))%>', 'quitladder', 150, 300, 'no');">Quit</A></TD>
						<%
					End If
				%>
			</TR>
			<%
			If bgc = bgcone then
				bgc = bgctwo
			else
				bgc = bgcone
			End IF
			oRs.MoveNext
		Loop
		oRs.NextRecordset 			
		%>
		</TABLE>
		<br />
		
		<table class="cssBordered" width="100%">
			<TR BGCOLOR="#000000"><TH COLSPAN=5>Recent History</TH></TR>
			<TR BGCOLOR="#000000">
				<TH>Ladder</TH>
				<TH>Opponent</TH>
				<TH>Result</TH>
				<TH>Date</TH>
			</tr>
			<%
			bgc=bgctwo
			strsql="select PPLLinkID from lnk_p_pL where PlayerID=" & intPlayerID
			oRS.Open strsql, oconn
			if not (oRS.EOF and oRS.BOF) then
				do while not (oRS.eof)
					linkID=oRS.Fields(0).Value
					strSQL="select TOP 2 * from vPlayerHistory where (matchwinnerid=" & linkID & " or matchloserid=" & linkID & ")  order by matchdate desc"
					ors2.Open strSQL, oconn
					if not (ors2.eof and ors2.BOF) then
						do while not ors2.EOF
							If ors2.Fields("MatchWinnerID") = linkID Then
								strEnemyName = oRS2.Fields("LoserName").Value 
								strResult = "Win"
							Else
								strEnemyName = oRS2.Fields("WinnerName").Value 
								strResult = "Loss"
							End If
							%>
							<tr bgcolor=<%=bgc%>><td>&nbsp;<a href=viewplayerladder.asp?ladder=<%=server.urlencode(oRS2.Fields("LadderName").Value )%>><%=Server.HTMLEncode(oRS2.Fields("LadderName").Value)%></a></td>
							<td><a href=viewplayer.asp?player=<%=server.urlencode(strEnemyName)%>><%=Server.HTMLEncode(strEnemyName & "")%></a></td>
							<td><%=strResult%></td>
							<td align=right><% If oRS2("MatchForfeit") = 1 Then Response.write "Forfeit" Else  Response.write ors2.Fields("MatchDate").Value End If%>&nbsp;</td></tr>
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
			oRS.NextRecordset 
		%>
		<TR BGCOLOR="#000000">
			<TD COLSPAN=4 ALIGN=CENTER><a href="/playerhistory.asp?player=<%=Server.URLEncode(strPlayerName)%>">Complete History</A></TD>
		</TR>
		</table>
		<%
		Call ContentEnd()
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