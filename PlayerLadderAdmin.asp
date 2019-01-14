<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Player Ladder Administration"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
Dim PlayerName, LadderName, PlayerID, LinkID, Wins, Losses
Dim Rank, IsActive, ForFeits, LadderID, Locked, Active, Challenge, LastLogin
Dim mStatus, enemyID, enemyName, mDate, xDate, aDate, Map1, MatchDate1, MatchDate2, MatchID
Dim MaxJump, MaxChallenge, tima, CurrentLadder, LadderChallenge
Dim grammer

CurrentLadder = Request.QueryString("ladder")
PlayerName = Request("Player")
LadderName = Request.QueryString("ladder")
if LadderName = "" then
	oConn.Close
	Set oCOnn = Nothing
	Set oRS = Nothing
	Response.Clear 
	Response.Redirect  "errorpage.asp?error=7"
end if

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Ladder Administration - " & PlayerName & " on the " & LadderName & " Ladder") %>

<%
	If PlayerName <> "" AND Not(bSysAdmin) AND NOT(IsPlayerLadderAdmin(LadderName))  And Not(Session("uName") = PlayerName) Then
		oConn.Close
		Set oCOnn = Nothing
		Set oRS = Nothing
		Response.clear
		Response.Redirect "/errorpage.asp?error=3"
	ElseIf PlayerName <> "" Then
		strSQL = "select l.PlayerLadderID, lnk.PPLLinkID, p.PlayerID, lnk.Wins, lnk.LastLogin, "
		strSQL = strSQL & " lnk.Losses, lnk.Rank, lnk.IsActive, lnk.ForFeits, l.Locked, l.Active, l.Challenge, lnk.status "
		strSQL = strSQL & " FROM lnk_p_pl lnk, tbl_players p, tbl_playerladders l "
		strsql = strSQL & " WHERE  P.PlayerHandle = '" & replace(PlayerName, "'", "''") & "' "
		strsql = strSQL & " AND p.PlayerID = lnk.PlayerID "
		strsql = strSQL & " AND l.PlayerLadderID = lnk.PlayerLadderID "
		strsql = strSQL & " AND lnk.IsActive = 1 "
		strsql = strSQL & " AND l.Active = 1 "
		strsql = strSQL & " AND l.PlayerLadderName = '" & replace(ladderName, "'", "''") & "'"
	Else
		strSQL = "select l.PlayerLadderID, lnk.PPLLinkID, p.PlayerID, lnk.Wins, lnk.LastLogin, "
		strSQL = strSQL & " lnk.Losses, lnk.Rank, lnk.IsActive, lnk.ForFeits, l.Locked, l.Active, l.Challenge, lnk.status "
		strSQL = strSQL & " FROM lnk_p_pl lnk, tbl_players p, tbl_playerladders l "
		strsql = strSQL & " WHERE  P.PlayerHandle = '" & replace(session("uName"), "'", "''") & "' "
		strsql = strSQL & " AND p.PlayerID = lnk.PlayerID "
		strsql = strSQL & " AND l.PlayerLadderID = lnk.PlayerLadderID "
		strsql = strSQL & " AND lnk.IsActive = 1 "
		strsql = strSQL & " AND l.PlayerLadderName = '" & replace(LadderName, "'", "''") & "'"
	End If
	ors.Open strSQL, oconn
	if not (ors.eof and ors.BOF) then
		Playerid=ors("PlayerID")
		LinkID = ors("PPLLinkID")
		Wins = ors("Wins")
		Losses = ors("Losses")
		Rank = ors("Rank")
		IsActive = ors("IsActive")
		ForFeits = ors("ForFeits")
		LadderID = ors("PlayerLadderID")
		Locked = ors("Locked")
		Active = ors("Active")
		Challenge = ors("Challenge")
		LastLogin = ors("LastLogin")
		mStatus = ors("Status")
	Else
		oConn.Close
		Set oCOnn = Nothing
		Set oRS = Nothing
		Response.Clear
		Response.Redirect "/errorpage.asp?error=7"
	End If
	ors.Close
	
	if (PlayerName <> "") then
		strsql = "update lnk_p_pL set LastLogin = '" & now() & "' where PPLLinkID= " & LinkID 
		oConn.Execute(strSQL)
	end if
	if (bSysAdmin) then
		response.write "<center><B>Last Login Time: " & LastLogin & "</b></center>"
	end if

	Response.Write "<center><p class=text><b>Current Rung:</b> " & Rank & "</p></center>"
If Locked = 0 Then
	if mstatus="Attacking" then
			strSQL = "Select MatchDefenderID, MatchDate, MatchChallengeDate, MatchAcceptanceDate, p.PlayerHandle, "
			strSQL = strSQL & " pm.MatchMap1ID, pm.MatchSelDate1, pm.MatchSelDate2, pm.PlayerMatchID "
			strSQL = strSQL & " FROM tbl_PlayerMatches pm, tbl_players p, lnk_p_pl lnk "
			strSQL = strSQL & " WHERE pm.MatchAttackerID = " & LinkID & " and pm.matchladderid=" & LadderID
			strSQL = strSQL & " AND lnk.PPLLinkID = pm.MatchDefenderID "
			strSQL = strSQL & " AND p.PlayerID = lnk.PlayerID "
			ors.Open strSQL, oconn
			if not (ors.EOF and ors.BOF) then
				enemyID = ors("MatchDefenderID")
				enemyName = ors("PlayerHandle")
				mDate = ors("MatchDate")
				xDate = ors("MatchChallengeDate")
				aDate = ors("MatchAcceptanceDate")
				map1=ors("MatchMap1ID")
				matchdate1 = ors("MatchSelDate1")
				matchdate2 = ors("MatchSelDate2")
				matchid = ors("PlayerMatchID")
			end if
			ors.Close
			Response.Write "<center><p class=small><b>Current Status:</b></font><font size=2 color=#c0c0c0> Attacking <a href=viewplayer.asp?player=" & server.urlencode(enemyname) & ">"&Server.HTMLEncode(enemyname)&"</a> <br>Map: " & Server.HTMLEncode(map1) & "<BR>"
			response.write "Match Date: " & mdate & "</center><center>"
			if (bsysadmin) and mdate <> "TBD" then %>
				<br><center><p class=small><input value="Clear Match Date" type=button class=brightgold style="width:200" onclick="window.location.href='ladder/playerladderengine.asp?clearDate=true&matchid=<%=matchid%>';" id=clearDate name=clearDate>
				<form name=changeDate action=ladder/playerladderengine.asp method=post>
					<input type="hidden" name="saveType" value="changeDate">
					<input type="hidden" name="matchid" value="<%=matchid%>">
					Month: <input type="Text" name="newMonth" value="" size="2" maxlength="2">&nbsp;Day:<input type="Text" name="newDay" value="" size="2" maxlength="2">&nbsp;<%=year(now)%><br>
					<input type="hidden" name="newYear" value="<%=year(now)%>">
					New Time: <input type="Text" name="newHour" value="" size="2" maxlength="2">:<input type="Text" name="newMinute" value="" size="2" maxlength="2">:00 PM<br>
					<input type="submit" class=brightgold style="width:150" value="Change Date" id=submit4 name=submit4>
				</form><%
			end if
			if matchdate1="TBD" then
				Response.Write "<center><p class=small><b>Awaiting match acceptance from " & enemyname & " (Challenged on " & xDate & ")</b></center>"
			elseif map1 = "TBD" then
				if right(ucase(enemyname),1)="S" then
					grammer=" have "
				else
					grammer=" has "
				end if
				Response.Write "<center><p class=small><b>Challenge accepted by " & enemyname & " on " &  aDate & "</b></center>"
				Response.Write "<center><p class=small>" & Server.HTMLEncode(enemyname) & grammer & "selected the match dates listed below. click on the date to accept</center>"
				%>
				<table align=center width=400 border=0 cellspacing=0>
					<form name=frmAccept action=../saveitem.asp method=post>
					<tr><td align=center>Chosen Dates: <select name=matchdate class=brightred><option><%=matchdate1%><option><%=matchdate2%></select></td></tr>
					<tr><td align=center>
						<input type=hidden name=matchid value=<%=matchid%>>
						<input type=hidden name=SaveType value=PlayerAcceptMatchDate>
						<input type=hidden name=LadderName value="<%=LadderName%>">
						<input type=hidden name=PlayerName value="<%=PlayerName%>">
						<input type=submit name=submit1 value="Confirm Date and Time" class=bright>
					</td></tr>
					</form>
				</table>
				<%
			else
				Response.Write "<br><table width=45% align=center ><tr bgcolor="& bgcone&" height=35 valign=center><td align=center><a href=PlayerMatchReportLoss.asp?matchid=" & matchid & "&player=" & server.urlencode(playerName) & "&linkid=" & LinkID & ">Report Loss</a></td></tr></table>"
			end if
	elseif mstatus="Defending" then
			strSQL = "Select MatchAttackerID, MatchDate, MatchChallengeDate, MatchAcceptanceDate, p.PlayerHandle, "
			strSQL = strSQL & " pm.MatchMap1ID, pm.MatchSelDate1, pm.MatchSelDate2, pm.PlayerMatchID "
			strSQL = strSQL & " FROM tbl_PlayerMatches pm, tbl_players p, lnk_p_pl lnk "
			strSQL = strSQL & " WHERE pm.MatchDefenderID = " & LinkID & " and pm.matchladderid=" & LadderID
			strSQL = strSQL & " AND lnk.PPLLinkID = pm.MatchAttackerID "
			strSQL = strSQL & " AND p.PlayerID = lnk.PlayerID "
			ors.Open strSQL, oconn
			if not (ors.EOF and ors.BOF) then
				enemyID = ors("MatchAttackerID")
				enemyName = ors("PlayerHandle")
				mDate = ors("MatchDate")
				xDate = ors("MatchChallengeDate")
				aDate = ors("MatchAcceptanceDate")
				map1=ors("MatchMap1ID")
				matchdate1 = ors("MatchSelDate1")
				matchdate2 = ors("MatchSelDate2")
				matchid = ors("PlayerMatchID")
			end if
			ors.Close
			Response.Write "<center><p class=small><b>Current Status:</b></font><font size=2 color=#c0c0c0> Defending vs <a href=viewplayer.asp?player=" & server.urlencode(enemyname) & ">"&Server.HTMLEncode(enemyname)&"</a> <br>Map: " & Server.HTMLEncode(map1) & "<BR>"
			response.write "<br>Match Date: " & mdate & "</font></center><center>"
			if (bsysadmin) and mdate <> "TBD" then %>
				<br><center><p class=small><input value="Clear Match Date" type=button class=brightgold style="width:200" onclick="window.location.href='ladder/playerladderengine.asp?clearDate=true&matchid=<%=matchid%>';" id=clearDate name=clearDate>
				<form name=changeDate action=playerladderengine.asp method=post>
					<input type="hidden" name="saveType" value="changeDate">
					<input type="hidden" name="matchid" value="<%=matchid%>">
					Month: <input type="Text" name="newMonth" value="" size="2" maxlength="2">&nbsp;Day:<input type="Text" name="newDay" value="" size="2" maxlength="2">&nbsp;<%=year(now)%><br>
					<input type="hidden" name="newYear" value="<%=year(now)%>">
					New Time: <input type="Text" name="newHour" value="" size="2" maxlength="2">:<input type="Text" name="newMinute" value="" size="2" maxlength="2">:00 PM<br>
					<input type="submit" class=brightgold style="width:150" value="Change Date" id=submit4 name=submit4>
				</form><%
			end if

			if  (matchdate1 = "TBD" and map1 = "TBD") then
				Response.Write "<center><font size=2><a href=playeracceptmatch.asp?matchid=" & matchid & "&player=" & server.urlencode(playerName) & "&ladder=" & server.urlencode(LadderName) & "&enemy=" & server.urlencode(enemyname) & ">Accept the Challenge from " & Server.HTMLEncode(enemyname) & "</a></font><center><center><font size=1>(You were challenged on " & xDate & ")</font></center>"
			elseif map1="TBD" then
				Response.Write "<center><font size=2>You have chosen <b>" & matchdate1 & "</b> and <b>" & matchdate2 & "</b> for match dates.</center>"
				Response.Write "<center>Awaiting acceptance from " & Server.HTMLEncode(enemyname) & "</font><center><center><font size=1>(You were challenged on " & xDate & ")</font></center>"
			else
				Response.Write "<br><table width=45% align=center ><tr bgcolor="& bgcone&" height=35 valign=center><td align=center><a href=PlayerMatchReportLoss.asp?matchid=" & matchid & "&player=" & server.urlencode(playerName) & "&linkid=" & LinkID & ">Report Loss</a></td></tr></table>"
			end if
	elseif (mstatus="Available" or left(mStatus,6)="Immune") then
		if (Locked = 0) then
			MaxJump=int(Rank/2) + 1
			if MaxJump < (Challenge + 1) then
				MaxJump = (Challenge + 1)	
			end if
			strSQL = "select lnk.Status, p.PlayerHandle, lnk.Rank, l.PlayerLadderName, lnk.LadderMode, lnk.PPLLinkID "
			strSQL = strSQL & "from lnk_P_pl lnk, tbl_playerladders l, tbl_players p "
			strSQL = strSQL & "where p.PlayerID = lnk.PlayerID AND "
			strSQL = strSQL & "lnk.Playerladderid = l.Playerladderid AND "
			strSQL = strSQL & "(lnk.status = 'Available' or left(lnk.Status,8)='Defeated') AND "
			strSQL = strSQL & "lnk.rank < " & Rank & " AND lnk.rank > " & (Rank - MaxJump) & " AND l.PlayerLadderName = '" & CheckString(LadderName) & "' AND "
			strSQL = strSQL & "lnk.LadderMode < 2  AND lnk.isActive = 1 "
			strSQL = strSQL & "AND " 
			strSQL = strSQL & "lnk.PPLLinkID NOT in ( "
			strSQL = strSQL & "select PPLLinkID = MatchWinnerID from tbl_Playerhistory where MatchLoserID = '" & linkid & "' AND DateDiff(d, matchDate, GetDate()) < 7 "
			strSQL = strSQL & "union "
			strSQL = strSQL & "select PPLLinkID = MatchLoserID from tbl_Playerhistory where MatchWinnerID= '" & linkid & "' AND DateDiff(d, matchdate, GetDate()) < 7 "
			strSQL = strSQL & ") "
			strSQL = strSQL & "order by lnk.rank "
			ors.Open strSQL, oconn
			Response.Write "<table align=center border=0 cellspacing=0><tr><td bgcolor="&bgcone&" align=center height=22 width=250><p class=small><b>" & mstatus & " - Click to Challenge</b></P></td></tr>"
			bgc=bgctwo
			if not (ors.EOF and ors.BOF) then
				do while not ors.EOF
					Response.Write "<tr bgcolor=" & bgc & " height=20><td align=center height=22 width=250>"
					Response.Write "<p class=small><a href=playerchallenge.asp?opponent=" & server.urlencode(ors.Fields("PlayerHandle").Value) 
					Response.Write "&ladder=" & server.urlencode(LadderName) & "&PlayerName=" & server.urlencode(PlayerName) & ">" & Server.HTMLEncode(ors.Fields("PlayerHandle").Value) & "</a> - Rung: " & ors.Fields("Rank").Value & "</p></td></tr>"		
					if bgc = bgcone then
						bgc=bgctwo
					else
						bgc=bgcone
					end if
					ors.MoveNext
				loop
			end if
			ors.Close 
			Response.Write "</table>"
		else
			if Locked = 1 then 
				Response.Write "<p align=center><b><i>This ladder is not open for challenging at this time.</i></b></p>"
			end if
		end if
	elseif left(mStatus,6)="Defeat" then
		strsql = "select modeflagtime from lnk_p_pl where PPllinkid=" & linkID
		ors.open strsql, oconn
		tima = ors.fields(0).value
		ors.close		
			Response.Write "<center><p class=text><b>Current Status:</b> " & mstatus & "</p></center>"
			response.write "<center><P class=text><b>Loss reported at: " & tima & "</p></center>"
			response.write "<center><P class=text><b>Able to challenge up 24-hours after report time.</p></center>"
	else
			Response.Write "<center><p class=text><b>Current Status:</b> " & mstatus & "</p></center>"
	end if
	if mStatus = "Attacking" or mStatus="Defending" then
		Response.Write "<br><br><center><font size=2><b><u>Match Communications</u></b></center>"
		Response.Write "<br><center><p class=small><input type=button name=matchcomm value='Add Match Communication' onclick=""window.location.href='playermatchcomms.asp?matchid=" & matchid & "&mode=add&ladder=" & server.urlencode(LadderName) & "&PlayerName=" & server.urlencode(PlayerName) & "';"">"
		strSQL = "select * from tbl_playerComms where (PlayermatchID=" & matchID & ") and (CommDead=0) order by CommID desc"
		ors.Open strSQL, oconn
		%>
		<table align=center width=580 border=0 cellspacing=0 cellpadding=1>
		<%
		bgc=bgcone
		if not (ors.EOF and ors.bof) then
			do while not ors.EOF
				Response.Write "<tr bgcolor="& bgc & "><td colspan=2><hr></td></tr><tr bgcolor="& bgc & "><td><p class=small>Author: <b>" & ors.Fields("CommAuthor").Value & " - Posted: " & formatdatetime(ors.Fields("CommTime").Value, 0) & "</p></td>"
				if bSysAdmin then
					Response.Write "<td align=right><p class=small><a href=playermatchcomms.asp?matchid=" & matchID & "&mode=edit&commid=" & ors.Fields("CommID").Value & ">Edit</a> - <a href=SaveItem.asp?commid=" & ors.Fields("CommID").Value & "&SaveType=playerDelete_Communications&ladder=" & Server.URLEncode(Request.QueryString("Ladder")) & "&player=" & Server.URLEncode(Request.QueryString("player")) & ">Delete</a>&nbsp;&nbsp;</td></tr>"
				else
					Response.Write "<td><p class=small>&nbsp;</p></td></tr>"
				end if
				Response.write "<tr bgcolor="& bgc &"><td colspan=2>&nbsp;" & ors.Fields("Comms").Value & "</td></tr>"
				if bgc = bgcone then
					bgc=bgctwo
				else
					bgc=bgcone
				end if
				ors.MoveNext
			loop
		end if
		%>
		</table>
		<%
		ors.Close 
	end if
Else
	%>
	<CENTER><B>Ladder is locked at this time.</B></CENTER>
	<%
End If
	%>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>
