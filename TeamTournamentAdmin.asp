<% 'Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Team Administration"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS1 = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% call ContentStart("Team Administration - " & Server.htmlencode(Request.QueryString("team")) & " on the " & Server.htmlencode(Request.QueryString("tournament"))) %>
<table border=0 cellspacing =0 cellpadding=0 align=center width="97%">
<%
	session("CurrentTeam") = Request.QueryString("team")
	session("CurrentTournament") = Request.QueryString("tournament")
	teamname = Request.QueryString("team")
	Tournamentname = Request.QueryString("tournament")
	if teamname = "" or Tournamentname = "" then
		response.redirect "errorpage.asp?error=7"
	end if
	LoggedIn=Session("LoggedIn")
	TeamFounder=IsTeamFounder(teamname)
	TeamCaptain=IsTournamentTeamCaptain(teamname, tournamentName)
	TournamentAdmin=IsTournamentAdmin(tournamentName)
	SysAdmin=IsSysAdmin()
	
	if not(LoggedIn) then
		response.clear
		response.redirect "teamlist.asp"
	end if
	strSQL="select teamid, teamtag from tbl_teams where teamname='" & replace(teamname,"'","''") & "'"
	ors.Open strSQL, oconn
	if not (ors.eof and ors.BOF) then
		teamid=ors.Fields(0).Value
		teamtag=ors.fields(1).value
	end if
	ors.Close
	
	Playerid=Session("PLayerID")
	
	strSQL = "select TournamentID, Maps, Locked from tbl_tournaments where tournamentName='" & replace(tournamentname, "'", "''") & "'"
	ors.Open strSQL, oconn
	if not (ors.eof and ors.BOF) then
		Maps = ors.fields(1).value
		Locked = ors.fields(2).value
		tournamentid=ors.Fields(0).Value
	end if
	ors.Close
	strSQL="select TMlinkID, LastLogin from lnk_T_M where tournamentid=" & tournamentID & " and teamid=" & teamid
	ors.Open strSQL, oconn
	if not (ors.eof and ors.BOF) then
		linkid=ors.Fields(0).Value
		logintime=ors.Fields(1).Value
	end if
	ors.Close
	strSQL="select teamfounderid from tbl_teams where teamid=" & teamid
	ors.Open strSQL, oconn
	if not (ors.eof and ors.BOF) then
		ownerid=ors.fields(0).value
		strsql="select playerhandle from tbl_players where playerid=" & ownerid
		ors2.open strsql,oconn
		if not (ors2.eof and ors2.bof) then
			ownername = ors2.fields(0).value
		else
			ownername="No Founder"
		end if
		ors2.close	
	end if
	ors.Close
	
if not(SysAdmin or TeamCaptain or TeamFounder or TournamentAdmin)  then
	response.clear
	response.redirect "errorpage.asp?error=3"
else
	if (TeamCaptain or TeamFounder) then
		strsql = "update lnk_t_m set LastLogin = '" & now() & "' where TMLinkID= " & linkid 
		ors.open strsql, oconn
	end if
	if (sysadmin or LeagueAdmin) then
		response.write "<TR><TD><center><B>Last Login Time: " & logintime & "</b></center></TD></TR>"
	end if
	if Locked = 1 then
		Response.Write "<TR><TD align=center><p class=small><B>Tournament is only open for signups at this time.</B></p></TD></TR>"
	else

		' Start Tournament Code here...
		' Find current round, and give options to choose maps and do match comms
		strsql = "select *, Team1Name = (select TeamName from tbl_teams t, lnk_t_m lnk where TMLinkID = Team1ID AND lnk.teamid = t.teamid), " &_
					" Team2Name = (select TeamName from tbl_teams t, lnk_t_m lnk where TMLinkID = Team2ID AND lnk.teamid = t.teamid) " &_
					" from tbl_rounds where (Team1ID = '" & linkID & "' or Team2ID = '" & linkid & "') AND WinnerID = 0 order by Round desc"
		ors.open strsql, oconn
		if not(ors.eof and ors.bof) then
			ServerName = oRs.Fields("ServerName").Value
			ServerIP = oRs.Fields("ServerIP").Value
			ServerJoinPassword = oRs.Fields("ServerJoinPassword").Value
			ServerRConPassword = oRs.Fields("ServerRConPassword").Value
			MatchTime = oRs.Fields("MatchTime").Value
			if ors("Team1ID") = linkid then
				RoundsID = ors("RoundsID")
				Team1 = true
				Team1ID = linkID
				Team2ID = ors("Team2ID")
				OpponentLinkID = ors("Team2ID")
				Team1Name = ors("Team1Name")
				Team2Name = ors("Team2Name")
				locationverb = "home"
				opponentname = Team2Name
				WinnerID = ors("WinnerID")
			else
				RoundsID = ors("RoundsID")
				Team1 = false
				Team1ID = ors("Team1ID")
				Team2ID = linkID
				OpponentLinkID = ors("Team1ID")
				Team1Name = ors("Team1Name")
				Team2Name = ors("Team2Name")
				locationverb = "visitor"
				opponentname = Team1Name
				WinnerID = ors("WinnerID")
			end if
			RoundNum = ors("Round")
		end if ' end select rounds
		ors.close
		if winnerID = "0" then
			if team1id = "0" or team2id = "0" then
				Waiting = true
				CurrentStatus = "Awaiting another team to be seeded into your bracket."
			else
				waiting = false
				CurrentStatus = "Challenging <a href=""/viewteam.asp?team=" & server.URLEncode(OpponentName)
				CurrentStatus = CurrentStatus & """>" & Server.htmlencode(OpponentName) & "</a> in round " & roundnum & "."
			end if
			Response.Write "<TR><TD align=center><p class=small><font color=""#DDDDDD""><center>Current Status:" & currentstatus & "</font></p></TD></TR>"
			Response.Write "<TR><TD align=center><p class=small>You are considered the " & locationverb & " team.</P></TD></TR>"
			'Map Stuff here			
			strsql = "select * from lnk_r_m where RoundsID='" & roundsID & "' order by MapOrder asc"
			ors1.open strsql, oconn
			i = 1
			choosemap = false
			dim curmaps (256)
			if not(ors1.eof and ors1.bof) then
				do while not(ors1.eof)
					curmaps(i) = ors1.fields("Map").value
					mapname = ors1.fields("Map").value
					select case MapName
						case "Home", "Visitor", "Random"
							MapName = MapName + " chioce"
					end select
					if ucase(ors1.fields("Map").value) = ucase(locationverb) then
						choosemap = true
					end if
					Response.Write "<TR><TD><p class=small>Map " & ors1.fields("MapOrder").value & ": " & mapname &  "</p></TD></TR>" & vbcrlf 
								
					ors1.movenext
					i = i + 1
				loop
			end if
			
			' need to choose a map
			if choosemap AND not(waiting) then
				ors1.movefirst
			%>
			<TR><TD>
			<form name=mapselection id=mapselection action=/tournament/savetournament.asp method=post>
			<table align=center border=0 cellspacing=0 cellpadding=0 width=200>
			<TR><TD align=center><B>Your team must choose:</b></TD></TR>
			<%
				'Build map list
				strSQL="select tbl_maps.mapname from tbl_maps inner join (lnk_M_M inner join " &_
							"tbl_tournaments on tbl_tournaments.tournamentID=lnk_M_M.tournamentID) " &_
							"on tbl_maps.mapid=lnk_M_M.mapid where tbl_tournaments.TournamentID='" & tournamentID & "'"
				ors2.open strSQL, oconn
				totalmaps = ors2.recordcount
				dim availmap(256)
				if not (ors2.eof and ors2.bof) then
					j=1
					do while not ors2.eof
						addmap = true
						for p = 1 to ubound(curmaps) 
							if ors2.fields(0).value = curmaps(p) then
								addmap = false
							end if
						next
						if addmap then
							availmap(j) = ors2.fields(0).value
							j=j+1
						end if
						ors2.movenext
					loop
				end if
				ors2.nextrecordset
				do while not(ors1.eof)
					mapname = ors1.fields("Map").value
					if ucase(MapName) = ucase(locationverb) then
						Response.Write "<TR bgcolor=" & bgcone & "><TD align=right><B>Map " & ors1.fields("MapOrder").value & ": </B><select name=Map" & ors1.fields("MapOrder").value & " id=Map" & ors1.fields("MapOrder").value & ">"
						Response.Write "<option value ="""">Select a Map</option>" & vbcrlf
						counta = 1
						do while counta < ubound(availmap) AND availmap(counta) <> ""
							Response.Write "<option value=""" & availmap(counta) & """>" & availmap(counta) & "</optoin>" & vbcrlf
							counta = counta + 1
						loop
						response.write "</select></TD></TR>"
							
					end if
		'				maps(i) = ors1.fields("Map").value
		'				select case MapName
		'					case "Home", "Visitor", "Random"
		'						MapName = MapName + " chioce"
		'				end select
		'				Response.Write (ucase(ors1.fields("Map").value) = "VISITOR")
		'				Response.Write "<TR><TD><p class=small>Map " & ors1.fields("MapOrder").value & ": " & mapname &  "</p></TD></TR>" & vbcrlf 
		'						
					ors1.movenext
		'				i = i + 1
				loop
			%>
			<TR bgcolor=<%=bgctwo%> height=35><TD align=center><input type=submit value="Choose Maps" id=submit name=submit></TD></TR>
			</table>
			<input name=team type=hidden id=team value="<%=teamname%>">
			<input name=tournament type=hidden id=tournament value="<%=tournamentname%>">

			<input name=savetype type=hidden id=savetype value="MapSave">
			<input name=roundsID type=hidden id=roundsID value="<%=roundsid%>">
			</FORM>
			</td></tr>
			<%
			end if
			ors1.nextrecordset
			' ENd Map Stuff
			if not (waiting) then
				' Report loss button
				%>
				<TR><TD align=center>
				<script>
					lossurl = "ReportTournamentLoss.asp?teamID=" + <%=TeamID%> + "&tournamentID=" + <%=tournamentid%> + "&roundsid=<%=roundsid%>&linkid=<%=linkid%>&url="+escape(this.location.href);
				</script>
				<form name=loss id=loss>
				<input type=button value="Report Loss" id=reportloss name=reportloss onclick="javascript:popup(lossurl, 'loss', 150, 300, 'no')" style='width:150'>
				</form>
				</td></tr>
			<%
			end if
			if bSysAdmin OR TournamentAdmin then
				' Report win button
				%>
				<TR><TD align=center>
				<script>
					lossurl2 = "ReportTournamentWin.asp?teamID=" + <%=TeamID%> + "&tournamentID=" + <%=tournamentid%> + "&roundsid=<%=roundsid%>&linkid=<%=linkid%>&url="+escape(this.location.href);
				</script>
				<form name=loss id=loss>
				<input type=button value="Report Win" id=reportwin name=reportwin onclick="javascript:popup(lossurl2, 'loss', 150, 300, 'no')" style='width:150'>
				</form>
				</td></tr>
			<%
			end if
			%>
			<TR><TD>
			<% If Not(IsNull(ServerName)) Then %>
			<table align=center border=0 cellspacing=0 cellpadding=0 bgcolor="#444444">
			<tr><td>
			<table align=center width=100% border=0 cellspacing=1 cellpadding=4>
			<tr>
				<th colspan="2" bgcolor="#000000">Server Information</th>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>"><b>Server Name:</b></td>
				<td bgcolor="<%=bgcone%>"><%=ServerName%></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>"><b>Server IP:</b></td>
				<td bgcolor="<%=bgcone%>"><%=ServerIP%></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>"><b>Join Password:</b></td>
				<td bgcolor="<%=bgcone%>"><%=ServerJoinPassword%></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>"><b>Rcon Password:</b></td>
				<td bgcolor="<%=bgcone%>"><%=ServerRConPassword%></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>"><b>Match Time:</b></td>
				<td bgcolor="<%=bgcone%>"><%=MatchTime%></td>
			</tr>
			</table></td></tr></table>
			<% End If %>
			<table align=center width=580 border=0 cellspacing=0 cellpadding=1>
			<%
			' Match comms
			if TeamCaptain or TeamFounder then
				Response.Write "<br><center><p class=small><input type=button name=matchcomm value='Add Match Communication' onclick=""window.location.href='roundcomms.asp?roundsid=" & roundsid & "&mode=add&tag=" & server.urlencode(teamtag) & "';"">"
			else
				Response.Write "<br><center><p class=small><input type=button name=matchcomm value='Add Match Communication' onclick=""window.location.href='roundcomms.asp?roundsid=" & roundsid & "&mode=add&tag=TWLAdmin';"">"
			end if
			strSQL = "select * from tbl_round_Comm where ((roundsid=" & roundsID & ") and (CommDead=0)) order by CommID desc"
			ors2.Open strSQL, oconn
			bgc=bgcone
			if not (ors2.EOF and ors2.bof) then
				do while not ors2.EOF
					Response.Write "<tr bgcolor="& bgc & "><td colspan=2><hr></td></tr><tr bgcolor="& bgc & "><td><p class=small>Author: <b>" & ors2.Fields("CommAuthor").Value & " - Posted: " & ors2.Fields("CommDate").Value & " " & formatdatetime(ors2.Fields("CommTime").Value, 3) & "</p></td>"
					if SysAdmin then
						Response.Write "<td align=right><p class=small><a href=roundcomms.asp?matchid=" & matchID & "&mode=edit&commid=" & ors2.Fields("CommID").Value & " onmouseover=""(window.status='Edit this communication.'); return true"" onmouseout=""(window.status='" & javatitle & "'); return true"">Edit</a> - <a href=/tournament/SaveTournament.asp?commid=" & ors2.Fields("CommID").Value & "&SaveType=Delete_Communications  onmouseover=""(window.status='Delete this communication.'); return true"" onmouseout=""(window.status='" & javatitle & "'); return true"">Delete&nbsp;&nbsp;</a></td></tr>"
					else
						Response.Write "<td><p class=small>&nbsp;</p></td></tr>"
					end if
					Response.write "<tr bgcolor="& bgc &"><td colspan=2>&nbsp;" & ors2.Fields("Comms").Value & "</td></tr>"
					if bgc = bgcone then
						bgc=bgctwo
					else
						bgc=bgcone
					end if
					ors2.MoveNext
				loop
			end if
			%>
			</table>
			</TD></TR>
			<%
			ors2.nextrecordset
			' End Match Comms
				
		end if ' WinnerID
	END IF ' Locked End if

Response.Write "</table>"
Call ContentEnd() 

' Roster...
strSQL = "select tbl_players.playerhandle from tbl_players inner join lnk_T_M_P on lnk_T_M_P.PlayerID=tbl_Players.playerid " &_
	"where (lnk_T_M_P.TMLinkID=" & linkID & " and lnk_T_M_P.isadmin=1) order by tbl_players.playerhandle"
ors.Open strSQL, oconn


Call Content3BoxStart(Server.htmlencode(request("tournament")) & " Captain Management") 

Response.Write "<table border=0 align=center width=97% cellspacing=0>"
Response.Write "<tr height=30 bgcolor="&bgcone&"><td align=center><p class=small><b>Current Captains</b></p></td></tr>"
if not (ors.EOF and ors.BOF) then
	bgc=bgcone
	do while not ors.EOF
		info=""
		if ors.fields(0).value = ownername then 
			info=" (founder)"
		end if	
		if bgc=bgctwo then
			bgc=bgcone
		else bgc=bgctwo
		end if
		Response.Write "<tr height=30 bgcolor="&bgc&"><td align=center>" & Server.htmlencode(ors.Fields(0).Value) & info & "</td></tr>"
		ors.MoveNext			
	loop
end if
response.write "</table>" & VBcRlF & VBcRlF & VBcRlF
ors.close

Call Content3BoxMiddle1()
		 
		strSQL = "select tbl_players.playerhandle, lnk_T_M_P.TMPLinkID from tbl_players " &_
					" inner join lnk_T_M_P on lnk_T_M_P.PlayerID=tbl_Players.playerid " &_
					" where (lnk_T_M_P.TMLinkID=" & linkID & " and lnk_T_M_P.isadmin=0) order by tbl_players.playerhandle"
		ors.Open strSQL, oconn
		if not(ors.EOF and ors.BOF) then
			Response.Write "<form name=promote action=/tournament/savetournament.asp method=post>"
			Response.Write "<table align=center border=0 width=97% cellspacing=0><tr bgcolor="&bgcone&" height=30><td align=center><p class=small><b>Promote Player to Captain</b></p></td></tr>"
			Response.Write "<tr bgcolor="&bgctwo&"><td height=30 align=center><select name=playerlist style='width:150'>"
			do while not ors.EOF
				Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.htmlencode(ors.Fields(0).Value)
				ors.MoveNext
			loop
			Response.Write "</select></td></tr><tr height=30 bgcolor="&bgcone&"><td align=center>"
			%>
			<input class=bright type=submit value='Promote'>
			<input type=hidden name=SaveType value=PromoteCaptain>
			<input type=hidden name=tournament value="<%=Server.htmlencode(Request.QueryString("tournament"))%>">
			<input type=hidden name=team value="<%=Server.htmlencode(teamname)%>">
			</td></tr></table></form>
			<%
		end if
ors.Close

Call Content3BoxMiddle2() 
		
		strSQL = "select tbl_players.playerhandle, lnk_T_M_P.TMPLinkID from tbl_players " &_
					"inner join lnk_T_M_P on lnk_T_M_P.PlayerID=tbl_Players.playerid " &_
					" where (lnk_T_M_P.TMLinkID=" & linkID & " and lnk_T_M_P.isadmin=1 " &_ 
					" and tbl_players.playerhandle <> '" & CheckString(session("uName")) &"') order by tbl_players.playerhandle"
		ors.Open strSQL, oconn
		if not(ors.EOF and ors.BOF) then
			Response.Write "<form name=demote action=/tournament/savetournament.asp method=post>"
			Response.Write "<table align=center border=0 width=97% cellspacing=0><tr bgcolor="&bgcone&" height=30>"
			Response.Write "<td align=center><p class=small><b>Demote Captain</b></p></td></tr>"
			Response.Write "<tr bgcolor="&bgctwo&" height=30><td align=center><select name=playerlist style='width:150'>"
			do while not ors.EOF
				if ors.fields(0).value <> ownername then
					Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.htmlencode(ors.Fields(0).Value & "") & "</option>" & vbCrLf
				end if
				ors.MoveNext
			loop
			Response.Write "</select></td></tr><tr bgcolor="&bgcone&" height=30><td align=center>"
			Response.Write "<input class=bright type=submit id=submit2 name=submit2 value=Demote>"
			Response.Write "<input type=hidden name=SaveType value=DemoteCaptain>"
			Response.Write "<input type=hidden name=tournament value=""" & Request.QueryString("tournament") & """>"
			Response.Write "<input type=hidden name=team value=""" & Request.QueryString("team") & """></td></tr></table></form>"	
		end if
		ors.Close

Call Content3BoxEnd()

Call ContentStart(Server.htmlencode(request("tournament")) & " Roster Management")

	strSQL="select PlayerHandle, lnk_T_M_P.DateJoined, lnk_T_M_P.IsAdmin, tbl_players.PlayerID from tbl_Players inner join "
	strsql= strsql & "lnk_T_M_P on lnk_T_M_P.PlayerID=tbl_players.playerid where lnk_T_M_P.TMLinkID=" 
	strsql= strsql & LinkID & " order by PlayerHandle"			
	ors.open strsql,oconn
	if not (ors.eof and ors.bof) then
		response.write "<form name=BootPlayer method=post action=/tournament/savetournament.asp><table width=50% align=center border=0 cellspacing=0 cellpadding=0>"
		response.write "<tr bgcolor="&bgcone&" height=125 valign=center><td align=center><input type=hidden name=savetype value=DropPlayer><select name=PlayerID size=5 class=brightred style='width:200'>"
		do while not ors.eof
			if ors.fields(3).value <> ownerid then 
				response.write "<option value=" & ors.fields(3).value & ">" & Server.htmlencode(ors.fields(0).value)
			end if
			ors.movenext
		loop
		response.write "</select></td></tr><tr bgcolor="&bgctwo&" height=35><td align=center><input type=hidden name=link value=" & linkid & "><input type=submit class=bright style='width:75' value='Kick Player'></td></tr></table></form>"
	end if
	ors.close
end if

Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRS2 = Nothing
%>
