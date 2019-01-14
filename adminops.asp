<% Option Explicit %>
<%
Response.Buffer = True
Response.Expires = -1440
Dim strPageTitle

strPageTitle = "TWL: Admin Operations"

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

If Not(bSysAdmin Or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim aType

aType = Request.Form("rAdmin")
if aType="" then
	aType= Request.QueryString("rAdmin")
end if
if atype="" then
	atype= Request.QueryString("aType")
end if
if aType="History" then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	response.redirect "edithistory.asp"
end if
if atype="Rank" then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	response.redirect "editrank.asp"
end if

Dim ErrorCode, xTra, sField, intGameID
intGameID = -1
errorcode = request("error")
'Match/Forfeit
Dim LadderName, LadderID, mDate, newMDate, j
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
if (aType = "Match" or aType="Forfeit") then
	Call ContentStart(aType & " Administration")
	if aType="Forfeit" then
		xtra=" and MatchAwaitingForfeit=1 "
	else
		xtra=""
	end if
	If bSysAdmin Then
		strSQL = " SELECT GameName, G.GameID, LadderName, LadderID, ForFeits = ( select count(*) from tbl_matches where matchawaitingforfeit=1 and left(forfeitreason, 5) <> 'Admin' AND matchladderid=l.LadderID) "
		strSQL = strSQL & " FROM tbl_ladders l, tbl_games g WHERE g.GameID = l.GameID AND LadderShown = 1 AND LadderActive = 1 ORDER BY G.GameName, LadderName"	
	Else
		strSQL = " SELECT GameName, G.GameID, LadderName, l.LadderID, ForFeits = ( select count(*) from tbl_matches where matchawaitingforfeit=1 and left(forfeitreason, 5) <> 'Admin' AND matchladderid=l.LadderID) "
		strSQL = strSQL & " FROM tbl_ladders l, lnk_l_a lnk, tbl_games G WHERE lnk.LadderID = l.LadderID AND LadderShown = 1 AND g.GameID = l.GameID AND lnk.PlayerID='" & Session("PlayerID") & "' AND LadderActive = 1 ORDER BY GameName, LadderName"	
	End If
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		bgc = bgcone
		j = 0
		do while not ors.EOF
			If intGameID <> oRS.Fields("GameID").Value THen
				If intGameID <> -1 Then 
					While j mod 3 <> 0
						Response.Write "<TD WIDTH=""33%"">&nbsp;</TD>"
						j = j + 1
					Wend
					Response.Write "</TR></TABLE></TD></TR></TABLE><BR><BR>"
				End If
				intGameID = oRS.Fields("GameID").Value 
				%>
				<table class="cssBordered" width="100%">
				<TR BGCOLOR="#000000">
					<TH COLSPAN=3><%=oRS.Fields("GameName").Value %></TH>
				</TR>
				<%
			End If
			If j Mod 3 = 0 Then 
				If j <> 0 Then
					Response.Write "</TR>"
				End If
				If bgc=bgcone then
					bgc=bgctwo
				else
					bgc=bgcone
				End If
				Response.Write "<tr bgcolor=" & bgc & ">"
			End If
			
			Response.Write "<td align=LEFT WIDTH=""33%"">"
			If oRS.Fields("ForFeits").Value > 0 Then
				Response.Write "<FONT COLOR=""RED"">" & oRS.Fields("ForFeits").Value & "</FONT>"
			Else
				Response.Write "0"
			End If
			Response.Write " - <a href=adminops.asp?rAdmin=" & aType & "&ladderid=" & ors.Fields("LadderID").Value & "&laddername=" & server.urlencode(ors.Fields("LadderName").Value) & ">" & Server.HTMLEncode(ors.Fields("LadderName").Value) & "</a></td>"
			j = j + 1
			ors.Movenext
		loop
		While j mod 3 <> 0
			Response.Write "<TD>&nbsp;</TD>"
			j = j + 1
		Wend
		Response.Write "</TR>"
		%>
		</TABLE>
		<%
	end if
	oRS.NextRecordset 

	sField = request.querystring("sfield")
	if sfield = "" then
		sField = "MatchID"
	End If
	LadderName = request("laddername")
	if Len(LadderName) = 0 Then
		LadderName = request("ladder")
	end if
	ladderid = Request.QueryString("ladderid")
	if ladderid <> "" then
		bgc = bgcone
		%>
		<form name=frmMatchAdmin action=saveItem.asp method=POST>
		<input type=hidden name=SaveType value=admKillMatch>
		<input type=hidden name=ladderid value=<%=ladderid%>>
		<input type=hidden name=rAdmin value=<%=aType%>>
		<input type=hidden name=ladderName value="<%=LadderName%>">
				<table class="cssBordered" width="100%">
		<TR BGCOLOR=#000000>
			<TH COLSPAN=6><%=LadderName%></TH>
		</TR>
		<tr bgcolor="#000000">
			<TH><a href="adminops.asp?rAdmin=<%=aType%>&sField=MatchID&ladderid=<%=ladderid%>&laddername=<%=Server.URLEncode(laddername & "")%>">Match ID</a></TH>
			<TH><a href="adminops.asp?rAdmin=<%=aType%>&sField=DefenderName&ladderid=<%=ladderid%>&laddername=<%=Server.URLEncode(laddername & "")%>">Defender</a></TH>
			<TH><a href="adminops.asp?rAdmin=<%=aType%>&sField=AttackerName&ladderid=<%=ladderid%>&laddername=<%=Server.URLEncode(laddername & "")%>">Attacker</a></TH>
			<TH><a href="adminops.asp?rAdmin=<%=aType%>&sField=MatchDate&ladderid=<%=ladderid%>&laddername=<%=Server.URLEncode(laddername & "")%>">Match Date</a></TH>
			<TH>Match Comms</TH>
			<TH>Operation</TH>
		</TR>
		<%
		bgc=bgctwo
		strSQL= "select * FROM vMatches WHERE MatchLadderId=" & ladderid & xtra & " ORDER BY " & sField
		oRs.Open strSQL, oConn
		if not (ors.EOF and ors.BOF) then
			do while not ors.EOF
				mDate=ors.fields("MatchDate").value
				if mdate <> "TBD" then
					newMDate = right(mDate, len(mDate)-instr(1, mDate, ","))
					newMDate = left(newmdate, len(newmdate)-4)
					newmdate=formatdatetime(newmdate, 2)
				else
					newmdate=mdate
				end if
				
				Response.Write "<tr bgcolor=" & bgc & "><td valign=top>" & ors.Fields("MatchID").Value & "</td>"
				Response.Write "<td valign=top><a href=viewteam.asp?team=" & server.urlencode(oRS.Fields("DefenderName").Value ) & ">" & oRS.Fields("DefenderName").Value  & "</a></td>"
				Response.Write "<td valign=top><a href=viewteam.asp?team=" & server.urlencode(oRS.Fields("AttackerName").Value ) & ">" & oRS.Fields("AttackerName").Value  & "</a></td>"
				Response.Write "<td valign=top>" & newMDate & "</td>"
				Response.Write "<td valign=top><a href=""teamladderadmin.asp?team=" & server.urlencode(oRS.Fields("AttackerName").Value ) & "&ladder=" & Server.URLEncode(laddername & "") & """>Match Comms</a></td>"
				Response.Write "<td valign=top rowspan=2><SELECT NAME=""AdminMatch_" & ors.Fields("MatchID").Value & """>"
				Response.Write "<OPTION SELECTED VALUE=""NOTHING"">Do Nothing</OPTION>" & vbCrLf 
				Response.Write "<OPTION VALUE=""OVERRIDE"">Admin Override</OPTION>" & vbCrLf 
				Response.Write "<OPTION VALUE=""KILL"">Kill Match</OPTION>" & vbCrLf 
				Response.Write "<OPTION VALUE=""FORFEITD"">Defender Forfeit</OPTION>" & vbCrLf 
				Response.Write "<OPTION VALUE=""FORFEITA"">Attacker Forfeit</OPTION>" & vbCrLf 
				Response.Write "</SELECT></td></tr>"
				if ors.fields("ForfeitReason").value = "xxx" then
					Response.Write "<tr><td colspan=5 valign=top bgcolor=" & bgc & " align=center>&nbsp;</td></tr>"
				else 
					Response.Write "<tr><td colspan=5 valign=top bgcolor=" & bgc & " align=center><font color=#ff0000>" & ors.fields("ForFeitReason").value & "</font></td></tr>"
				end if
				ors.MoveNext
				if bgc=bgcone then
					bgc=bgctwo
				else
					bgc=bgcone
				end if
			loop
		end if
		ors.Close 
		%>
		<tr bgcolor=<%=bgc%>><td colspan=6 align=center><input type=submit name=submit1 value=Submit class=bright></td></tr>
		</table>
		</form>
		<%
	end if
	Call ContentEnd()
end if

if (aType = "PMatch" or aType="PForfeit") then
	Call ContentStart(aType & " Administration")
	if aType="PForfeit" then
		xtra=" and MatchAwaitingForfeit=1 "
	else
		xtra=""
	end if
	If bSysAdmin Then
		strSQL = "SELECT g.GameName, g.GameID, l.PlayerLadderID, PlayerLaddername, ForFeits = ( select count(PlayerMatchID) from vPlayerMatches where matchawaitingforfeit=1 and left(forfeitreason, 5) <> 'Admin' AND matchladderid=l.PlayerLadderID) "
		strSQL = strSQL & " FROM tbl_playerLadders l, tbl_games g WHERE g.GameID = l.GameID AND Active > 0 ORDER BY g.GameName, PlayerLaddername"
	Else
		strSQL = "SELECT g.GameName, g.GameID, l.PlayerLadderID, PlayerLaddername, ForFeits = ( select count(PlayerMatchID) from vPlayerMatches where matchawaitingforfeit=1 and left(forfeitreason, 5) <> 'Admin' AND matchladderid=l.PlayerLadderID) "
		strSQL = strSQL & " FROM tbl_playerLadders l, tbl_games g, lnk_pl_a lnk WHERE l.GameID = g.GameID AND lnk.PlayerLadderID = l.PlayerLadderID AND lnk.PlayerID='" & Session("PlayerID") & "' AND Active > 0 ORDER BY GameName, PlayerLaddername"
	End If
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		bgc = bgcone
		do while not ors.EOF
			if intGameID <> oRS.Fields("GameID").Value Then
				If intGameID <> -1 Then
					Response.Write "</TABLE></TD></TR></TABLE><BR><BR>"
				End If
				%>
				<table class="cssBordered" width="100%">
				<TR BGCOLOR="#000000">
					<TH><%=oRs.Fields("GameName").Value%></TH>
				</TR>
			<%
				intGameID = oRs.Fields("GameID").Value
			End If

			Response.Write "<tr bgcolor=" & bgc & ">"
			Response.Write "<td align=LEFT>"
			If oRS.Fields("FOrfeits").Value > 0 Then
				Response.Write "<FONT COLOR=""RED"">" & oRS.Fields("Forfeits").Value & "</FONT>"
			Else
				Response.Write "0"
			End If
			Response.Write " - <a href=adminops.asp?rAdmin=" & aType & "&ladderid=" & ors.Fields("PlayerLadderID").Value & "&laddername=" & server.urlencode(ors.Fields("PlayerLaddername").Value) & ">" & Server.HTMLEncode(ors.Fields("PlayerLadderName").Value) & "</a></td></tr>"
			If bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.Movenext
		loop
		%>
		</TABLE>
		<%
	end if
	oRS.NextRecordset 

	sField = request.querystring("sfield")
	if sfield = "" then
		sField = "PlayerMatchID"
	end if
	laddername = request("laddername")
	ladderid = Request.QueryString("ladderid")
	if ladderid <> "" then
		bgc = bgcone
		%>
		<form name=frmMatchAdmin action=saveItem.asp method=POST>
		<input type=hidden name=SaveType value=PlayeradmKillMatch>
		<input type=hidden name=ladderid value=<%=ladderid%>>
		<input type=hidden name=rAdmin value=<%=aType%>>
		<input type=hidden name=ladderName value="<%=LadderName%>">
				<table class="cssBordered" width="100%">
		<TR BGCOLOR="#000000">
			<TH COLSPAN=6><%=LadderName%></TH>
		</TR>
		<tr bgcolor="#000000">
			<TH><a href=adminops.asp?rAdmin=<%=aType%>&sField=MatchID&ladderid=<%=ladderid%>&laddername=<%=laddername%>>Match ID</a></TH>
			<TH><a href=adminops.asp?rAdmin=<%=aType%>&sField=DefenderName&ladderid=<%=ladderid%>&laddername=<%=laddername%>>Defender</a></TH>
			<TH><a href=adminops.asp?rAdmin=<%=aType%>&sField=AttackerName&ladderid=<%=ladderid%>&laddername=<%=laddername%>>Attacker</a></TH>
			<TH><a href=adminops.asp?rAdmin=<%=aType%>&sField=MatchDate&ladderid=<%=ladderid%>&laddername=<%=laddername%> >Match Date</a></TH>
			<TH>Comms</a></TH>
			<TH>Operation</TH>
		</TR>
		<%
		bgc=bgctwo
		strSQL= "SELECT * FROM vPlayerMatches WHERE MatchLadderId=" & ladderid & xtra & " ORDER BY " & sField
		oRs.Open strSQL, oConn
		
		if not (ors.EOF and ors.BOF) then
			do while not ors.EOF
				mDate=ors.fields("MatchDate").value
				if mdate <> "TBD" then
					newMDate = right(mDate, len(mDate)-instr(1, mDate, ","))
					newMDate = left(newmdate, len(newmdate)-4)
					newmdate=formatdatetime(newmdate, 2)
				else
					newmdate=mdate
				end if
				Response.Write "<tr bgcolor=" & bgc & "><td valign=top>" & ors.Fields("PlayerMatchID").Value & "</td>"
				Response.Write "<td valign=top><a href=viewplayer.asp?player=" & server.urlencode(oRS.Fields("DefenderName").Value) & ">" & oRS.Fields("DefenderName").Value & "</a></td>"
				Response.Write "<td valign=top><a href=viewplayer.asp?player=" & server.urlencode(oRS.Fields("AttackerName").Value) & ">" & oRS.Fields("AttackerName").Value & "</a></td>"
				Response.Write "<td valign=top>" & newMDate & "</td>"
				Response.Write "<td valign=top><a href=""PlayerLadderAdmin.asp?player=" & server.urlencode(oRS.Fields("AttackerName").Value) & "&ladder=" & Server.URLEncode(laddername) & """>Match Comms</a></td>"
				Response.Write "<td valign=top rowspan=2><SELECT NAME=""AdminMatch_" & ors.Fields("PlayerMatchID").Value & """>"
				Response.Write "<OPTION SELECTED VALUE=""NOTHING"">Do Nothing</OPTION>" & vbCrLf 
				Response.Write "<OPTION VALUE=""OVERRIDE"">Admin Override</OPTION>" & vbCrLf 
				Response.Write "<OPTION VALUE=""KILL"">Kill Match</OPTION>" & vbCrLf 
				Response.Write "<OPTION VALUE=""FORFEITD"">Defender Forfeit</OPTION>" & vbCrLf 
				Response.Write "<OPTION VALUE=""FORFEITA"">Attacker Forfeit</OPTION>" & vbCrLf 
				Response.Write "</SELECT></td></tr>"
				if ors.fields("ForFeitReason").value = "xxx" then
					Response.Write "<tr><td colspan=5 valign=top bgcolor=" & bgc & " align=center>&nbsp;</td></tr>"
				else 
					Response.Write "<tr><td colspan=5 valign=top bgcolor=" & bgc & " align=center><font color=#ff0000>" & ors.fields("ForFeitReason").value & "</font></td></tr>"
				end if
				ors.MoveNext
				if bgc=bgcone then
					bgc=bgctwo
				else
					bgc=bgcone

				end if
			loop
		end if
		ors.Close 
		%>
		<tr bgcolor=<%=bgc%>><td colspan=6 align=center><input type=submit name=submit1 value=Submit class=bright></td></tr>
		</TABLE>
		</form>
		<%
	end if
	Call ContentEnd()
end if

if atype = "Player" then
	if not IsSysAdminLevel2() then
		response.clear
		response.redirect "errorpage.asp?error=3"
	end if
	Call ContentStart("Player Administration")
	If Request.QueryString("Player") <> "" Then
		strsql="select playerhandle, tbl_players.playerid from tbl_players WHERE playerHandle like '%" & CheckString(SearchString(request("player"))) & "%' order by playerhandle"
		%>
			<form name=frmPlayerStuff action=saveItem.asp method=post>
				<table class="cssBordered" width="100%">
			<TR BGCOLOR="#000000">
				<TH>Delete Player</TH>
			</TR>
			<TR>
				<TD ALIGN=CENTER BGCOLOR="<%=bgcone%>"><select class=brightred name=player size=4 style='width:200'>
			<%
			ors.open strSQL, oconn
			if not (ors.eof and ors.bof) then
				do while not ors.eof
					%>
					<option value="<%=replace(ors.fields(0).value, """", "&quot;")%>"><%=Server.HTMLEncode(ors.fields(0).value)%>
					<%
					ors.movenext
				loop
			end if
			oRS.NextRecordset 
			%>
			</SELECT>
			</TD></TR>
			<TR BGCOLOR="<%=bgctwo%>">
				<TD ALIGN=CENTER><input type=submit name=submit1 value='Delete Selected Player' class=bright></TD>
			</TR>
			</TABLE>
			<input type=hidden name=SaveType value=DeletePlayer>
			</FORM>
			<BR><BR>
		<%
	Else
		%>
				<table class="cssBordered" width="100%">
			<FORM METHOD="Get" ACTION="adminops.asp" id=form1 name=form1>
			<INPUT TYPE=HIDDEN name="atype" value="Player">
			<TR BGCOLOR="#000000">
				<TH>Search for Player</TH>
			</TR>
			<TR BGCOLOR="<%=bgcone%>">
				<TD ALIGN=CENTER><INPUT TYPE=TEXT STYLE="width:200px" NAME="player" id="player"></TD>
			</TR>
			<TR	BGCOLOR="<%=bgctwo%>">
				<TD ALIGN=CENTER><INPUT TYPE=Submit VALUE="Search"></TD>
			</TR>
			</FORM>
			</TABLE>
		<BR><BR>
		<%
	End if
	
	strSQL = "SELECT p.PlayeriD, p.PlayerHandle from sysadmins s, tbl_players p WHERE p.PlayerID = s.AdminID AND p.PlayerHandle <> 'Triston' AND p.PlayerHandle <> 'Polaris' ORDER BY p.PlayerHandle"
	oRs.Open strSQL, oConn
	if not (ors.EOF and ors.BOF) then
		%>
				<table class="cssBordered" width="100%">
			<form action=saveItem.asp method=post name=FrmDelAdminPlayers>
			<input type=hidden value=DeleteSysadmin name=SaveType>
			<TR BGCOLOR="#000000">
				<TH>System Admins</TH>
			</TR>
			<TR BGCOLOR="<%=bgcOne%>">
				<TD ALIGN=CENTER><select name=SysadminDel size=4 class=brightred style='width:200'>
				<%
				do while not ors.eof
					%>
					<option value="<%=oRS.fields("PlayerID").value%>"><%=Server.HTMLEncode(oRS.fields("PlayerHandle").value)%></option>
					<%
					ors.movenext
				loop
				%>
				</SELECT>
				</TD></TR>
				<tr bgcolor=<%=bgctwo%>>
					<td align=center><input class=bright value='De-Sysadmin Player' name=submit2 type=submit></td></tr>
				</FORM>
		</TABLE><BR><BR>
		<%
	End If
	ors.NextRecordset 
	
	If Request.QueryString("Player") <> "" Then
		strsql="select playerhandle, tbl_players.playerid from tbl_players WHERE playerHandle like '%" & CheckString(SearchString(request("player"))) & "%' order by playerhandle"
		%>
				<table class="cssBordered" width="100%">
			<form name=frmPlayerStuff action=saveItem.asp method=post>
			<TR BGCOLOR="#000000">
				<TH>SysAdmin Player</TH>
			</TR>
			<TR>
				<TD ALIGN=CENTER BGCOLOR="<%=bgcone%>"><select class=brightred name=SysadminAdd size=4 style='width:200'>
			<%
			ors.open strSQL, oconn
			if not (ors.eof and ors.bof) then
				do while not ors.eof
					%>
					<option value="<%=replace(ors.fields(1).Value, """", "&quot;")%>"><%=Server.HTMLEncode(ors.fields(0).value)%>
					<%
					ors.movenext
				loop
			end if
			oRS.NextRecordset 
			%>
			</SELECT>
			</TD></TR>
			<TR BGCOLOR="<%=bgctwo%>">
				<TD ALIGN=CENTER><input type=submit name=submit1 value='SysAdmin Selected Player' class=bright></TD>
			</TR>
			</TABLE>
			<input type=hidden name=SaveType value=AddSysadmin>
			</FORM>
			<BR><BR>
		<%
	Else
		%>
				<table class="cssBordered" width="100%">
			<FORM METHOD="Get" ACTION="adminops.asp" id=form1 name=form1>
			<INPUT TYPE=HIDDEN name="atype" value="Player">
			<TR BGCOLOR="#000000">
				<TH>Search for Player to SysAdmin</TH>
			</TR>
			<TR BGCOLOR="<%=bgcone%>">
				<TD ALIGN=CENTER><INPUT TYPE=TEXT STYLE="width:200px" NAME="player" id="player"></TD>
			</TR>
			<TR	BGCOLOR="<%=bgctwo%>">
				<TD ALIGN=CENTER><INPUT TYPE=Submit VALUE="Search" id=Submit4 name=Submit4></TD>
			</TR>
			</FORM>
			</TABLE>
		<BR><BR>
		<%
	End if
	Call ContentEnd()
end if

if atype="Team" then
	if not IsSysAdminLevel2() then
		response.clear
		response.redirect "errorpage.asp?error=3"
	end if
	Call Content2BoxStart("Team Administration")
	%>
	<table width=780 border="0" cellspacing="0" cellpadding="0" BACKGROUND="">
	<tr>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	<td width=380>
	<%
	If Request.QueryString("Team") <> "" Then 
		%>
				<table class="cssBordered" width="100%">
			<form name=frmTeamStuff action=saveItem.asp method=post>
			<TR BGCOLOR="#000000">
				<TH>Active Teams</TH>
			</TR>
			<%
			strSQL="Select teamid, Teamname from tbl_teams where teamactive=1 AND (TeamName like '%" & CheckString(SearchString(request.querystring("team")))  & "%') order by TeamName "
			ors.open strSQL, oconn
			Response.Write "<tr bgcolor=" & bgctwo & "><td align=center><select name=TeamId class=bright style='width:250' size=15>"
			if not (ors.eof and ors.bof) then
				do while not ors.eof
					Response.Write "<option value=" & ors.fields(0).value & ">" & Server.HTMLEncode(ors.fields(1).value)
					ors.movenext
				loop
			end if
			Response.Write "<tr bgcolor=" & bgcone & " height=25><td align=center><input type=hidden name=SaveType value=DeleteTeam><input type=submit name=submit1 value='Delete Selected Team' class=bright></td></tr>"
			ors.NextRecordset 
			%>
			</FORM>
		</TABLE>
		<%
	Else
		%>
				<table class="cssBordered" width="100%">
			<FORM METHOD="Get" ACTION="adminops.asp" id=form1 name=form1>
			<INPUT TYPE=HIDDEN name="atype" value="Team">
			<TR BGCOLOR="#000000">
				<TH>Search for Team</TH>
			</TR>
			<TR BGCOLOR="<%=bgcone%>">
				<TD ALIGN=CENTER><INPUT TYPE=TEXT STYLE="width:200px" NAME="team" id="team"></TD>
			</TR>
			<TR	BGCOLOR="<%=bgctwo%>">
				<TD ALIGN=CENTER><INPUT TYPE=Submit VALUE="Search" id=Submit4 name=Submit4></TD>
			</TR>
			<%
			if Errorcode=1 then
				response.write "<TR BGCOLOR=""#000000""><TD><B>Cannot delete teams with active challenges, please remove their challenges first.</b></TD></TR>"
			end if	
			%>
			</FORM>
			</TABLE>
		<BR><BR>
		<%
	End If
	%>
	</td>
	<td><img src="/images/spacer.gif" width="10" height="1"></td>
	<td width=379>

				<table class="cssBordered" width="100%">
	<form name=frmTeamDel action=saveItem.asp method=post>
	<TR BGCOLOR="#000000">
		<TH>Deleted Teams</TH>
	</TR>
	<%
	strsql="select teamid, teamname from tbl_teams where teamactive=0 order by teamname"
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		Response.Write "<tr bgcolor=" & bgctwo & " valign=center><td align=center><select name=TeamId class=bright size=15 style='width:250'>"
		do while not ors.eof
			Response.Write "<option value=" & ors.fields(0).value & ">" & Server.HTMLEncode(ors.fields(1).value)
			ors.movenext
		loop
		response.write "</select></td></tr>"
		Response.Write "<tr bgcolor=" & bgcone & "><td align=center><input type=hidden name=SaveType value=RestoreTeam><input type=submit name=submit1 value='Undelete Team' class=bright></td></tr>"
	else	
		response.write "<TR BGCOLOR=" & bgcone & "><TD align=center><b><font color=red>There are no deleted teams in the database</font></b></TD></TR>"
	end if
	ors.NextRecordSet
	%>
	</FORM></TABLE>

	</td>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	</tr>
	</table>
	<%
	Call Content2BoxEnd()
	
end if	

'---------------------------------
' Ladder Admin
'---------------------------------
if aType="Ladder" then
	Call ContentStart("Ladder Administration")
	
	If bSysAdmin Then
		strSQL = "SELECT l.*, g.GameName, g.GameID "
		strSQL = strSQL & " from tbl_ladders l, tbl_games g WHERE g.GameID = l.GameID AND LadderShown = 1 ORDER BY GameName, LadderActive DESC, LadderName"
	Else
		strSQL="Select l.*, g.GameName, g.GameID "
		strSQL = strSQL & " FROM tbl_ladders l, lnk_l_a lnk, tbl_games G "
		strSQL = strSQL & " WHERE lnk.LadderID = l.LadderID AND g.GameID = l.GameID AND  LadderShown = 1 AND lnk.PlayerID ='" & Session("PlayerID") & "' order by GameName, LadderActive DESC, LadderName"
	End If
	ors.open strSQL, oconn
	if not (ors.eof and ors.bof) then
		bgc=bgcone
		do while not ors.eof
			if intGameID <> oRs.Fields("GameID").Value Then
				if intGameID <> -1 Then
					Response.Write "</TABLE></TD></TR></TABLE><BR><BR>"
				End If
				intGameID = oRs.Fields("GameID").Value
				%>
				<table class="cssBordered" width="100%">
				<TR BGCOLOR="#000000">
					<TH <% If bSysAdmin Then Response.Write " COLSPAN=10 " Else Response.Write " COLSPAN=8" End If %>><%=oRS.Fields("GameName").Value%></TH>
				</TR>					
				<TR BGCOLOR="#000000">
					<TH WIDTH=200>Ladder Name</TH>
					<TH WIDTH=50>Active</TH>
					<TH WIDTH=50>Locked</TH>
					<% If bSysAdmin Then %>
					<TH WIDTH=60>Activate</TH>
					<% End If %>
					<TH WIDTH=75>Max Roster</TH>
					<TH WIDTH=75>Min Roster</TH>
					<TH WIDTH=75>Reset Rest</TH>
					<TH WIDTH=75>Map List</TH>
					<TH WIDTH=100>Roster Report</TH>
					<% If bSysAdmin Then %>
					<TH WIDTH=25>Edit</TH>
					<% End If %>
				</TR>
				<%
			End If
				%>
				<tr bgcolor=<%=bgc%>>
				<td><a href="viewladder.asp?ladder=<%=Server.URLEncode(ors.fields("LadderName").value)%>"><%=Server.HTMLEncode(ors.fields("LadderName").value)%></A></td>
				<TD><%
				If oRs.Fields("LadderActive").Value  = 1 Then 
					Response.Write "Yes"
				Else
					Response.Write "<FONT COLOR=""RED"">No</FONT>"
				End If
				%></TD>
				<TD><%
				If oRs.Fields("LadderLocked").Value  = 0 Then 
					Response.Write "No"
				Else
					Response.Write "<FONT COLOR=""RED"">Yes</FONT>"
				End If
				%></TD>
				<% 
				If bSysAdmin Then
					If oRs.Fields("LadderActive").Value = 1 Then
						Response.Write "<TD><A HREF=""saveitem.asp?savetype=HaltLadder&ladder=" & oRS.Fields("LadderID").Value & """>Deactivate</A></TD>"
					Else
						Response.Write "<TD><A HREF=""saveitem.asp?savetype=StartLadder&ladder=" & oRS.Fields("LadderID").Value & """>Activate</A></TD>"
					End If
				End If
				%>
				<TD ALIGN=RIGHT><%
				If oRS.Fields("RosterLimit").Value <> 0 Then
					Response.Write "<FONT COLOR=""RED"">" & oRS.Fields("RosterLimit").Value & "</FONT>"
				Else
					Response.Write "0"
				End If
				%></TD>
				<TD ALIGN=RIGHT><%
				If oRS.Fields("MinPlayer").Value <> 0 Then
					Response.Write oRS.Fields("MinPlayer").Value
				Else
					Response.Write "<FONT COLOR=""RED"">0</FONT>"
				End If
				%></TD>
				<TD><A href="saveitem.asp?SaveType=LadderResetRest&ladderid=<%=oRs.fIelds("LadderID").Value%>&ladder=<%=Server.URLEncode(ors.Fields("LadderName").Value)%>">Reset Rest</A></TD>
				<TD><A href="maplist.asp?ladder=<%=Server.URLEncode(ors.Fields("LadderName").Value)%>">Edit Map List</A></TD>
				<TD><A href="/reports/rosterreport.asp?ladder=<%=Server.URLEncode(ors.Fields("LadderName").Value)%>&numplayer=<%=ors.Fields("MinPlayer").Value%>">View Report</A></TD>
				<% If bSysAdmin Then %>
				<TD><A href="/addladder.asp?IsEdit=true&ladder=<%=server.URLEncode(oRs.Fields("LadderName").Value)%>">Edit</A></TD>
				<% End If %>
			</TR>				
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.movenext
		loop
	end if	
	ors.NextRecordSet		
	%>
	</table>
	<%
	Call ContentEnd()
end if

'---------------------------------
' Ladder Admin
'---------------------------------
if aType="PLadder" then
	Call ContentStart("Player Ladder Administration")
	%>
				<table class="cssBordered" width="100%">
	<TR BGCOLOR="#000000">
		<TH>Ladder Name</TH>
		<TH>Active</TH>
		<TH>Locked</TH>
		<% If bSysAdmin Then %>
		<TH>Activate</TH>
		<% End If %>
		<TH>Map List</TH>
		<% If bSysAdmin Then %>
		<TH>Edit</TH>
		<% End If %>
	</TR>
	<%
	If bSysAdmin Then
		strSQL = "SELECT l.* "
		strSQL = strSQL & " from tbl_playerladders l ORDER BY Active DESC, PlayerLadderName"
	Else
		strSQL="Select l.* "
		strSQL = strSQL & " FROM tbl_playerladders l, lnk_pl_a lnk "
		strSQL = strSQL & " WHERE lnk.PlayerLadderID = l.PlayerLadderID AND lnk.PlayerID ='" & Session("PlayerID") & "' order by Active DESC, PlayerLadderName"
	End If
	ors.open strSQL, oconn
	if not (ors.eof and ors.bof) then
		bgc=bgcone
		do while not ors.eof
			%>
				<tr bgcolor=<%=bgc%>>
				<td><a href="viewplayerladder.asp?ladder=<%=Server.URLEncode(ors.fields("PlayerLadderName").value)%>"><%=Server.HTMLEncode(ors.fields("PlayerLadderName").value)%></A></td>
				<TD><%
				If oRs.Fields("Active").Value  = 1 Then 
					Response.Write "Yes"
				Else
					Response.Write "<FONT COLOR=""RED"">No</FONT>"
				End If
				%></TD>
				<TD><%
				If oRs.Fields("Locked").Value  = 0 Then 
					Response.Write "No"
				Else
					Response.Write "<FONT COLOR=""RED"">Yes</FONT>"
				End If
				%></TD>
				<% 
				If bSysAdmin Then
					If oRs.Fields("Active").Value = 1 Then
						Response.Write "<TD><A HREF=""saveitem.asp?savetype=PHaltLadder&ladder=" & oRS.Fields("PlayerLadderID").Value & """>Deactivate</A></TD>"
					Else
						Response.Write "<TD><A HREF=""saveitem.asp?savetype=PStartLadder&ladder=" & oRS.Fields("PlayerLadderID").Value & """>Activate</A></TD>"
					End If
				End If
				%>
				<TD><A href="maplistplayer.asp?ladder=<%=Server.URLEncode(ors.Fields("PlayerLadderName").Value)%>">Edit Map List</A></TD>
				<% If bSysAdmin Then %>
				<TD><A href="/addplayerladder.asp?IsEdit=true&name=<%=server.URLEncode(oRs.Fields("PlayerLadderName").Value)%>">Edit</A></TD>
				<% End If %>
			</TR>				
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.movenext
		loop
	end if	
	ors.NextRecordSet		
	%>
	</table>
	<%
	Call ContentEnd()
	
end if
%>

<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>