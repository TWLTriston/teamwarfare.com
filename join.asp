<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Join New Competition"

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

Dim jType, strTeamName, intTeamID, full, intGameID 
strTeamName = Request.QueryString ("team")
If Not(bSysAdmin OR IsTeamFounder(strTeamName)) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Set oRS2 = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("") %>
From the four options below, decide whether you are joining a ladder, <br />
league or tournament. Then select which you wish to join by choosing it <br />
in the select box and clicking join.
<br /><br />
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
	<TR><TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH="400">
	<form name=frmLadderJoin action=saveItem.asp method=post>
	<tr>
		<th bgcolor="#000000">Join A Rung Based Ladder With <%=Server.HTMLEncode(strTeamName)%></th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="center"><select name="LadderToJoin" id="LadderToJoin">
	<%
	strSQL = "SELECT TeamID FROM tbl_teams WHERE TeamName='" & CheckString(strTeamName) & "'"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		intTeamID = oRS.Fields("TeamID").Value
	End If
	oRS.NextRecordset
	
	strSQL = "SELECT tbl_games.GameName, tbl_games.GameID, tbl_games.GameAbbreviation,    "
  	strSQL = strSQL & " tbl_ladders.LadderName, "
  	strSQL = strSQL & " tbl_ladders.LadderID,   "
  	strSQL = strSQL & " tbl_ladders.LadderAbbreviation   "
 	strSQL = strSQL & " FROM tbl_games, tbl_ladders   "
 	strSQL = strSQL & " WHERE tbl_games.GameID = tbl_ladders.GameID   "
  	strSQL = strSQL & " AND LadderActive = 1   "
	strSQL = strSQL & " AND LadderID NOT IN (SELECT LadderID FROM lnk_t_l WHERE TeamID = " & intTeamID & " AND IsActive = 1) "
	strSQL = strSQL & "  ORDER BY GameName ASC, LadderName ASC   "
	ors.Open strSQL, oConn
	bgc=bgcone
	intGameID = -1
	if not (ors.EOF and ors.BOF) then
		do while not ors.EOF
			%>
			<option value="<%=Server.HTMLEncode(ors.Fields("LadderName").Value)%>"><%=Server.HTMLEncode(ors.Fields("LadderName").Value)%></option>
			<%
			ors.MoveNext
		loop
	end if	
	ors.NextRecordset 
	%>
	</select></td>
	</tr>
	<tr>
		<th bgcolor="#000000"><INPUT id=submit1 name=submit1 type=submit value='Join Ladder' class=bright ></th>
	</tr>
	</table>
	<input id=TeamID name=TeamID type=hidden value="<%=intTeamID%>">
	<input id=SaveType name=SaveType type=hidden value=LadderJoin>
	</form>
	</TD></TR>
	</TABLE>
<br /><br />
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
	<TR><TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH="400">
	<form name=frmEloLadderJoin action="/scrim/saveItem.asp" method=post>
	<tr>
		<th bgcolor="#000000">Join A Power Rating Ladder With <%=Server.HTMLEncode(strTeamName)%></th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="center"><select name="LadderToJoin" id="LadderToJoin">
	<%
	strSQL = "SELECT tbl_games.GameName, tbl_games.GameID, tbl_games.GameAbbreviation,    "
  	strSQL = strSQL & " tbl_elo_ladders.EloLadderName, "
  	strSQL = strSQL & " tbl_elo_ladders.EloLadderID,   "
  	strSQL = strSQL & " tbl_elo_ladders.EloAbbreviation "
 	strSQL = strSQL & " FROM tbl_games, tbl_Elo_ladders   "
 	strSQL = strSQL & " WHERE tbl_games.GameID = tbl_elo_ladders.EloGameID   "
  	strSQL = strSQL & " AND EloActive = 1   "
	strSQL = strSQL & " AND EloLadderID NOT IN (SELECT EloLadderID FROM lnk_elo_team WHERE TeamID = " & intTeamID & " AND Active = 1) "
	strSQL = strSQL & "  ORDER BY GameName ASC, EloLadderName ASC   "
	ors.Open strSQL, oConn
	bgc=bgcone
	intGameID = -1
	if not (ors.EOF and ors.BOF) then
		do while not ors.EOF
			%>
			<option value="<%=Server.HTMLEncode(ors.Fields("EloLadderName").Value)%>"><%=Server.HTMLEncode(ors.Fields("EloLadderName").Value)%></option>
			<%
			ors.MoveNext
		loop
	Else
		Response.write "<option value="""">No scirm ladders available at this time.</option>"
	End If	
	ors.NextRecordset 
	%>
	</select></td>
	</tr>
	<tr>
		<th bgcolor="#000000"><INPUT id=submit1 name=submit1 type=submit value='Join Ladder' class=bright ></th>
	</tr>
	</table>
	<input id=TeamID name=TeamID type=hidden value="<%=intTeamID%>">
	<input id=SaveType name=SaveType type=hidden value="EloLadderJoin">
	</form>
	</TD></TR>
	</TABLE>

	<br /><br />
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
	<TR><TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH="400">
	<form name=frmLeagueJoin action=/tournament/savetournament.asp method=post>
	<TR>
		<Th BGCOLOR="#000000">Join a Tournament with <%=Server.HTMLEncode(strteamName)%></TH>
	</TR>
	<tr>
		<td bgcolor="<%=bgcone%>" align="center">
	<select name="TournamentToJoin" id="TournamentToJoin">
	<%
	strSQL = "SELECT Distinct tournamentName, tournamentID, Divisions, TeamsPerDiv FROM tbl_tournaments t WHERE SignUp = 1 AND active = 1 AND LanID = 0 "
	strSQL = strSQL & " AND tournamentID NOT IN (SELECT tournamentID FROM lnk_t_m lnk WHERE lnk.tournamentID = t.tournamentID AND TeamID='" & intTeamID & "')"
	strSQL = strSQL & " ORDER BY tournamentname"
	'response.write strSQL
	ors.Open strSQL, oConn
	bgc=bgcone
	Dim intTournaments
	intTournaments = 0
	if not (ors.EOF and ors.BOF) then
		do while not ors.EOF
			strsql = "select count(*) from lnk_T_M where tournamentID = " & ors.fields(1).value
'			response.write strSQL
			ors2.open strsql, oconn
			if ors2.fields(0).value = (ors.fields(2).value * ors.fields(3).value) then
				full = true
			else
				full = false
			end if
			ors2.nextrecordset
			if not full then 
				intTournaments = intTournaments +1
				%>
				<option value="<%= Server.htmlencode(ors.Fields(1).Value) %>"><%= Server.htmlencode(ors.Fields(0).Value) %></option>
				<%
			end if
			ors.MoveNext
		loop
		if intTournaments = 0 Then
			Response.write "<option value="""">No tournaments available at this time.</option>"
		End If
	Else
		Response.write "<option value="""">No tournaments available at this time.</option>"
	End If	
	oRS.NextRecordset 
	%>
	</select></td>
	</tr>
	<%
	if intTournaments > 0 Then
		%>
		<tr>
			<th bgcolor="#000000"><input type="submit" value="Join Tournament" class="bright" /></th>
		</tr>
		<%
	End if
	%>
	</table>
	<input id=hidden name=TeamID type=hidden value=<%=intTeamID%>>
	<input id=hidden name=TeamName type=hidden value="<%=strTeamName%>">
	<input id=hidden name=SaveType type=hidden value=TournamentJoin>
	</form>
	</TD></TR>
	</TABLE>
	<br /><br />
<form name="frmJoinLeague" id="frmJoinTournament" action="saveitem.asp" method="post">
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
<tr>
	<td>
	<table border="0" cellspacing="1" cellpadding="4" width="400">
	<tr>
		<th bgcolor="#000000" colspan="2">Join a League with <%=server.htmlencode(strteamName)%></th>
	</tr>
	<tr>
		<td bgcolor=<%=bgcone%> align="Center"><select name="LeagueConferenceID" id="LeagueConferenceID">
	<%
	strSQL = "SELECT LeagueName, l.LeagueID, LeagueInviteOnly, ConferenceName, LeagueConferenceID FROM tbl_leagues l "
	strSQL = strSQL & " INNER JOIN tbl_League_conferences c "
	strSQL = strSQL & " ON c.LeagueID = l.LeagueID"
	strSQL = strSQL & " WHERE LeagueActive = 1 AND SignUp = 1 "
	strSQL = strSQL & " AND l.LeagueID NOT IN (SELECT LeagueID FROM lnk_league_team WHERE TeamID = '" & intTeamID & "' AND Active=1)"
	strSQL = strSQL & " ORDER BY LeagueName, ConferenceName"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		Do While Not(oRs.EOF)
			%>
			<option value="<%=oRs.Fields("LeagueConferenceID").Value%>"><%=Server.HTMLEncode(oRs.Fields("LeagueName").Value & "")%> League &raquo; <%=Server.HTMLEncode(oRs.Fields("ConferenceName").Value & "")%> Conference</option>
			<%
			oRs.MoveNext
		Loop
		%>
		</select></td>
		</tr>
		<tr>
			<th bgcolor="#000000"><input type="submit" value="Join League"></th>
		</tr>
		<%
	Else
		%>
		<option value="">No Open Leagues</option></select></td></tr>
		<%	
	End If
	oRs.Close
	%>
</table>
</td></tr></table>	
	<input id=hidden name=TeamID type=hidden value="<%=intTeamID%>">
	<input id=hidden name=TeamName type=hidden value="<%=Server.HTMLEncode(strTeamName & "")%>">
	<input id=hidden name=SaveType type=hidden value="LeagueJoin">
</form>
		
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>
