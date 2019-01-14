<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Security Audit"

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

If Not(IsSysAdminLevel2()) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Security Audit") %>
<table width=760 border=0 cellspacing=0 cellpadding=0 align=center BGCOLOR="#444444">
<TR><TD>
<table width=100% border=0 cellspacing=1 cellpadding=2 align=center>
<tr BGCOLOR="#000000">
	<TH><a href="securityaudit.asp?sort=type">Type</a></th>
	<!-- th>ID</th -->
	<TH><a href="securityaudit.asp?sort=name">Name</a></th>
	<TH>Active</th>
	<!-- TH>Player ID</th -->
	<TH><a href="securityaudit.asp?sort=playerhandle">Player Name</a></th>
	<TH><a href="securityaudit.asp?sort=playeremail">Email</a></th>
</tr>
	<%
	strSQL = "SELECT data.*, tbl_players.playerhandle, tbl_players.playeremail FROM " & _ 
		" (" & _
		" SELECT 'LadderAdmin' AS 'Type', lnk_l_a.LadderID AS 'ID', LadderName AS 'Name', LadderActive AS 'Active', PlayerID FROM lnk_l_a" & _
		"	LEFT OUTER JOIN tbl_ladders ON tbl_ladders.ladderID = lnk_l_a.ladderid" & _
		" UNION " & _
		" SELECT 'PlayerLadderAdmin', lnk_pl_a.PlayerLadderID, PlayerLadderName, Active, PlayerID FROM lnk_pl_a" & _
		" LEFT OUTER JOIN tbl_playerladders ON tbl_playerladders.PlayerladderID = lnk_pl_a.Playerladderid" & _
		" UNION " & _
		" SELECT 'EloLadderAdmin', lnk_elo_admin.EloLadderID, EloLadderName, EloActive, PlayerID FROM lnk_elo_admin " & _
		" LEFT OUTER JOIN tbl_elo_ladders ON tbl_elo_ladders.eloladderid = lnk_elo_admin.eloladderid" & _
		" UNION " & _
		" SELECT 'LeagueAdmin', lnk_league_admin.LeagueID, LeagueName, LeagueActive, PlayerID FROM lnk_league_admin " & _
		"	LEFT OUTER JOIN tbl_leagues ON tbl_leagues.leagueid = lnk_league_admin.leagueid" & _
		" UNION " & _
		" SELECT 'TournamentAdmin', lnk_m_a.TournamentID, TournamentName, Active, PlayerID FROM lnk_m_a" & _
		"	LEFT OUTER JOIN tbl_tournaments ON tbl_tournaments.TournamentID = lnk_m_a.TournamentID" & _
		" UNION " & _
		" SELECT 'Sysadmin', '', CAST(AdminLevel AS Varchar(20)), 1, AdminID FROM sysadmins" & _
		" ) AS data" & _
		" LEFT OUTER JOIN tbl_players ON data.playerid = tbl_players.playerid"
	If ( Len(Request("sort") ) > 0  ) Then
		Select Case Request("sort")
			Case "type"
				strSQL = strSQL & " ORDER BY Type, Name, playerhandle"
			Case "name"
				strSQL = strSQL & " ORDER BY Name, Type, playerhandle"
			Case "playerhandle"
				strSQL = strSQL & " ORDER BY playerhandle, Type, Name, Active"
			Case "playeremail"
				strSQL = strSQL & " ORDER BY PlayerEmail, playerhandle, Type, Name"
			Case Else
				strSQL = strSQL & " ORDER BY type, active asc, name, data.PlayerID"
		End Select
	Else
			strSQL = strSQL & " ORDER BY type, active asc, name, data.PlayerID"
	End If
	oRS.Open strSQL, oConn
	bgc = bgctwo
	if not (ors.eof and ors.bof) then
		do while not ors.EOF 
			Response.write "<tr bgcolor=" & bgcone & ">"
			Response.write "<td>" & Server.HTMLEncode( oRs.Fields("Type").Value & "") & "</td>"
			Response.write "<td>" & Server.HTMLEncode( oRs.Fields("Name").Value & "") & "</td>"
			' Response.write "<td>" & Server.HTMLEncode( oRs.Fields("ID").Value & "") & "</td>"
			Response.write "<td>" & Server.HTMLEncode( oRs.Fields("Active").Value & "") & "</td>"
			' Response.write "<td>" & Server.HTMLEncode( oRs.Fields("PlayerID").Value & "") & "</td>"
			Response.write "<td><a href=""/viewplayer.asp?player=" & Server.URLEncode(ors.Fields("PlayerHandle").Value & "") & """>" & Server.HTMLEncode(ors.Fields("PlayerHandle").Value & "") & "</A></TD>"
			Response.write "<td>" & Server.HTMLEncode( oRs.Fields("PlayerEmail").Value & "") & "</td>"
			ors.MoveNext 
			if bgc = bgcone then
				bgc = bgctwo
			else
				bgc = bgcone
			end if 
		loop
	end if
	ors.close
	%>
</table></TD></TR>
</TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

