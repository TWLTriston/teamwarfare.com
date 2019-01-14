<%
''''''''''''''''''''''''''''''''''''''''''
' Security Functions
''''''''''''''''''''''''''''''''''''''''''
Function CheckCookie()
' REDACTED '
End Function



Function IsSysAdmin()
	If Len(Session("SysAdmin")) = 0 Then
		CheckCookie()
		Dim sSQL, objRSs
		Set objRSs = Server.CreateObject("ADODB.RecordSet")
		Session("SysAdmin") = False
		IsSysAdmin = False
		sSQL = "SELECT AdminLevel FROM SysAdmins WHERE AdminID = '" & Session("PlayerID") & "'"
		objRSs.Open sSQL, oConn
		If Not(objRSs.EOF AND objRSs.BOF) Then
			IsSysAdmin = True
			Session("SysAdmin") = True
			Session("SysAdminLevel") = objRSs.Fields("AdminLevel").Value
		End If
		objRSs.Close
		Set objRSs = Nothing
	Else
		IsSysAdmin = Session("SysAdmin")
	End If
End Function

Function IsSysAdminLevel2()
	If Session("SysAdminLevel") = "2" Then
		IsSysAdminLevel2 = True
	Else
		IsSysAdminLevel2 = False
	End If
End Function

Function IsAnyLadderAdmin()
	If Len(Session("PlayerID")) = 0 Then
		IsAnyLadderAdmin = False
	ElseIf Len(Session("AnyLadderAdmin")) = 0 Then
		Dim sSQL, objRSs
		Set objRSs = Server.CreateObject("ADODB.RecordSet")

		CheckCookie()
		IsAnyLadderAdmin = False
		Session("AnyLadderAdmin") = False

		sSQL = "SELECT 'AdminID' = LALinkID from lnk_l_a lnk WHERE lnk.PlayerID = '" & Session("PlayeriD") & "'"
		sSQL = sSQL & " UNION "
		sSQL = sSQL & " SELECT 'AdminID' =  PLAdminID from lnk_pl_a lnk2 WHERE lnk2.PlayerID = '" & Session("PlayeriD") & "'"
		sSQL = sSQL & " UNION "
		sSQL = sSQL & " SELECT 'AdminID' =  PlayerID from lnk_elo_admin lnk3 WHERE lnk3.PlayerID = '" & Session("PlayeriD") & "'"
		sSQL = sSQL & " UNION "
		sSQL = sSQL & " SELECT 'AdminID' =  PlayerID from lnk_league_admin lnk4 WHERE lnk4.PlayerID = '" & Session("PlayeriD") & "'"
		objRSs.open sSQL, oConn
		If Not(objRSs.BOF AND objRSs.EOF) Then
			IsAnyLadderAdmin = True
			Session("AnyLadderAdmin") = True
		End If
		objRSs.Close
		Set objRSs = Nothing
	Else
		IsAnyLadderAdmin = Session("AnyLadderAdmin")
	End If
End Function

Function IsAnyTeamLadderAdmin()
	If Len(Session("AnyTeamLadderAdmin")) = 0 Then
		Dim sSQL, objRSs
		Set objRSs = Server.CreateObject("ADODB.RecordSet")

		CheckCookie()
		IsAnyTeamLadderAdmin = False
		Session("IsAnyTeamLadderAdmin") = False

		sSQL = "SELECT 'AdminID' = LALinkID from lnk_l_a lnk WHERE lnk.PlayerID = '" & Session("PlayeriD") & "'"
		objRSs.open sSQL, oConn
		If Not(objRSs.BOF AND objRSs.EOF) Then
			IsAnyTeamLadderAdmin = True
			Session("AnyTeamLadderAdmin") = True
		End If
		objRSs.Close
		Set objRSs = Nothing
	Else
		IsAnyTeamLadderAdmin = Session("AnyTeamLadderAdmin")
	End If
End Function

' 01/30/04 Support for /admin/default.asp	-jkb
Function IsAnyPlayerLadderAdmin()
	If Len(Session("AnyPlayerLadderAdmin")) = 0 Then
		Dim sSQL, objRSs
		Set objRSs = Server.CreateObject("ADODB.RecordSet")

		CheckCookie()
		IsAnyPlayerLadderAdmin = False
		Session("IsAnyPlayerLadderAdmin") = False

		sSQL = "SELECT 'AdminID' = PLAdminID from lnk_pl_a lnk WHERE lnk.PlayerID = '" & Session("PlayeriD") & "'"
		objRSs.open sSQL, oConn
		If Not(objRSs.BOF AND objRSs.EOF) Then
			IsAnyPlayerLadderAdmin = True
			Session("IsAnyPlayerLadderAdmin") = True
		End If
		objRSs.Close
		Set objRSs = Nothing
	Else
		IsAnyPlayerLadderAdmin = Session("IsAnyPlayerLadderAdmin")
	End If
End Function

' 02/06/04 Support for /admin/default.asp	-jkb
Function IsAnyLeagueAdmin()
	If Len(Session("IsAnyLeagueAdmin")) = 0 Then
		Dim sSQL, objRSs
		Set objRSs = Server.CreateObject("ADODB.RecordSet")

		CheckCookie()
		IsAnyLeagueAdmin = False
		Session("IsAnyLeagueAdmin") = False

		sSQL = "SELECT LeagueAdminID FROM lnk_league_admin WHERE PlayerID = '" & Session("PlayeriD") & "'"
		objRSs.open sSQL, oConn
		If Not(objRSs.BOF AND objRSs.EOF) Then
			IsAnyLeagueAdmin = True
			Session("IsAnyLeagueAdmin") = True
		End If
		objRSs.Close
		Set objRSs = Nothing
	Else
		IsAnyLeagueAdmin = Session("IsAnyLeagueAdmin")
	End If
End Function

' 02/06/04 Support for /admin/default.asp	-jkb
Function IsAnyTournamentAdmin()
	If Len(Session("IsAnyTournamentAdmin")) = 0 Then
		Dim sSQL, objRSs
		Set objRSs = Server.CreateObject("ADODB.RecordSet")

		CheckCookie()
		IsAnyTournamentAdmin = False
		Session("IsAnyTournamentAdmin") = False

		sSQL = "SELECT MALinkID FROM lnk_m_a WHERE PlayerID = '" & Session("PlayeriD") & "'"
		objRSs.open sSQL, oConn
		If Not(objRSs.BOF AND objRSs.EOF) Then
			IsAnyTournamentAdmin = True
			Session("IsAnyTournamentAdmin") = True
		End If
		objRSs.Close
		Set objRSs = Nothing
	Else
		IsAnyTournamentAdmin = Session("IsAnyTournamentAdmin")
	End If
End Function

' 02/06/04 Support for /admin/default.asp	-jkb
Function IsAnyScrimLadderAdmin()
	If Len(Session("IsAnyScrimLadderAdmin")) = 0 Then
		Dim sSQL, objRSs
		Set objRSs = Server.CreateObject("ADODB.RecordSet")

		CheckCookie()
		IsAnyScrimLadderAdmin = False
		Session("IsAnyScrimLadderAdmin") = False

		sSQL = "SELECT EloLadderAdminID from lnk_elo_admin WHERE PlayerID = '" & Session("PlayeriD") & "'"
		objRSs.open sSQL, oConn
		If Not(objRSs.BOF AND objRSs.EOF) Then
			IsAnyScrimLadderAdmin = True
			Session("IsAnyScrimLadderAdmin") = True
		End If
		objRSs.Close
		Set objRSs = Nothing
	Else
		IsAnyScrimLadderAdmin = Session("IsAnyScrimLadderAdmin")
	End If
End Function

Function IsAnyServerAdmin()
	If Len(Session("AnyServerAdmin")) = 0 Then
		Dim sSQL, objRSs
		Set objRSs = Server.CreateObject("ADODB.RecordSet")
		IsAnyServerAdmin = False
		Session("AnyServerAdmin") = False
		CheckCookie()
		sSQL = "SELECT sp.SPLinkID FROM lnk_s_p sp WHERE sp.PlayerID ='" & Session("PlayerID") & "'"
		objRSs.open sSQL, oconn
		If Not(objRSs.BOF And objRSs.eof) Then
			IsAnyServerAdmin = True
			Session("AnyServerAdmin") = True
		End if
		objRSs.Close
		Set objRSs = Nothing
	Else
		IsAnyServerAdmin = Session("AnyServerAdmin")
	End If
End Function

Function IsTeamFounder(byVal sTeamName)
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsTeamFounder = False
	CheckCookie()
	sSQL = "SELECT TeamID, TeamFounderID "
	sSQL = sSQL & " FROM tbl_teams WHERE "
	sSQL = sSQL & " TeamFounderID = '" & Session("PlayerID") & "' "
	sSQL = sSQL & " AND TeamName='" & CheckString(sTeamName) & "'"
	objRSs.Open sSQL, oConn
	If Not(objRSs.EOF AND objRSs.BOF) Then
		IsTeamFounder = True
	End If
	objRSs.Close
	Set objRSs = Nothing
End Function

Function IsTeamCaptainByID(byVal iTeamID, byVal iLadderID)
	Dim sSQL, objRSs
	Dim TLLinkID
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsTeamCaptainByID = False
	CheckCookie()
	sSQL = "SELECT TLLinkID FROM lnk_T_L WHERE LadderID='" & iLadderID & "' AND TeamID='" & iTeamID & "'"
	objRSs.Open sSQL, oConn
	If Not(objRSs.EOF AND objRSs.BOF) Then
		TLLinkID = objRSs.Fields("TLLinkID").Value
		objRSs.NextRecordSet

		sSQL ="SELECT IsAdmin from lnk_T_P_L where TLLinkID='" & TLLinkID & "' and PlayerID='" & Session("PlayerID") & "'"
		objRSs.Open sSQL, oConn
		If Not(objRSs.EOF AND objRSs.BOF) Then
			if objRSs.fields(0).value = 1 Then
				IsTeamCaptainByID=true
			End If
		end if
		objRSs.close
	Else
		objRSs.close
	End If
	Set objRSs = Nothing
End Function

Function IsEloTeamCaptainByID(byVal iTeamID, byVal iLadderID)
	Dim sSQL, objRSs
	Dim intEloTeamID
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsEloTeamCaptainByID = False
	CheckCookie()
	sSQL = "SELECT lnkEloTeamID FROM lnk_elo_team WHERE EloLadderID='" & iLadderID & "' AND TeamID='" & iTeamID & "'"
	objRSs.Open sSQL, oConn
	If Not(objRSs.EOF AND objRSs.BOF) Then
		intEloTeamID = objRSs.Fields("lnkEloTeamID").Value
		objRSs.NextRecordSet

		sSQL ="SELECT IsAdmin from lnk_elo_team_player where lnkEloTeamID='" & intEloTeamID & "' and PlayerID='" & Session("PlayerID") & "'"
		objRSs.Open sSQL, oConn
		If Not(objRSs.EOF AND objRSs.BOF) Then
			if objRSs.fields(0).value = 1 Then
				IsEloTeamCaptainByID=true
			End If
		end if
		objRSs.close
	Else
		objRSs.close
	End If
	Set objRSs = Nothing
End Function

Function IsEloLadderAdminByID(byVal iLadderID)
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsEloLadderAdminByID = False
	CheckCookie()
	sSQL="SELECT * from lnk_elo_admin WHERE ELoLadderid = '" & iLadderID & "' AND Playerid='" & Session("PlayerID") & "'"
	objRSs.Open sSQL, oConn
	If Not (objRSs.bof AND objRSs.eof) Then
		IsEloLadderAdminByID=True
	End If
	objRSs.close
	Set objRSs = Nothing
End Function

Function IsLadderAdminByID(byVal iLadderID)
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsLadderAdminByID = False
	CheckCookie()
	sSQL="SELECT * from lnk_L_A WHERE Ladderid = '" & iLadderID & "' AND playerid='" & Session("PlayerID") & "'"
	objRSs.Open sSQL, oConn
	If Not (objRSs.bof AND objRSs.eof) Then
		IsLadderAdminByID=True
	End If
	objRSs.close
	Set objRSs = Nothing
End Function

Function IsLeagueAdminByID(byVal iLeagueID)
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsLeagueAdminByID = False
	CheckCookie()
	sSQL="select LeagueAdminID from lnk_league_admin where LeagueID='" & iLeagueID & "' and playerid='" & Session("PlayerID") & "'"
	objRSs.open sSQL, oConn
	if not(objRSs.eof and objRSs.bof) then
		IsLeagueAdminByID=True
	end if
	objRSs.Close
	Set objRSs = Nothing
End Function

Function IsLeagueTeamCaptainByID(byVal iTeamID, byVal iLeagueID)
	Dim sSQL, objRSs, lnkLeagueTeamID
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsLeagueTeamCaptainByID = False
	CheckCookie()
	sSQL="select lnkLeagueTeamID from lnk_league_team where LeagueID='" & iLeagueID & "' and TeamID='" & iTeamID & "'"
	objRSs.open sSQL, oConn
	if not(objRSs.eof and objRSs.bof) then
		lnkLeagueTeamID = objRSs.fields(0).value
		objRSs.close
		sSQL="select isadmin from lnk_league_team_player where lnkLeagueTeamID='" & lnkLeagueTeamID & "' and PlayerID='" & Session("PlayerID") & "'"
		objRSs.open sSQL, oConn
		if not(objRSs.eof and objRSs.bof) then
			if objRSs.fields(0).value = 1 then
				IsLeagueTeamCaptainByID=true
			end if
		end if
		objRSs.Close
	else
		objRSs.Close
	end if
	Set objRSs = Nothing
End Function

Function IsLeagueTeamCaptainByLinkID(byVal iLinkID)
	Dim sSQL, objRSs, lnkLeagueTeamID
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsLeagueTeamCaptainByLinkID = False
	CheckCookie()
	sSQL="select isadmin from lnk_league_team_player where lnkLeagueTeamID='" & iLinkID & "' and PlayerID='" & Session("PlayerID") & "'"
'	Response.Write sSQL
	objRSs.open sSQL, oConn
	if not(objRSs.eof and objRSs.bof) then
		if objRSs.fields(0).value = 1 then
			IsLeagueTeamCaptainByLinkID=true
		end if
	end if
	objRSs.Close
	Set objRSs = Nothing
End Function

Function IsEloTeamCaptainByLinkID(byVal iLinkID)
	Dim sSQL, objRSs, intEloTeamID
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsEloTeamCaptainByLinkID = False
	CheckCookie()
	sSQL="select isadmin from lnk_elo_team_player where lnkEloTeamID='" & iLinkID & "' and PlayerID='" & Session("PlayerID") & "'"
'	Response.Write sSQL
	objRSs.open sSQL, oConn
	if not(objRSs.eof and objRSs.bof) then
		if objRSs.fields(0).value = 1 then
			IsEloTeamCaptainByLinkID = true
		end if
	end if
	objRSs.Close
	Set objRSs = Nothing
End Function

Function IsTournamentAdminByID(byVal iTournamentID)
	Dim sSQL, objRSs, TGLinkID
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsTournamentAdminByID = False
	CheckCookie()
	sSQL="select * from lnk_M_A where TournamentID='" & iTournamentID & "' and playerid='" & Session("PlayerID") & "'"
	objRSs.open sSQL, oConn
	if not(objRSs.eof and objRSs.bof) then
		IsTournamentAdminByID=True
	end if
	objRSs.close
	Set objRSs = Nothing
End Function

Function IsTournamentTeamCaptainByID(byVal iTeamID, byVal iTournamentID)
	Dim sSQL, objRSs, tmlinkid
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsTournamentTeamCaptainByID = False
	CheckCookie()
	sSQL="select TMLinkID from lnk_T_m where TournamentID=" & iTournamentID & " and TeamID='" & iTeamID & "'"
	objRSs.open sSQL, oConn
	if not (objRSs.eof and objRSs.bof) then
		tmlinkid=objRSs.fields(0).value
		objRSs.close
		sSQL="select isadmin from lnk_T_M_P where TMLinkID='" & TMLinkID & "' and PlayerID='" & Session("PlayerID") & "'"
		objRSs.open sSQL, oConn
		if not (objRSs.eof and objRSs.bof) then
			if objRSs.fields(0).value = 1 then
				IsTournamentTeamCaptainByID=true
			end if
		end if
		objRSs.close
	else
		objRSs.close
	end if
	Set objRSs = Nothing
End Function

Function IsPlayerLadderAdminByID(byVal iLadderID)
	Dim sSQL, objRSs, tmlinkid
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	CheckCookie()
	sSQL="select * from lnk_pL_A where PlayerLadderid='" & iLadderID & "' and Playerid='" & Session("PlayerID") & "'"
	objRSs.open sSQL, oConn
	if not(objRSs.eof and objRSs.bof) then
		IsPlayerLadderAdminByID=True
	end if
	objRSs.close
	Set objRSs = Nothing
End Function

Function IsTeamCaptain(byVal sTeamName, byVal sLadderName)
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsTeamCaptain = False
	CheckCookie()
	sSQL = "select IsAdmin from lnk_t_p_l lp, tbl_ladders l, tbl_players p, tbl_teams t, lnk_t_l lt where l.laddername='" & CheckString(sLadderName) & "' and l.ladderid = lt.ladderid "
	sSQL = sSQL & "AND lt.TLLinkID = lp.TLLinkID AND lp.playerid = '" & Session("PlayerID") & "' AND t.teamid = lt.teamid AND t.teamname='" & CheckString(sTeamName) & "'"
	objRSs.open sSQL, oConn
	if not (objRSs.bof and objRSs.eof) then
		if objRSs.fields("IsAdmin").value = 1 then
			IsTeamCaptain=true
		end if
	end if
	objRSs.close
	Set objRSs = Nothing
End Function

Function IsEloTeamCaptain(byVal sTeamName, byVal sLadderName)
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsEloTeamCaptain = False
	CheckCookie()
	sSQL = "select IsAdmin from lnk_elo_team_player lp, tbl_elo_ladders l, tbl_players p, tbl_teams t, lnk_elo_team lt where l.Eloladdername='" & CheckString(sLadderName) & "' and l.Eloladderid = lt.Eloladderid "
	sSQL = sSQL & "AND lt.lnkEloTeamID = lp.lnkEloTeamID AND lp.playerid = '" & Session("PlayerID") & "' AND t.teamid = lt.teamid AND t.teamname='" & CheckString(sTeamName) & "'"
	objRSs.open sSQL, oConn
	if not (objRSs.bof and objRSs.eof) then
		if objRSs.fields("IsAdmin").value = 1 then
			IsEloTeamCaptain=true
		end if
	end if
	objRSs.close
	Set objRSs = Nothing
End Function

Function IsLadderAdmin(byVal sLadderName)
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsLadderAdmin = False
	CheckCookie()
	sSQL = "SELECT lnk.LadderID FROM lnk_l_a lnk, tbl_ladders l WHERE l.LadderID = lnk.LadderID AND lnk.PlayerID = '" & Session("PlayerID") & "' AND l.LadderName = '" & CheckString(sLadderName) & "'"
	objRSs.Open sSQL, oConn
	if not(objRSs.bof and objRSs.eof) then
		IsLadderAdmin=True
	end if
	objRSs.Close
	Set objRSs = Nothing
End Function

Function IsEloLadderAdmin(byVal sLadderName)
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsEloLadderAdmin = False
	CheckCookie()
	sSQL = "SELECT lnk.EloLadderID FROM lnk_elo_admin lnk, tbl_elo_ladders l WHERE l.EloLadderID = lnk.EloLadderID AND lnk.PlayerID = '" & Session("PlayerID") & "' AND l.EloLadderName = '" & CheckString(sLadderName) & "'"
	objRSs.Open sSQL, oConn
	if not(objRSs.bof and objRSs.eof) then
		IsEloLadderAdmin=True
	end if
	objRSs.Close
	Set objRSs = Nothing
End Function

Function IsPlayerLadderAdmin(byVal sLadderName)
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsPlayerLadderAdmin = False
	CheckCookie()
	sSQL = "SELECT lnk.* FROM lnk_pl_a lnk, tbl_playerladders pl WHERE lnk.PlayerLadderID = pl.PlayerLadderID AND pl.Playerladdername='" & CheckString(sLadderName) & "' AND lnk.PlayerID='" & Session("PlayerID") & "'"
	objRSs.Open sSQL, oConn
	if not(objRSs.bof and objRSs.eof) then
		IsPlayerLadderAdmin=True
	end if
	objRSs.Close
	Set objRSs = Nothing
End Function

Function IsAnyTeamFounder()
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	IsAnyTeamFounder = False
	CheckCookie()
	sSQL = "SELECT TOP 1 teamfounderid FROM tbl_teams where TeamFounderID='" & Session("PlayerID") & "'"
	objRSs.Open sSQL, oConn
	if not(objRSs.bof and objRSs.eof) then
		IsAnyTeamFounder=true
	end if
	objRSs.Close
	Set objRSs = Nothing
end function


Function IsLeagueAdmin(LeagueName)
	Dim sec, sec2, membername, playeridseccheck, lid
	Set Sec = Server.Createobject("ADODB.RecordSet")
	Set Sec2 = Server.Createobject("ADODB.RecordSet")
	IsLeagueAdmin = False
	CheckCookie()
	membername=session("uName")
	strsql="select PlayerID from tbl_players where PlayerHandle='" & replace(membername, "'", "''") & "'"
	sec.open strsql, oconn
	if not (sec.bof and sec.eof) then
		playeridseccheck=sec.fields(0).value
		strsql="select LeagueID from tbl_leagues where LeagueName='" & replace(LeagueName, "'", "''") & "'"
		sec.close
		sec.open strsql, oconn
		if not(sec.bof and sec.eof) then
			lid = sec.fields(0).value
			strsql="select LeagueAdminID from lnk_league_admin where LeagueID=" & lid & " and playerid=" & playeridseccheck
			sec.close
			sec.open strsql, oconn
			if not(sec.eof and sec.bof) then
				IsLeagueAdmin=True
			end if
			sec.close
		else
			sec.close
		end if
	else
		sec.close
	end if
End Function

Function IsTournamentAdmin(TournamentName)
	Dim sec, sec2, membername
	Set Sec = Server.Createobject("ADODB.RecordSet")
	Set Sec2 = Server.Createobject("ADODB.RecordSet")
	IsTournamentAdmin = False
	CheckCookie()
	strsql = "SELECT p.PlayerID FROM lnk_m_a m INNER JOIN tbl_tournaments t ON t.TournamentID = m.TournamentID "
	strsql = strSQL & " INNER JOIN tbl_players p ON p.PlayerID = m.PlayerID WHERE p.PlayerHandle= '" & CheckString(session("uName")) & "' AND t.TournamentName = '" & CheckString(TournamentName) & "'"
	sec.open strsql, oconn
	if not (sec.bof and sec.eof) then
		IsTournamentAdmin=True
	end if
	sec.close
End Function

Function IsTeamFounderByID(TeamID)
	Dim sec, sec2
	Set Sec = Server.Createobject("ADODB.RecordSet")
	IsTeamFounderByID = False
	CheckCookie()

	strsql="select TeamFounderID from tbl_teams where TeamID='" & TeamID & "' AND TeamFounderID = '" & Session("PlayerID") & "'"
	sec.open strsql,oconn
	if not (sec.eof and sec.bof) then
		IsTeamFounderByID=True
	End If
	sec.close
	Set Sec = Nothing
	' look up memberID and Team ID
	' check against tbl_teams
End Function

Function IsLeagueTeamCaptain(TeamName, LeagueName)
	Dim sec
	Set Sec = Server.Createobject("ADODB.RecordSet")
	IsLeagueTeamCaptain = False
	CheckCookie()
	strSQL = "SELECT lnkLeagueTeamPlayerID FROM lnk_league_team_player ltp "
	strSQL = strSQL & " INNER JOIN lnk_league_team lt ON lt.lnkLeagueTeamID = ltp.lnkLeagueTeamID "
	strSQL = strSQL & " INNER JOIN tbl_leagues l ON l.LeagueID = lt.LeagueID  "
	strSQL = strSQL & " INNER JOIN tbl_teams t ON t.TeamID = lt.TeamID "
	strSQL = strSQL & " WHERE ltp.PlayerID = '" & Session("PlayerID") & "' "
	strSQL = strSQL & " AND LeagueName = '" & CheckString(LeagueName) &  "' "
	strSQL = strSQL & " AND ltp.IsAdmin = 1 "
	strSQL = strSQL & " AND TeamName = '" & CheckString(TeamName) & "' "
	sec.open strsql,oconn
	if not (sec.bof and sec.eof) then
		IsLeagueTeamCaptain = True
	end if
	sec.close
	Set sec = Nothing
End Function

Function IsTournamentTeamCaptain(TeamName, TournamentName)
	Dim sec, sec2
	Set Sec = Server.Createobject("ADODB.RecordSet")
	Set Sec2 = Server.Createobject("ADODB.RecordSet")
	IsTournamentTeamCaptain = False
	CheckCookie()
	membername=session("uName")
	strsql="select PlayerID from tbl_players where PlayerHandle='" & replace(membername, "'", "''") & "'"
	sec.open strsql,oconn
	if not (sec.bof and sec.eof) then
		playeridseccheck=sec.fields(0).value
		sec.close
		strsql="select TeamId from tbl_teams where TeamName='" & replace(teamname, "'", "''") & "'"
		sec.open strsql,oconn
		if not(sec.eof and sec.bof) then
			tid=sec.fields(0).value
			sec.close
			strsql="select TournamentID from tbl_tournaments where TournamentName='" & replace(TournamentName, "'", "''") & "'"
			sec.open strsql,oconn
			if not (sec.eof and sec.bof) then
				tourid=sec.fields(0).value
				sec.close
				strsql="select TMLinkID from lnk_T_M where TournamentID=" & tourid & " and TeamID=" & tid
				sec.open strsql,oconn
				if not (sec.eof and sec.bof) then
					tmlinkid=sec.fields(0).value
					sec.close
					strsql="select isadmin from lnk_T_M_P where TMLinkID=" & TMLinkID & " and PlayerID=" & playeridseccheck
					sec.open strsql,oconn
					if not (sec.eof and sec.bof) then
						if sec.fields(0).value = 1 then
							IsTournamentTeamCaptain=true
						end if
					end if
					sec.close
				else
					sec.close
				end if
			else
				sec.close
			end if
		else
			sec.close
		end if
	else
		sec.close
	end if
End Function

%>
