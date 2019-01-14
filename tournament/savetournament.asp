<%' Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: SaveItem"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<body bgcolor="#000000" text="#FFFFFF">
<%
LoggedIn= Session("LoggedIn")
SysAdmin = IsSysAdmin()

if not(LoggedIn) and request.form("SaveType") <> "player" then
	' Require login to perform action.
	Response.Clear
	Response.Redirect "errorpage.asp?error=2"
end if

'Response.Write "<font color=#ffffff>Querystring: " & Request.QueryString
'Response.Write "<br>Form data: " & Request.Form & "</font></br>"

'-----------------------------------------------
' Tournament Team Join
'-----------------------------------------------
if Request.Form("SaveType")="TournamentJoin" then
	If Request("TournamentToJoin") <> "" Then
		tourID=Request("TournamentToJoin")
	Else
		Response.Clear
		Response.Redirect "/errorpage.asp?error=9"
	End If

	ClearToJoin = False
	strsql = "select TeamFounderID from tbl_teams where teamid='" & request("TeamID") & "'"
	ors.Open strsql, oconn
	if not (ors.EOF and ors.BOF) then
		FounderID = ors.Fields(0).Value
	end if
	ors.Close
	strsql = "select TournamentID from lnk_t_M inner join lnk_t_m_p on lnk_t_m_P.TMLinkID = lnk_t_M.TMLinkID "
	strsql = strsql & "where lnk_t_m_p.PlayerID = '" & FounderID & "' and lnk_t_m.TournamentID='" & tourid & "'"
	ors.Open strsql, oconn
	if ors.EOF and ors.BOF then
		ClearToJoin = True
	end if
	ors.nextrecordset
	if clearToJoin then
		strSQL = "select Top 1 TeamsInDiv = (SELECT count(distinct lnk.teamid) FROM lnk_t_m lnk "
		strSQL = strSQL & " WHERE lnk.DivisionID = div.DivisionID "
		strSQL = strSQL & " AND lnk.TournamentID = div.TournamentID), div.divisionID, tbl_tournaments.TournamentName, Maps "
		strsql = strsql & "from tbl_tdivisions div, tbl_tournaments "
		strsql = strsql & "where div.tournamentID = '" & tourid & "'" 
		strsql = strsql & "AND tbl_tournaments.tournamentid = div.tournamentID "
		strsql = strsql & "Group BY div.DivisionID, TournamentName, Maps, div.tournamentid "
		strsql = strsql & "order by TeamsInDiv asc, DivisionID ASC "
		ors.Open strsql, oconn
		If Not(orS.EOF AND oRS.BOF) Then
			Maps = ors.fields("Maps").Value 
			TournamentName = ors.Fields("TournamentName").Value 
			assigneddivision = ors.Fields("DivisionID").Value
			TeamsInDivision = ors.Fields("TeamsInDiv").Value + 1
		End If
		ors.NextRecordset 
		strsql = "insert into lnk_t_m (TeamID, TournamentID, DivisionID) values ('" & Request.Form("TeamID") & "', "
		strsql = strsql & "'" & TourID & "', '" & AssignedDivision & "')"
		oconn.execute (strsql)
		strsql = "select TMLinkID from lnk_t_m where TeamID='" & Request.Form("TeamID") & "' AND TournamentID = '" & tourid & "'"
		ors.Open strsql, oconn
		if not(ors.EOF and ors.BOF) then
			TMLinkID = ors.Fields(0).Value
		end if
		ors.NextRecordset
		if TeamsInDivision mod 2 = 1 then
			FirstPosition = True
		else
			FirstPosition = False
		end if
		
		SeedOrder = round((TeamsInDivision/2)-0.1,0)

		strSQL = "insert into lnk_t_m_p (TMLinkID, PlayerID, IsAdmin, DateJoined) values ('" & TMLinkID & "', '" & founderID & "', '1', '" & now() & "')"
		oconn.execute strSQL
		blnEmptySlot = false
		strSQL = "SELECT TOP 1 RoundsID FROM tbl_rounds WHERE [Round] = 1 AND DivisionID = '" & AssignedDivision & "' AND TournamentID = '" & tourID & "' AND Team1ID = 0 "
		oRs.Open strSQL, oConn
		If Not(ors.EOF AND oRs.BOF) Then
			strSQL = "UPDATE tbl_rounds SET Team1ID = '" & TMLinkID & "' WHERE RoundsID = '" & oRs.FieldS("RoundsID").Value & "'"
			oConn.Execute(strSQL)
		Else
			ors.NextRecordSet
			strSQL = "SELECT TOP 1 RoundsID FROM tbl_rounds WHERE [Round] = 1 AND DivisionID = '" & AssignedDivision & "' AND TournamentID = '" & tourID & "' AND Team2ID = 0 "
			oRs.Open strSQL, oConn
			If Not(ors.EOF AND oRs.BOF) Then
				strSQL = "UPDATE tbl_rounds SET Team2ID = '" & TMLinkID & "' WHERE RoundsID = '" & oRs.FieldS("RoundsID").Value & "'"
				oConn.Execute(strSQL)
			Else
				ors.NextRecordSet
				if FirstPosition then
					strsql = "insert into tbl_rounds (SeedOrder, Team1ID, Team2ID, TournamentID, DivisionID, Round) values "
					strsql = strsql & "('" & SeedOrder & "', '" & TMLinkID & "', '0', '" & tourid & "', '" & AssignedDivision & "', '1')"
					oconn.execute (strsql)
					strsql = "select RoundsID from tbl_rounds where Team1ID='" & TMLinkID & "'"
					ors.Open strsql, oconn
					if not(ors.EOF and ors.BOF) then
						RoundsID = ors.Fields(0).Value
					end if
					ors.nextrecordset
					strsql = "select TMapID, Map, MapOrder from tbl_map_tour where TournamentID = '" & TourID & "' order by MapOrder"
					ors.Open strsql, oconn
					if not(ors.eof and ors.BOF) then
						do while not(ors.EOF)
							strsql = "insert into lnk_r_m (RoundsID, Map, Maporder, TMapID) values ('" & roundsID 
							strsql = strsql & "', '" & ors.Fields(1).value & "', '" & ors.Fields(2).Value & "', '" & ors.Fields(0).Value & "')"
							oconn.Execute strsql
							ors.MoveNext
						loop
					end if
					ors.NextRecordset
					
				else
					strsql = "update tbl_rounds SET Team2ID = '" & TMLinkID & "' where "
					strsql = strsql & "TournamentID='" & TourID & "' AND Round = '1' AND SeedOrder='" 
					strsql = strsql & (SeedOrder - 1) & "' AND DivisionID='" & assigneddivision & "'"
					oconn.execute (strsql)
					strsql = "select RoundsID from tbl_rounds where Team2ID = '" & TMLinkID & "' AND TournamentID = " & TourID
					ors.open strsql, oconn
					if not(ors.eof and ors.bof) then
						RoundsID = ors.fields(0).value
					end if
					ors.nextrecordset
				end if		
			End If
		End If
	else
		Response.Clear
		Response.Redirect "/errorpage.asp?error=9"
	end if
	Response.Clear
	Response.Redirect "/tournament/default.asp?tournament=" & server.urlencode(TournamentName) & "&page=brackets&div=" & assigneddivision
end if
'-----------------------------------------------
' Kick Player from Roster
'-----------------------------------------------
if request.form("savetype")="DropPlayer" then
	
	TMLinkiD=request.form("link")
	playerid=request.form("playerid")
	strsql="select TournamentID, teamid from lnk_t_m where TMLinkID=" & TMLinkID
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		tid = ors.fields(1).value
		Tourid = ors.fields(0).value
	end if
	ors.close
	if not(IsSysAdmin()) and not(IsTeamFounderByID(tid)) and not(IsTournamentAdminByID(Tourid)) and not(IsTournamentTeamCaptainByID(tid, Tourid)) then
		response.clear
		response.redirect "errorpage.asp?error=3"
	else
		if tid <> "" and Tourid <> "" then
			strsql="select teamname from tbl_teams where teamid=" & tid
			ors.open strsql, oconn
			tname=ors.fields(0).value
			ors.close
			strsql="select TournamentName from tbl_tournaments where tournamentid=" & Tourid
			ors.open strsql, oconn
			mname=ors.fields(0).value
			ors.close
		end if
		if playerid <> "" then
			strsql="delete from lnk_t_m_p where tmlinkid=" & tmlinkid & " and playerid=" & playerid
			ors.open strsql, oconn
		end if
'		response.write htmlencode(tname) & htmlencode(mname)
		response.redirect "/teamtournamentadmin.asp?team=" & server.urlencode(tname) & "&tournament=" & server.urlencode(mname)
	end if
end if
'-----------------------------------------------
' Promote to Captain
'-----------------------------------------------
if Request.Form("SaveType") = "PromoteCaptain" then
	if not(IsSysAdmin()) and not(IsTournamentTeamCaptain(request.form("team"), request.form("tournament"))) and not(IsTeamFounder(request.form("team"))) and not(IsTournamentAdmin(request.form("Tournament"))) then
		response.clear
		response.redirect "errorpage.asp?error=3"
	end if
	
	strSQL="UPDATE lnk_T_M_P set isadmin=1 where tMPlinkid=" & Request.Form("playerlist")
	Response.write strSQL
	ors.open strSQL, oconn
	Response.Clear
	Response.Redirect "/teamtournamentadmin.asp?tournament=" & server.urlencode(Request.form("tournament")) & "&team=" & server.urlencode(Request.Form("team"))
end if
'-----------------------------------------------
' Demote Captain
'-----------------------------------------------
if Request.Form("SaveType") = "DemoteCaptain" then
	if not(IsSysAdmin()) and not(IsTournamentTeamCaptain(request.form("team"), request.form("tournament"))) and not(IsTeamFounder(request.form("team"))) and not(IsTournamentAdmin(request.form("Tournament"))) then
		response.clear
		response.redirect "errorpage.asp?error=3"
	end if
	
	strSQL="select lnk_T_M.TMLinkID, playerid from lnk_T_M inner join lnk_T_M_P on lnk_T_M_P.TMLinkID=lnk_T_M.TMLinkID where lnk_T_M_P.TMPLinkID=" & Request.Form("playerlist")
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		plid=ors.Fields(0).Value
		pid=ors.Fields(1).Value 
	end if
	ors.Close
	strSQL="select teamfounderid from tbl_teams inner join lnk_T_M on lnk_T_M.teamid=tbl_teams.teamid where lnk_T_M.TMLinkID=" & plid
	ors.open strsql, oconn
	if not (ors.EOF and ors.BOF) then
		tfid=ors.Fields(0).Value
	end if
	ors.Close 
	if tfid <> pid then
		strSQL="UPDATE lnk_T_M_P set isadmin=0 where tMPlinkid=" & Request.Form("playerlist")
		ors.open strSQL, oconn
	end if
	Response.Clear
	Response.Redirect "/teamtournamentadmin.asp?tournament=" & server.urlencode(Request.form("tournament")) & "&team=" & server.urlencode(Request.Form("team"))
end if
'-----------------------------------------------
' Map saving
'-----------------------------------------------
if Request.Form("SaveType") = "MapSave" then
	if not(IsSysAdmin()) and not(IsTournamentTeamCaptain(request.form("team"), request.form("tournament"))) and not(IsTeamFounder(request.form("team"))) and not(IsTournamentAdmin(request.form("Tournament"))) then
		response.clear
		response.redirect "errorpage.asp?error=3"
	end if
	roundsid = Request.Form("roundsid")
	for i = 1 to 15
		if Request.Form("Map" & i) <> "" then
			strsql = "update lnk_r_m set Map='" & Request.Form("Map" & i) & "' where RoundsID=" & roundsID & " and maporder=" & i
			oconn.execute strsql
			Response.Write strsql
		end if
	next
	Response.Clear
	Response.Redirect "/teamtournamentadmin.asp?tournament=" & server.urlencode(Request.form("tournament")) & "&team=" & server.urlencode(Request.Form("team"))
	
end if
'-----------------------------------------------
' Add Comms
'-----------------------------------------------
if Request.form("SaveType") = "Add_Communications" then
	strSQL= "select commid from tbl_round_comm order by commid"
	ors.open strSQL, oconn
	strSQL="INSERT INTO tbl_round_comm ( RoundsID, CommDate, CommAuthor, Comms, CommTime ) values ('" 
	strSQL = strSQL & CheckString(Request.form("roundsid")) & "','" & Checkstring(Request.Form("commdate")) & "','"  & CheckString(Request.Form("commauthor")) & "','"
	strSQL = strSQL & replace(Request.Form("comms"), "'", "''") & "','" & Request.Form("commtime") & "')"
	ors.close
	oRs.Open strSQL, oConn
set ors = nothing
set oConn = nothing	
set ors2 = nothing	
	Response.Clear 
	Response.Redirect "/TeamTournamentAdmin.asp?tournament=" & server.urlencode(session("CurrentTournament")) & "&team=" & server.urlencode(session("CurrentTeam"))
end if
'-----------------------------------------------
' Edit Comms
'-----------------------------------------------
if Request.form("SaveType") = "Edit_Communications" then
	if not(IsSysAdmin()) then
		response.clear
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL= "update tbl_round_comm set Comms='" & replace(Request.Form("comms"), "'", "''") &  "', CommDate='" & date & "', CommTime='" & time & "' where commid=" & Request.Form("commid")
	ors.open strSQL, oconn
	set ors = nothing
	set oConn = nothing	
	set ors2 = nothing	
	Response.Clear 
	Response.Redirect "/TeamTournamentAdmin.asp?tournament=" & server.urlencode(session("CurrentTournament")) & "&team=" & server.urlencode(session("CurrentTeam"))
end if
'-----------------------------------------------
' Delete Comms
'-----------------------------------------------
if Request.QueryString("SaveType") = "Delete_Communications" then
	if not(IsSysAdmin()) then
		response.clear
		response.redirect "errorpage.asp?error=3"
	end if
	if Request.QueryString("commid") <> "" then
		strSQL= "delete from tbl_round_comm where commid=" & Request.QueryString("commid")
		Response.Write strSQl
		oRs.Open strSQL, oConn
	end if
set ors = nothing
set oConn = nothing	
set ors2 = nothing	
	Response.Clear 
	Response.Redirect "/TeamTournamentAdmin.asp?tournament=" & server.urlencode(session("CurrentTournament")) & "&team=" & server.urlencode(session("CurrentTeam"))
end if
'-----------------------------------------------
' Report Loss
'-----------------------------------------------
if Request.Form("SaveType") = "ReportLoss" then
	RoundsID = Request.Form("RoundsID")
	TournamentID = Request.Form("TournamentID")
	LinkID = Request.Form("LinkID")
	
	strsql = "select * from tbl_rounds where roundsid = " & roundsID
	ors.open strsql, oconn
	
	if not(ors.eof and ors.bof) then
		SeedOrder = ors.fields("SeedOrder").value
		Team1ID = ors.fields("Team1ID").value
		Team2ID = ors.fields("Team2ID").value
		
		if cint(Team1ID) = cint(LinkID) then
			Team1Loser = true
		elseif cint(Team2ID) = cint(LinkID) then
			Team1Loser = false		
		end if
		If Request.Form("AWin") = "Yes" Then
			Team1Loser = Not(Team1Loser)
		End If
		DivisionID = ors.fields("DivisionID").value
		OtherRound = ors.fields("Round").value
		NextRound = OtherRound + 1
	end if
	ors.nextrecordset
	
	strSQL = "SELECT * FROM tbl_tournaments WHERE TournamentID = '" & CheckString(TournamentID) & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intTeamsPerDiv = oRs.Fields("TeamsPerDiv").Value
		chrTournamentStyle = oRs.Fields("TournamentStyle").Value
		chrFinalsStyle = oRs.Fields("FinalsStyle").Value
	End If
	oRs.NextRecordSet
	
	If cStr(DivisionID) = "0" Then 
		chrStyle = chrFinalsStyle
	Else 
		chrStyle = chrTournamentStyle
	End If
	

	if Team1Loser then
		strsql = "update tbl_rounds set WinnerID = " & Team2ID & " where RoundsID = " & roundsID & ";"
		If chrStyle = "S" Then
			strsql = strsql & "update lnk_t_m SET Active = 0 WHERE TMLinkID = '" & Team1ID & "'"
		End If
		WinnerID = Team2ID
		LoserID = Team1ID
	else
		strsql = "update tbl_rounds set WinnerID = " & Team1ID & " where RoundsID = " & roundsID & ";"
		If chrStyle = "S" Then
			strsql = strsql & "update lnk_t_m SET Active = 0 WHERE TMLinkID = '" & Team2ID & "'"
		End If
		WinnerID = Team1ID
		LoserID = Team2ID
	end if
	oconn.execute (strSQL)

'	Ptavv's Seed Determination Method ('cause Triston's doesn't work) 6/6/01
'	In changing this I changed all of the FirstPosition determination and Update
'	checking too, it's more refined now (heh, that's kind of an oxymoron)

	If SeedOrder mod 2 = 0 then
		NewSeed = SeedOrder / 2
		FirstPosition = true
		Team1ID = WinnerID
		Team2ID = 0
	Elseif SeedOrder mod 2 = 1 then
		NewSeed = round((SeedOrder/2)-0.1,0)
		FirstPosition = false
		Team1ID = 0
		Team2ID = WinnerID
	End If
	

	If Log2(intTeamsPerDiv) <= int(OtherRound) Then
		intNextDiv = 0 
		If Log2(intTeamsPerDiv) = OtherRound Then
			NextSeedOrder = DivisionID - 1
			If NextSeedOrder mod 2 = 0 then
				NewSeed = NextSeedOrder / 2
				FirstPosition = true
				Team1ID = WinnerID
				Team2ID = 0
			Elseif NextSeedOrder mod 2 = 1 then
				NewSeed = round((NextSeedOrder/2)-0.1,0)
				FirstPosition = false
				Team1ID = 0
				Team2ID = WinnerID
			End If
		'	Response.Write "Entering outside division... "
		End If
	'	Response.Write "Next Seed: " & NewSeed
	Else
		intNextDiv = DivisionID
	End If

	If chrStyle = "D" Then
		intLoserSeedOrder = -1
		intLoserRound = -1
		blnLoserFirstPosition = True
		If cStr(SeedOrder) = "0" AND cStr(OtherRound) = "1" Then
			intLoserSeedOrder = 4
			intLoserRound = 1
			blnLoserFirstPosition = True
		ElseIf cStr(SeedOrder) = "1" AND cStr(OtherRound) = "1" Then
			intLoserSeedOrder = 4
			intLoserRound = 1
			blnLoserFirstPosition = False
		ElseIf cStr(SeedOrder) = "2" AND cStr(OtherRound) = "1" Then
			intLoserSeedOrder = 5
			intLoserRound = 1
			blnLoserFirstPosition = True
		ElseIf cStr(SeedOrder) = "3" AND cStr(OtherRound) = "1" Then
			intLoserSeedOrder = 5
			intLoserRound = 1
			blnLoserFirstPosition = False
		ElseIf cStr(SeedOrder) = "0" AND cStr(OtherRound) = "2" Then
			intLoserSeedOrder = 3
			intLoserRound = 2
			blnLoserFirstPosition = True
		ElseIf cStr(SeedOrder) = "1" AND cStr(OtherRound) = "2" Then
			intLoserSeedOrder = 2
			intLoserRound = 2
			blnLoserFirstPosition = False
			
		ElseIf cStr(SeedOrder) = "0" AND cStr(OtherRound) = "3" Then
			intLoserSeedOrder = 1
			intLoserRound = 4
			blnLoserFirstPosition = True
			NewSeed = 0
			intNextDiv = DivisionID
			FirstPosition = True
			Team1ID = WinnerID
			Team2ID = 0
		ElseIf cStr(SeedOrder) = "0" AND cStr(OtherRound) = "4" Then
			If Team1Loser Then
				intLoserSeedOrder = 0
				intLoserRound = 5
				blnLoserFirstPosition = False
				
				' Force the winner to stay in the current bracket
				NewSeed = 0
				intNextDiv = DivisionID
				FirstPosition = True
				Team1ID = WinnerID
				Team2ID = LoserID
			Else
				NextRound = 4
			End If	
		ElseIf cStr(SeedOrder) = "0" AND cStr(OtherRound) = "5" Then
			SeedOrder = DivisionID - 1
			If SeedOrder mod 2 = 0 then
				NewSeed = SeedOrder / 2
				FirstPosition = true
				Team1ID = WinnerID
				Team2ID = 0
			Elseif SeedOrder mod 2 = 1 then
				NewSeed = round((SeedOrder/2)-0.1,0)
				FirstPosition = false
				Team1ID = 0
				Team2ID = WinnerID
			End If
			NextRound = 4

		ElseIf cStr(SeedOrder) = "4" AND cStr(OtherRound) = "1" Then
			FirstPosition = False
		ElseIf cStr(SeedOrder) = "5" AND cStr(OtherRound) = "1" Then
			FirstPosition = False
			NewSeed = 3
		ElseIf cStr(SeedOrder) = "2" AND cStr(OtherRound) = "2" Then
			FirstPosition = False
		ElseIf cStr(SeedOrder) = "3" AND cStr(OtherRound) = "2" Then
			FirstPosition = False
		ElseIf cStr(SeedOrder) = "1" AND cStr(OtherRound) = "3" Then
			FirstPosition = False
			NewSeed = 1
			intNextDiv = DivisionID
			Team2ID = WinnerID

		ElseIf cStr(SeedOrder) = "1" AND cStr(OtherRound) = "4" Then
			intNextDiv = DivisionID
			NewSeed = 0
			NextRound = 4
			FirstPosition = False
		End If
		'Response.Write cStr(SeedOrder) & "-" & cStr(OtherRound) & "-" & NewSeed
		'Response.End
		If intLoserSeedOrder = -1 Then
			' Deactive the loser.
			strSQL = "UPDATE lnk_t_m SET Active = 0 WHERE TMLinkID = '" & LoserID & "'"
			oConn.Execute(strSQL)
		Else
			'' This will fill the loser bracket from the winner bracket
			strsql = "SELECT RoundsID from tbl_rounds WHERE SeedOrder=" & intLoserSeedOrder & " AND Round=" & intLoserRound 
			strsql = strsql & " AND DivisionID = " & DivisionID & " AND TournamentID = " & TournamentID
			ors.Open strsql, oconn
			If not(ors.eof and ors.bof) then
				LoserUpdate = True
				LoserRoundsID = oRs.Fields("roundsID").Value
			End If
			oRs.NextRecordSet
			If LoserUpdate Then
				If blnLoserFirstPosition Then
					strsql = "update tbl_rounds set Team1ID = '" & LoserID & "' where RoundsID = " & LoserRoundsID
				Else
					strsql = "update tbl_rounds set Team2ID = '" & LoserID & "' where RoundsID = " & LoserRoundsID
				End If
				oConn.Execute(strSQL)
			Else
				If blnLoserFirstPosition Then
						strsql = "insert into tbl_rounds (Round, Team1ID, Team2ID, WinnerID, TournamentID, DivisionID, SeedOrder) values " &_
							"('" & (intLoserRound) & "', '" & LoserID & "', 0, '0', '" & TournamentID & "', '" & DivisionID & "', '" & intLoserSeedOrder & "')"
				Else
						strsql = "insert into tbl_rounds (Round, Team1ID, Team2ID, WinnerID, TournamentID, DivisionID, SeedOrder) values " &_
							"('" & (intLoserRound) & "', 0, '" & LoserID & "', '0', '" & TournamentID & "', '" & DivisionID & "', '" & intLoserSeedOrder & "')"
				End If
				oConn.Execute(strSQL)
			End If
		End If
	End If			
		strsql = "SELECT RoundsID from tbl_rounds WHERE SeedOrder=" & NewSeed & " AND Round=" & NextRound 
		strsql = strsql & " AND DivisionID = " & intNextDiv & " AND TournamentID = " & TournamentID
		ors.Open strsql, oconn
	
		if not(ors.eof and ors.bof) then
			Update = true
			NextRoundID = ors.fields(0).value
			'response.write "Updating: " & NextRoundID
		else
			Update = false
		'	Response.write "Inserting"
		end if
		ors.nextrecordset
		
		if Update then
			if FirstPosition then
				strsql = "update tbl_rounds set Team1ID = '" & WinnerID & "' where RoundsID = " & NextRoundID
			else
				strsql = "update tbl_rounds set Team2ID = '" & WinnerID & "' where RoundsID = " & NextRoundID
			end if
			'response.write strSQL
			'response.end
			oconn.execute (strsql)
		else
			strsql = "insert into tbl_rounds (Round, Team1ID, Team2ID, WinnerID, TournamentID, DivisionID, SeedOrder) values " &_
					"('" & (NextRound) & "', '" & Team1ID & "', '" & Team2ID & "', '0', '" & TournamentID & "', '" & intNextDiv & "', '" & NewSeed & "')"
			oconn.execute (Strsql)
		end if
	%>
	<script language="javascript">
		window.opener.location.href='<%=request.form("fromurl")%>';
		window.close();
	</script>
	<%	
	Response.End
end if
If Request.Form("SaveType") = "DivisionNames" Then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	strSQL = ""
	For i = 1 to Request.Form("hdnDivisionID").Count
		strSQL = strSQL & "UPDATE tbl_tdivisions SET DivisionName = '" & Request.Form("txtDivisionName")(i) & "' WHERE TDivisionID = '" & Request.Form("hdnDivisionID")(i) & "';"
	Next
	oConn.Execute(strSQL)
	Response.Clear
	Response.Redirect("admintournament.asp")
End If
If Request.QueryString("SaveType") = "OpenSignup" Then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	strSQL = "UPDATE tbl_tournaments SET SignUp = 1 WHERE TournamentID = '" & Request.QuerySTring("TournamentID") & "'"
	oConn.Execute(strSQL)
	Response.Clear
	Response.Redirect("admintournament.asp")
End If	

If Request.QueryString("SaveType") = "CloseSignup" Then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	strSQL = "UPDATE tbl_tournaments SET SignUp = 0 WHERE TournamentID = '" & Request.QuerySTring("TournamentID") & "'"
	oConn.Execute(strSQL)
	Response.Clear
	Response.Redirect("admintournament.asp")
End If	

If Request.QueryString("SaveType") = "Lock" Then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	strSQL = "UPDATE tbl_tournaments SET Locked = 1 WHERE TournamentID = '" & Request.QuerySTring("TournamentID") & "'"
	oConn.Execute(strSQL)
	Response.Clear
	Response.Redirect("admintournament.asp")
End If	

If Request.QueryString("SaveType") = "Unlock" Then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	strSQL = "UPDATE tbl_tournaments SET Locked = 0 WHERE TournamentID = '" & Request.QuerySTring("TournamentID") & "'"
	oConn.Execute(strSQL)
	Response.Clear
	Response.Redirect("admintournament.asp")
End If	

If Request.Form("SaveType") = "EditSeedings" Then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	DivisionID = Request.Form("DivisionID")
	TournamentID = Request.Form("TournamentID")
	strSQL = "DELETE FROM tbl_rounds WHERE TournamentID = '" & TournamentID & "' AND DivisionID = '" & DivisionID & "';"
	For i = 1 to Request.Form("seedorder").Count
		strSQL = strSQL & "INSERT INTO tbl_rounds (Round, Team1ID, Team2ID, WinnerID, TournamentID, DivisionID, SeedOrder) VALUES ("
		strSQL = strSQL & "1, " & Request.Form("Team1")(i) & ", " & Request.Form("Team2")(i) & ", 0, " & TournamentID & ", " & DivisionID & ", " & Request.Form("seedOrder")(i) & ");"
	Next
	oConn.Execute(strSQL)
	Response.Clear
	Response.Redirect "admintournament.asp"
	
End If

If Request.QueryString("SaveType") = "Deactivate" Then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	strSQL = "UPDATE tbl_tournaments SET Active = 0 WHERE TournamentID = '" & Request.QuerySTring("TournamentID") & "'"
	oConn.Execute(strSQL)
	Response.Clear
	Response.Redirect("admintournament.asp")
	
End If

If Request.Form("SaveType") = "EditTournament" Then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	strSQL = "UPDATE tbl_tournaments SET "
	strSQL = strSQL & " TournamentName  = '" & CheckString(Request.Form("TournamentName")) & "'"
	strSQL = strSQL & ", ForumID  = '" & CheckString(Request.Form("ForumID")) & "'"
	strSQL = strSQL & ", GameID  = '" & CheckString(Request.Form("GameID")) & "'"  
	strSQL = strSQL & ", Active  = '" & CheckString(Request.Form("Active")) & "'"  
	strSQL = strSQL & ", Locked  = '" & CheckString(Request.Form("Locked")) & "'" 
	strSQL = strSQL & ", Signup   = '" & CheckString(Request.Form("Signup")) & "'"
	strSQL = strSQL & ", RulesName  = '" & CheckString(Request.Form("RulesName")) & "'"
	strSQL = strSQL & ", HasPrizes  = '" & CheckString(Request.Form("HasPrizes")) & "'"
	strSQL = strSQL & ", HasSponsors  = '" & CheckString(Request.Form("HasSponsors")) & "'"
	strSQL = strSQL & ", RosterLock  = '" & CheckString(Request.Form("RosterLock")) & "'"   
	strSQL = strSQL & ", HeaderURL  = '" & CheckString(Request.Form("HeaderURL")) & "'"
	strSQL = strSQL & " WHERE TournamentID = '" & CheckString(Request.Form("TournamentID")) & "'"
	oConn.Execute(strSQL)
	Response.Redirect("AdminTournament.asp")

End If

If Request.Form("SaveType") = "EditContent" Then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	strSQL = "UPDATE tbl_tournaments SET Content" & CheckString(Request.Form("ContentName")) & " = '" & CheckString(Request.Form("Content")) & "' WHERE TournamentID = '" & CheckString(Request.Form("TournamentID")) & "'"
	oConn.Execute(strSQL)

	Response.Redirect("EditContent.asp?Tournament=" & Server.URLEncode(Request.Form("TournamentName")))
End If

'-----------------------------------------------
' LeagueAssignAdmin
'-----------------------------------------------
if Request.Form("SaveType") = "TournamentAssignAdmin" then
	if not(bSysAdmin) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	If Len(Request.Form("PlayerID")) > 0 Then
		strSQL = "INSERT INTO lnk_m_a (PlayerID, TournamentID) VALUES ('" & Request.Form("PlayerID") & "', '" & Request.Form("LeagueID") & "')"
		oConn.Execute(strSQL)		
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "AssignAdmin.asp"
end if
'-----------------------------------------------
' LeagueRemoveAdmin
'-----------------------------------------------
if Request.Form("SaveType") = "TournamentRemoveAdmin" then
	if not(bSysAdmin) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	If Len(Request.Form("selTournamentAdminID")) > 0 Then
		strSQL = "DELETE FROM lnk_m_a WHERE MALinkID = '" & Request.Form("selTournamentAdminID") & "'"
		oConn.Execute(strSQL)		
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "AssignAdmin.asp"
end if

If Request.Form("SaveType") = "ServerInfo" Then
	If Not(bSysAdmin Or IsTournamentAdmin(Request.Form("Tournament"))) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	End If
	strSQL = "UPDATE tbl_rounds SET "
	strSQL = strSQL & " ServerName = '" & CheckString(Request.Form("txtServerName")) & "' ,"
	strSQL = strSQL & " ServerIP = '" & CheckString(Request.Form("txtServerIP")) & "' ,"
	strSQL = strSQL & " ServerJoinPassword = '" & CheckString(Request.Form("txtServerJoinPassword")) & "' ,"
	strSQL = strSQL & " ServerRConPassword = '" & CheckString(Request.Form("txtServerRConPassword")) & "' ,"
	strSQL = strSQL & " MatchTime = '" & CheckString(Request.Form("txtMatchTime")) & "', "
	strSQL = strSQL & " BracketBlurb = '" & CheckString(Request.Form("txtBracketBlurb")) & "' "
	strSQL = strSQL & " WHERE RoundsID = '" & CheckString(Request.Form("RoundsID")) & "' "
	oConn.Execute(strSQL)
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "DisplayServers.asp?tournament=" & Server.URLEncode(Request.Form("Tournament"))
End If
'-----------------------------------------------
' Housekeeping
'-----------------------------------------------
set ors = nothing
set oConn = nothing	
set ors2 = nothing	
Response.Clear
Response.Redirect "/"
Response.End 
%>	
</BODY>
</HTML>
