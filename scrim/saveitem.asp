<% Option Explicit %>
<%
Server.ScriptTimeout = 1000
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: SaveItem"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADoDB.Connection")
Set oRS = Server.CreateObject("ADoDB.RecordSet")
Set oRS2 = Server.CreateObject("ADoDB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, blnLoggedIn, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<%
blnLoggedIn = Session("LoggedIn")

If Not(blnLoggedIn) Then
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	' Require login to perform action.
	Response.Clear
	Response.Redirect "/errorpage.asp?error=2"
End If

'' LadderJoin
Dim strTeamName, intTeamID, intFounderID
Dim intLadderID, strLadderName, blnClearToJoin 
Dim intEloTeamID, intPlayerID

'' Ladder Quit
Dim bTeamCaptain, bTeamFounder

'' Map List
Dim i

'' Change Map Time
Dim strMatchDate, dtmMatchDate, intMatchID

If Request.Form("SaveType")="EloLadderJoin" Then
	If Not(bSysAdmin) AND Not(IsTeamFounderByID(request.form("TeamID"))) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	Else
		strSQL = "select TeamID, TeamFounderID, TeamName, TeamEmail from tbl_Teams where TeamID='" & CheckString(Request.form("TeamID")) & "'"
		oRs.Open strSQL, oConn
		If Not (oRs.EOF AND oRs.BOF) Then
			Do while Not oRs.EOF
				strTeamName = oRs.fields("TeamName").value
				intTeamID = oRs.fields("TeamID").value
				intFounderID = oRs.fields("TeamFounderID").value
				oRs.MoveNext
			Loop
		End If
		oRs.Close

		strSQL = "select EloLadderID, EloLadderName from tbl_elo_Ladders where EloLadderName='" & CheckString(Request.Form("LadderToJoin")) & "'"
		oRs.Open strSQL, oConn
		If Not (oRs.EOF AND oRs.BOF) Then
			Do while Not oRs.EOF
				intLadderID = oRs.fields("EloLadderID").value
				strLadderName = oRs.fields("EloLadderName").Value
				oRs.MoveNext
			Loop
		End If
		oRs.Close
		
		blnClearToJoin = True
		
		strSQL = "SELECT PlayerID FROM lnk_elo_team_player etp "
		strSQL = strSQL & "	INNER JOIN lnk_elo_team et ON et.lnkEloTeamID = etp.lnkEloTeamID "
		strSQL = strSQL & "	WHERE PlayerID = '" & intFounderID & "'"
		strSQL = strSQL & "	AND Active = 1 AND EloLadderID = '" & intLadderID & "'"
		oRs.Open strSQL, oConn
		If Not (oRs.BOF AND oRs.EOF) Then
			blnClearToJoin = False '' Founder on a team on the ladder
		End If
		oRs.NextRecordSet

		If blnClearToJoin Then
			'-------------
			' Ladder Join
			'-------------
			strSQL = "SELECT lnkEloTeamID FROM lnk_elo_team WHERE TeamID='" & intTeamID & "' AND EloLadderID='" & intLadderID & "'"
			oRs.Open strSQL, oConn
			If Not(oRs.EOF AND oRs.BOF) Then
				intEloTeamID = oRs.Fields("lnkEloTeamID").Value
				oRs.NextRecordSet
				
				strSQL = "UPDATE lnk_elo_team SET Active = 1, Rating = 1000, WinStreak = 0, LossStreak = 0, Wins = 0, Losses = 0 WHERE lnkEloTeamID='" & intEloTeamID & "'"
				strSQL = strSQL & " DELETE FROM lnk_elo_team_player WHERE lnkEloTeamID = '" & intEloTeamID & "'"
				oConn.Execute(strSQL)
			Else
				oRs.NextRecordSet
				strSQL = "INSERT INTO lnk_elo_team (TeamID, EloLadderID, JoinDate, Rating, WinStreak, LossStreak, Active, Wins, Losses) VALUES ("
				strSQL = strSQL & "'" & intTeamID & "',"
				strSQL = strSQL & "'" & intLadderID & "',"
				strSQL = strSQL & "GetDate(), 1000, 0, 0, 1, 0, 0) "
				oConn.Execute(strSQL)
				strSQL = "SELECT 'lnkEloTeamID' = @@IDENTITY"
				oRs.Open strSQL, oConn
				intEloTeamID = oRs.Fields("lnkEloTeamID").Value
				oRs.NextRecordSet
			End If
			strSQL = "INSERT INTO lnk_elo_team_player (lnkEloTeamID, PlayerID, IsAdmin, JoinDate) VALUES ('" & intEloTeamID & "', '" & intFounderID & "', 1, GetDate())"
			oConn.Execute(strSQL)
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "/viewTeam.asp?team=" & server.urlencode(strTeamName)
		Else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "/errorpage.asp?error=1"
		End If
	End If
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	response.redirect "default.asp"
End If

'-----------------------------------------------
' Team Drop From Ladder
'-----------------------------------------------
if request.form("SaveType") = "QuitLadder" then
	
	strLadderName = Request.Form("LadderName")
	strTeamName = Request.Form("TeamName")

	bLadderAdmin = IsEloLadderAdmin(strLadderName)
	bTeamFounder = IsTeamFounder(strTeamName)
	bTeamCaptain = IsEloTeamCaptain(strTeamName, strLadderName)
	
	if not(bSysAdmin OR bLadderAdmin OR bTeamFounder OR bTeamCaptain) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	else
		' then check for active challenge
		intTeamID = Request.Form("TeamID")	
		strsql = "select EloLadderID from tbl_elo_ladders where Eloladdername='" & CheckString(strLadderName) & "'"
		ors.open strsql, oconn
		if not (ors.eof and ors.bof) then
			intLadderID = ors.fields("EloLadderID").value
		end if
		ors.close
	
		strsql = "Select lnkEloTeamID FROM lnk_elo_team where TeamID=" & intTeamID & " and EloladderID=" & intLadderID
		ors.open strSQL, oconn
		if not (ors.eof and ors.bof) then
			intEloTeamID = ors.fields(0).value
		end if
		ors.close
		
		'---------------
		' Delete Roster
		'---------------
		strSQL = "delete from lnk_elo_team_player where lnkEloTeamID=" & intEloTeamID
		oConn.Execute(strSQL)

		'-----------------
		' Deactivate Team
		'-----------------
		strSQL = "UPDATE lnk_elo_team SET Active = 0 WHERE lnkEloTeamID = '" & intEloTeamID & "'"
		oConn.Execute(strSQL)

		'--------------------
		' Deactivate Matches
		'--------------------
		strSQL = "UPDATE tbl_elo_matches SET MatchActive = 0 WHERE DefenderEloTeamID = '" & intEloTeamID & "' OR AttackerEloTeamID = '" & intEloTeamID & "'"
		oConn.Execute(strSQL)
		
		Response.Clear
		%>
		<script language="javascript">
			window.opener.location = window.opener.location.href;
			window.close();
		</script>
		<%
		Response.End 
	End If
	set ors = nothing
	set oConn = nothing	
	set ors2 = nothing	
	%>
	<script language="javascript">
		window.opener.location = window.opener.location.href;
		window.close();
	</script>
	<%	
	Response.End 
End If

'-----------------------------------------------
' Kick Player from Roster
'-----------------------------------------------
if request.form("savetype")="DropPlayer" then
	intEloTeamID=request.form("link")
	intPlayerID = request.form("playerid")

	strsql="select Eloladderid, teamid from lnk_elo_team where lnkEloTeamID=" & intEloTeamID
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		intTeamID = ors.fields(1).value
		intLadderID = ors.fields(0).value
	end if
	ors.close

	bLadderAdmin = IsEloLadderAdminById(intLadderID)
	bTeamFounder = IsTeamFounderByID(intTEamID)
	bTeamCaptain = IsEloTeamCaptainByID(intTeamID, intLadderID)
	
	if not(bSysAdmin OR bLadderAdmin OR bTeamFounder OR bTeamCaptain) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	else
		strsql="select teamname from tbl_teams where teamid=" & intTeamID
		ors.open strsql, oconn
		strTeamName =ors.fields(0).value
		ors.close
		strsql="select Eloladdername from tbl_elo_ladders where Eloladderid=" & intLadderID
		ors.open strsql, oconn
		strLadderName = ors.fields(0).value
		ors.close
		strsql="delete from lnk_elo_team_player where lnkEloTeamID=" & intEloTeamID & " and playerid=" & intPlayerID
		ors.open strsql, oconn
		'response.write Server.HTMLEncode(tname) & Server.HTMLEncode(lname)
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/teamScrimladderadmin.asp?team=" & server.urlencode(strTeamName) & "&ladder=" & server.urlencode(strLadderName)
	end if
end if

'-----------------------------------------------
' Promote to Captain
'-----------------------------------------------
if Request.Form("SaveType") = "PromoteScrimCaptain" then

	strSQL = "SELECT lnk.EloLadderID, lnk.lnkEloTeamID FROM lnk_elo_team lnk INNER JOIN lnk_elo_team_player l ON l.lnkEloTeamID = lnk.lnkEloTeamID "
	strSQL = strSQL & " WHERE lnkEloTeamPlayerID='" & Request.Form("PlayerList") & "'"
	oRs.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		intLadderID = oRs.FieldS("EloLadderID").Value
		intEloTeamID = oRs.Fields("lnkEloTeamID").Value
	End If
	oRS.NextRecordSet
		
	If Not(bSysAdmin OR IsEloLadderAdminByID(intLadderID) OR IsEloTeamCaptainByLinkID(intEloTeamID) OR IsTeamFounder(Request.Form("Team"))) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	ENd If

	strSQL="UPDATE lnk_elo_team_player set isadmin=1 where lnkEloTeamPlayerID='" & Request.Form("playerlist") & "'"
	'Response.write strSQL
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/TeamScrimLadderAdmin.asp?ladder=" & server.urlencode(Request.form("ladder")) & "&team=" & server.urlencode(Request.Form("team"))
end if
'-----------------------------------------------
' DemoteLeagueCaptain
'-----------------------------------------------
if Request.Form("SaveType") = "DemoteScrimCaptain" then

	strSQL = "SELECT lnk.EloLadderID, lnk.lnkEloTeamID FROM lnk_elo_team lnk INNER JOIN lnk_elo_team_player l ON l.lnkEloTeamID = lnk.lnkEloTeamID "
	strSQL = strSQL & " WHERE lnkEloTeamPlayerID='" & Request.Form("PlayerList") & "'"
	oRs.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		intLadderID = oRs.FieldS("EloLadderID").Value
		intEloTeamID = oRs.Fields("lnkEloTeamID").Value
	End If
	oRS.NextRecordSet
		
	strSQL="select teamfounderid from tbl_teams WHERE TeamName = '" & CHeckString(Request.Form("Team")) & "'"
	ors.open strsql, oconn
	if not (ors.EOF and ors.BOF) then
		intFounderID=ors.Fields(0).Value
	end if
	ors.Close 

	If Not(bSysAdmin OR IsEloLadderAdminByID(intLadderID) OR IsEloTeamCaptainByLinkID(intEloTeamID) OR IsTeamFounder(Request.Form("Team"))) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	ENd If

	if intFounderID <> intPlayerID then
		strSQL="UPDATE lnk_elo_team_player set isadmin=0 where lnkEloTeamPlayerID='" & Request.Form("playerlist") & "'"
		ors.open strSQL, oconn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/TeamScrimLadderAdmin.asp?ladder=" & server.urlencode(Request.form("ladder")) & "&team=" & server.urlencode(Request.Form("team"))
end if

'-----------------------------------------------
' Initiate a Challenge
'-----------------------------------------------
If Request.Form("SaveType") = "ChallengeTeam" Then
	strLadderName = Request.Form("Ladder")
	strTeamName = Request.Form("Team")

	bLadderAdmin = IsEloLadderAdmin(strLadderName)
	bTeamFounder = IsTeamFounder(strTeamName)
	bTeamCaptain = IsEloTeamCaptain(strTeamName, strLadderName)
	
	If not(bSysAdmin OR bLadderAdmin OR bTeamFounder OR bTeamCaptain) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	Else
		If Len(Request.Form("selLinkID") > 0) Then
			strSQL = "INSERT INTO tbl_elo_matches (DefenderEloTeamID, AttackerEloTeamID, ChallengeDate, EloLadderID, MatchActive) VALUES ( "
			strSQL = strSQL & "'" & CheckString(Request.Form("selLinkID")) & "', '" & CheckString(Request.Form("LinkID")) & "', GetDate(), '" & CheckString(Request.Form("LadderID")) & "', 1)"
			oConn.Execute(strSQL)
		End If

		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "/TeamScrimLadderAdmin.asp?ladder=" & server.urlencode(strLadderName) & "&team=" & server.urlencode(strTeamName)
	End If
End If

'-----------------------------------------------
' Save the Map Listing for a Scrim Ladder
'-----------------------------------------------
if Request.Form("SaveType") = "LadderMapList" then
	intLadderID = Request.Form("LadderID")
	if intLadderID <> "" then 
		if not(bSysAdmin or IsEloLadderAdminById(intLadderID)) then
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "/errorpage.asp?error=3"
		else
			
			strsql="delete from lnk_elo_maps where EloLadderID=" & intLadderID
			oconn.Execute strSQL

			strSQL = ""
			For i = 1 To (Request.Form("frm_current_maplist_map0").Count)
				strSQL = strSQL & "INSERT INTO lnk_elo_maps (MapID, EloLadderID) VALUES ('" & CheckString(Request.Form("frm_current_maplist_map0")(i)) & "', '" & CheckString(intLadderID) & "'); "
			Next
			
'			Response.Write Request.Form("frm_current_maplist") & " -- Form <BR>"
'			Response.Write Request.Form("frm_current_maplist").Count & " -- Count <BR>"
			If Len(strSQL) > 0 Then
				oConn.Execute(strSQL)
			End If
'			Response.Write strSQL
'			Response.End
			
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "GeneralAdmin.asp"
		end if
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/default.asp"
	end if
end if

'-----------------------------------------------
' Add Comms
'-----------------------------------------------
if Request.form("SaveType") = "Add_Communications" then
	strSQL="INSERT INTO tbl_elo_Comms ( ElomatchID, CommDate, CommAuthor, CommDead, Comms ) values ('" 
	strSQL = strSQL & Request.form("matchID") & "',GetDate(),'"  & CheckString(Request.Form("commauthor")) & "',0,'"
	strSQL = strSQL & CheckString(Request.Form("comms")) & "')"
	oConn.Execute(strSQL)
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/TeamScrimLadderAdmin.asp?ladder=" & server.urlencode(request("Ladder")) & "&team=" & server.urlencode(request("Team"))
end if
'-----------------------------------------------
' Edit Comms
'-----------------------------------------------
if Request.form("SaveType") = "Edit_Communications" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	end if
	strSQL= "update tbl_ELO_comms set Comms='" & replace(Request.Form("comms"), "'", "''") &  "' where Elocommid=" & Request.Form("commid")
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/TeamScrimLadderAdmin.asp?ladder=" & server.urlencode(request("Ladder")) & "&team=" & server.urlencode(request("Team"))
end if
'-----------------------------------------------
' Delete Comms
'-----------------------------------------------
if Request.QueryString("SaveType") = "Delete_Communications" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	end if
	if Request.QueryString("commid") <> "" then
		strSQL= "delete from tbl_elo_comms where Elocommid=" & Request.QueryString("commid")
		'Response.Write strSQl
		oRs.Open strSQL, oConn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect"/TeamScrimLadderAdmin.asp?ladder=" & server.urlencode(request("ladder")) & "&team=" & server.urlencode(request("team"))
end if

If Request.Form("SaveType") = "ChangeMatchTime" Then
	strLadderName = Request.Form("Ladder")
	strTeamName = Request.Form("Team")

	bLadderAdmin = IsEloLadderAdmin(strLadderName)
	bTeamFounder = IsTeamFounder(strTeamName)
	bTeamCaptain = IsEloTeamCaptain(strTeamName, strLadderName)
	
	If not(bSysAdmin OR bLadderAdmin OR bTeamFounder OR bTeamCaptain) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	Else
		strMatchDate = Request.Form("selMonth") & "/" & Request.Form("selDay") & "/" & Request.Form("selYear") & " " & Request.Form("selHour") & ":" & Request.Form("selMinute") & " " & Request.Form("selAMPM")

		strSQL = "UPDATE tbl_elo_matches SET MatchDate = '" & CheckString(strMatchDate) & "' WHERE EloMatchID = '" & CheckString(Request.Form("MatchID")) & "'"
		oConn.Execute(strSQL)

		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "/TeamScrimLadderAdmin.asp?ladder=" & server.urlencode(strLadderName) & "&team=" & server.urlencode(strTeamName)
	End If
End If

If Request.Form("SaveType") = "ChangeMaps" Then
	strLadderName = Request.Form("Ladder")
	strTeamName = Request.Form("Team")

	bLadderAdmin = IsEloLadderAdmin(strLadderName)
	bTeamFounder = IsTeamFounder(strTeamName)
	bTeamCaptain = IsEloTeamCaptain(strTeamName, strLadderName)
	
	If not(bSysAdmin OR bLadderAdmin OR bTeamFounder OR bTeamCaptain) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	Else

		strSQL = "UPDATE tbl_elo_matches SET "
		strSQL = strSQL & " Map1 = '" & CheckString(Request.Form("selMap1")) & "', "
		strSQL = strSQL & " Map2 = '" & CheckString(Request.Form("selMap2")) & "', "
		strSQL = strSQL & " Map3 = '" & CheckString(Request.Form("selMap3")) & "', "
		strSQL = strSQL & " Map4 = '" & CheckString(Request.Form("selMap4")) & "', "
		strSQL = strSQL & " Map5 = '" & CheckString(Request.Form("selMap5")) & "' "
		strSQL = strSQL & " WHERE EloMatchID = '" & CheckString(Request.Form("MatchID")) & "'"
		oConn.Execute(strSQL)

		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "/TeamScrimLadderAdmin.asp?ladder=" & server.urlencode(strLadderName) & "&team=" & server.urlencode(strTeamName)
	End If
End If

if Request.Form("SaveType") = "ReportMatch" then
	intMatchID = Request.Form("matchid")
	dtmMatchDate = Request.Form("matchdate")
	intLadderID = Request.Form("ladderid")
	strTeamName = Request.Form("matchlosername")
	strLadderName = Request.Form("LadderName")
	If Not(bSysAdmin or IsEloLadderAdmin(strLadderName) or IsTeamFounder(strTeamName) OR IsEloTeamCaptainByLinkID(Request.Form("MatchLoserID")) ) Then 
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	else
		strSQL = "select Elomatchid from tbl_elo_matches where Elomatchid=" & intMatchID
		ors.open strSQL, oconn
		if (ors.eof and ors.bof) Then
			Response.Redirect "/viewteam.asp?team=" & server.URLEncode(strTeamName)
		End If
		oRs.NextRecordSet

		Dim intLoserRating, intWinnerRating, intTeamCount

		strSQL = "SELECT Rating FROM lnk_elo_team WHERE lnkEloTeamID = '" & Request.Form("MatchLoserID") & "'"
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRs.BOF) Then
			intLoserRating = oRs.Fields("Rating").Value
		End If
		oRs.NextRecordSet

		strSQL = "SELECT Rating FROM lnk_elo_team WHERE lnkEloTeamID = '" & Request.Form("MatchWinnerID") & "'"
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRs.BOF) Then
			intWinnerRating = oRs.Fields("Rating").Value
		End If
		oRs.NextRecordSet
		
		strSQL = "SELECT 'TeamCount' = COUNT(lnkEloTeamID) FROM lnk_elo_team WHERE EloLadderID = '" & intLadderID & "' AND Active = 1 "
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRs.BOF) Then
			intTeamCount = oRs.Fields("TeamCount").Value
		End If
		oRs.NextRecordSet
		
		Dim intRatingDiff, intWinnerDefending
		
		intRatingDiff = round((intLoserRating * (intTeamCount / 2)) / intWinnerRating)
		If cBool(Request.Form("matchwinnerdefending")) Then
			intWinnerDefending = 1
		Else
			intWinnerDefending = 0
		End If
		strSQL = "insert into tbl_elo_History ("
		strSQL = strSQL & " EloMatchID, EloLadderID, AttackerEloTeamID, DefenderEloTeamID, "
		strSQL = strSQL & " Map1, Map2, Map3, Map4, Map5, "
		strSQL = strSQL & " Map1DefenderScore, Map2DefenderScore, Map3DefenderScore, Map4DefenderScore, Map5DefenderScore, "
		strSQL = strSQL & " Map1AttackerScore, Map2AttackerScore, Map3AttackerScore, Map4AttackerScore, Map5AttackerScore, "
		strSQL = strSQL & " Map1OT, Map2OT, Map3OT, Map4OT, Map5OT, "
		strSQL = strSQL & " MatchDate, AttackerRatingDiff, DefenderRatingDiff, MatchWinnerDefending) values ("
		strSQL = strSQL & intMatchID & ", " & intLadderID & ",'" & CheckString(Request.Form("AttackerID")) & "','" & CheckString(Request.Form("DefenderID")) & "',"
		strSQL = strSQL & "'" & CheckString(Request.Form("Map1")) & "', '" & CheckString(Request.Form("Map2")) & "','" & CheckString(Request.Form("Map3")) & "','" & CheckString(Request.Form("Map4")) & "','" & CheckString(Request.Form("Map5")) & "',"
		strSQL = strSQL & "'" & CheckString(Request.Form("Map1DefScore")) & "','" & CheckString(Request.Form("Map2DefScore")) & "','" & CheckString(Request.Form("Map3DefScore")) & "','" & CheckString(Request.Form("Map4DefScore")) & "','" & CheckString(Request.Form("Map5DefScore")) & "',"
		strSQL = strSQL & "'" & CheckString(Request.Form("Map1AttScore")) & "','" & CheckString(Request.Form("Map2AttScore")) & "','" & CheckString(Request.Form("Map3AttScore")) & "','" & CheckString(Request.Form("Map4AttScore")) & "','" & CheckString(Request.Form("Map5AttScore")) & "',"
		strSQL = strSQL & "'" & CheckString(Request.Form("intMap1OT")) & "','" & CheckString(Request.Form("intMap2OT")) & "','" & CheckString(Request.Form("intMap3OT")) & "','" & CheckString(Request.Form("intMap4OT")) & "','" & CheckString(Request.Form("intMap5OT")) & "',"
		strSQL = strSQL & "'" & CheckString(dtmMatchDate) & "','" & CheckString(intRatingDiff) & "','" & CheckString(intRatingDiff) & "', '" & CheckString(intWinnerDefending) & "')" 
		
'Response.Write strSQL & "<br><br>"
'		Response.End 

		oConn.Execute(strSQL)

		strSQL = "delete from tbl_elo_matches where Elomatchid=" & intmatchid
		oConn.Execute(strSQL)

		strSQL = "UPDATE lnk_elo_team SET Losses = Losses + 1, LossStreak = LossStreak + 1, WinStreak = 0, Rating = Rating - " & intRatingDiff & " WHERE lnkEloTeamID = '" & Request.Form("MatchLoserID") & "'"
'		Response.Write strSQL & "<br><br>"
'		Response.End 
		oConn.Execute(strSQL)
		
		strSQL = "UPDATE lnk_elo_team SET Wins = Wins + 1, LossStreak = 0, WinStreak = WinStreak + 1, Rating = Rating + " & intRatingDiff & " WHERE lnkEloTeamID = '" & Request.Form("MatchWinnerID") & "'"
		oConn.Execute(strSQL)
		'Response.Write strsql & "<br><br>"
		
		strSQL="update tbl_elo_Comms set CommDead=1 where Elomatchid=" & intMatchID
		oConn.Execute(strSQL)

	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/viewteam.asp?team=" & server.URLEncode(strTeamName)
end if

'-----------------------------------------------
' LeagueAssignAdmin
'-----------------------------------------------
if Request.Form("SaveType") = "AssignAdmin" then
	if not(bSysAdmin) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	end if
	If Len(Request.Form("PlayerID")) > 0 Then
		strSQL = "INSERT INTO lnk_elo_admin (PlayerID, EloLadderID) VALUES ('" & Request.Form("PlayerID") & "', '" & Request.Form("LadderID") & "')"
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
if Request.Form("SaveType") = "RemoveAdmin" Then
	if not(bSysAdmin) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	end if
	If Len(Request.Form("selEloLadderAdminID")) > 0 Then
		strSQL = "DELETE FROM lnk_elo_admin WHERE EloLadderAdminID = '" & Request.Form("selEloLadderAdminID") & "'"
		oConn.Execute(strSQL)		
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "AssignAdmin.asp"
end if

If Request.QueryString("SaveType") = "KillMatch" Then
	intMatchID = Request.QueryString("MatchID")
	strTeamName = Request.QueryString("Team")
	strLadderName = Request.QueryString("Ladder")
	intEloTeamID = Request.QueryString("LinkID")

	If Not(bSysAdmin or IsEloLadderAdmin(strLadderName) or IsTeamFounder(strTeamName) OR IsEloTeamCaptain(strTeamName, strLadderName) ) Then 
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	else
		strSQL = "UPDATE tbl_elo_matches SET MatchActive = 0 WHERE EloMatchID = '" & CheckString(intMatchID) & "' AND (DefenderEloTeamID = '" & intEloTeamID & "' OR AttackerEloTeamID = '" & intEloTeamID & "')"
		'Response.Write (strSQL)
		'Response.End
		oConn.Execute(strSQL)
	End If
	Response.Redirect "/TeamScrimLadderAdmin.asp?ladder=" & server.urlencode(strLadderName) & "&team=" & server.urlencode(strTeamName)
End If

oConn.Close 
Set oConn = Nothing
Set oRs = Nothing
Response.Clear
response.redirect "/default.asp"
%>