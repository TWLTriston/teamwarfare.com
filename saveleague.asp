<%
 	response.buffer=true
%>
<!-- #include file="include/i_funclib.asp" -->
<% 	
	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	Set oRS2 = Server.CreateObject("ADODB.RecordSet")
	
	oConn.ConnectionString = Application("ConnectStr")
	oConn.Open 	
 	
 	if request.form("saveType")="LeagueAdd" then
 		strLeagueName = request.form("txtleagueName")
 		intGameID = cint(request.form("selGameID"))
 		intInviteOnly = cint(request.form("selInviteOnly"))
 		intWinPoints = cint(request.form("txtWinPoints"))
 		intLossPoints = cint(request.form("txtLossPoints"))
 		intDrawPoints = cint(request.form("txtDrawPoints"))
		intNoShowPoints = cint(request.form("txtNoShowPoints"))
 		intMapWinPoints = cint(request.form("txtMapWinPoints"))
 		intMapLossPoints = cint(request.form("txtMapLossPoints"))
 		intMapDrawPoints = cint(request.form("txtMapDrawPoints"))
		intMapNoShowPoints = cint(request.form("txtMapNoShowPoints"))
 		intLocked = cint(request.form("selLocked"))
 		intSignUp = cint(request.form("selSignUp"))
 		intActive = cint(request.form("selActive"))
 		intScoring = cint(request.form("selScoring"))
		intLeagueID = cint(0 & request.Form("LeagueID"))
 		intNumConferences = cint(request.form("txtNumConferences"))
		intRosterLocked = cint(Request.Form("selRosterLocked"))
		intIdentifierID = Request.Form("selIdentifierID")
 		If request.Form("SaveMode") = "Edit" Then
 			'' Edit!
 			strSQL = "UPDATE tbl_leagues SET "
 			strSQL = strSQL & " LeagueName = '" & CheckString(strLeagueName) & "', "
 			strSQL = strSQL & " LeagueGameID = '" & CheckString(intGameID) & "', "
 			strSQL = strSQL & " LeagueInviteOnly = '" & CheckString(intInviteOnly) & "', "
 			strSQL = strSQL & " WinPoints = '" & CheckString(intWinPoints) & "', "
 			strSQL = strSQL & " LossPoints = '" & CheckString(intLossPoints) & "', "
 			strSQL = strSQL & " DrawPoints = '" & CheckString(intDrawPoints) & "', "
 			strSQL = strSQL & " NoShowPoints = '" & CheckString(intNoShowPoints) & "', "
 			strSQL = strSQL & " MapWinPoints = '" & CheckString(intMapWinPoints) & "', "
 			strSQL = strSQL & " MapLossPoints = '" & CheckString(intMapLossPoints) & "', "
 			strSQL = strSQL & " MapDrawPoints = '" & CheckString(intMapDrawPoints) & "', "
 			strSQL = strSQL & " MapNoShowPoints = '" & CheckString(intMapNoShowPoints) & "', "
 			strSQL = strSQL & " LeagueLocked = '" & CheckString(intLocked) & "', "
 			strSQL = strSQL & " IdentifierID = '" & CheckString(intIdentifierID) & "', "
 			strSQL = strSQL & " LeagueActive = '" & CheckString(intActive) & "', "
 			strSQL = strSQL & " Scoring = '" & CheckString(intScoring) & "', "
 			strSQL = strSQL & " SignUp = '" & CheckString(intSignUp) & "', "
 			strSQL = strSQL & " RosterLock = '" & CheckString(intRosterLocked) & "' "
 			
 			strSQL = strSQL & " WHERE LeagueID ='" & intLeagueID & "'"
 			oConn.Execute(strSQL)
 			oConn.CLose
 			Set oConn = nothing
 			Response.Clear
 			Response.Redirect "leagueadmin.asp"
 		Else
	 		strSQL="insert into tbl_leagues(LeagueName, LeagueGameID, IdentifierID, LeagueInviteOnly, WinPoints, "
	 		strSQL = strSQL & " LossPoints, DrawPoints, NoShowPoints, MapWinPoints, MapLossPoints, MapDrawPoints, MapNoShowPoints, LeagueLocked, LeagueActive, RosterLock, Scoring, SignUp) "
 	 		strSQL = strSQL & "values ('" & CheckString(strLeagueName) & "',"
	 		strSQL = strSQL & "'" & intGameID & "',"
	 		strSQL = strSQL & "'" & IdentifierID & "',"
	 		strSQL = strSQL & "'" & intInviteOnly & "',"
	 		strSQL = strSQL & "'" & intWinPoints & "',"
	 		strSQL = strSQL & "'" & intLossPoints & "',"
	 		strSQL = strSQL & "'" & intDrawPoints & "',"
	 		strSQL = strSQL & "'" & intNoShowPoints & "',"
	 		strSQL = strSQL & "'" & intMapWinPoints & "',"
	 		strSQL = strSQL & "'" & intMapLossPoints & "',"
	 		strSQL = strSQL & "'" & intMapDrawPoints & "',"
	 		strSQL = strSQL & "'" & intMapNoShowPoints & "',"
	 		strSQL = strSQL & "'" & intLocked & "',"
	 		strSQL = strSQL & "'" & intActive & "',"
	 		strSQL = strSQL & "'" & intRosterLocked & "',"
	 		strSQL = strSQL & "'" & intSignUp & "',"
	 		strSQL = strSQL & "'" & intScoring & "')"
	 		
	 		oConn.execute strSQL
	 		strSQL = "select @@identity"
	 		oRS.open strSQL, oConn
	 		intLeagueID = oRS.fields(0)
	 		oRS.close
	 		
	 		for intCounter = 1 to intNumConferences
	 			strSQL = "insert into tbl_league_conferences(LeagueID) values ('" & intLeagueID & "')"
	 			oConn.execute strSQL
	 		next
	 		
	 		response.clear
	 		response.redirect "leagueConferenceConfig.asp?LeagueID=" & intLeagueID
	 	End If
 	end if
 	
 	if request.form("SaveType")="ConferenceSettings" then
 		intLeagueID = request.form("intLeagueID")
		strSQL = ""
 		for intCounter = 1 to Request.Form("hdnConferenceID").Count
			intConferenceID = Request.Form("hdnConferenceID")(intCounter) 		

			strConferenceName = request.Form("txtConferenceName")(intCounter)
			strSQL = strSQL & "update tbl_league_conferences set ConferenceName='" & CheckString(strConferenceName) & "' "
			strSQL = strSQL & "where LeagueConferenceID=" & intConferenceID & ";"
	 				
			intDivCount = request.Form("txtConferenceDivCount")(intCounter)
			for intCounter2 = 1 to intDivCount
				strSQL = strSQL & "insert into tbl_league_divisions(LeagueConferenceID, LeagueID, DivisionName) values "
				strSQL = strSQL & "('" & intConferenceID & "',"
				strSQL = strSQL & "'" & intLeagueID & "',"
				strSQL = strSQL & "'None Assigned');"
 				
			next
		next
		If len(strSQL) > 0 Then	
			oConn.execute strSQL
		End If
		response.clear
		response.redirect "leagueDivisionConfig.asp?intLeagueID=" & intLeagueID
 	end if
 	
 	if request.form("SaveType")="Divisions" then
 		for intCounter = 1 to request.form("hdnDivID").count
 			intDivID = request.form("hdnDivID")(intCounter)
 			strDivName = request.form("txtDivName")(intCounter)
 			strSQL = "update tbl_league_divisions set DivisionName='" & CheckString(strDivName) & "' "
 			strSQL = strSQL & "where LeagueDivisionID='" & intDivID & "'"
 			oConn.execute strSQL
 		next
		response.clear
		response.redirect "leagueadmin.asp"
 	End if

 	if request.form("SaveType")="LeagueEditHistory" then
		if not(IsSysAdmin() OR IsLeagueAdminByID(intLeagueID)) Then
			oConn.Close
			Set oConn = Nothing
			Set oRS = Nothing
			response.clear
			response.redirect "errorpage.asp?error=3"
		End If
		Dim intMapHomeScore(6), intMapVisitorScore(6)
		intLeagueID = Request.Form("LeagueID")
		strMatchDate = Request.Form("MatchDate")
		strHomeName = Request.Form("HomeTeamName")
		strVisitorName = Request.Form("VisitorTeamName")
		strLeagueName = Request.Form("LeagueName")
		intMapHomeScore(1) = Request.Form("HMapScore1")
		intMapVisitorScore(1) = Request.Form("VMapScore1")
		intMapHomeScore(2) = Request.Form("HMapScore2")
		intMapVisitorScore(2) = Request.Form("VMapScore2")
		intMapHomeScore(3) = Request.Form("HMapScore3")
		intMapVisitorScore(3) = Request.Form("VMapScore3")
		intMapHomeScore(4) = Request.Form("HMapScore4")
		intMapVisitorScore(4) = Request.Form("VMapScore4")
		intMapHomeScore(5) = Request.Form("HMapScore5")
		intMapVisitorScore(5) = Request.Form("VMapScore5")
		intHistoryID = Request.Form("HistoryID")
		strFromURL = Request.Form("FromURL")
		
		For i = 1 to 5
			If Len(intMapHomeScore(i)) = 0 Then
				intMapHomeScore(i) = 0
			End If
			If Len(intMapVisitorScore(i)) = 0 Then
				intMapVisitorScore(i) = 0
			End If
		Next
		strSQL = "UPDATE tbl_league_history SET "
		strSQL = strSQL & " Map1HomeScore = '" & intMapHomeScore(1) & "', "
		strSQL = strSQL & " Map2HomeScore = '" & intMapHomeScore(2) & "', "
		strSQL = strSQL & " Map3HomeScore = '" & intMapHomeScore(3) & "', "
		strSQL = strSQL & " Map4HomeScore = '" & intMapHomeScore(4) & "', "
		strSQL = strSQL & " Map5HomeScore = '" & intMapHomeScore(5) & "', "
		strSQL = strSQL & " Map1VisitorScore = '" & intMapVisitorScore(1) & "', "
		strSQL = strSQL & " Map2VisitorScore = '" & intMapVisitorScore(2) & "', "
		strSQL = strSQL & " Map3VisitorScore = '" & intMapVisitorScore(3) & "', "
		strSQL = strSQL & " Map4VisitorScore = '" & intMapVisitorScore(4) & "', "
		strSQL = strSQL & " Map5VisitorScore = '" & intMapVisitorScore(5) & "' "
		strSQL = strSQL & " WHERE LeagueHistoryID = '" & intHistoryID & "'"
'response.write strsql
'response.end
		oConn.Execute(strSQL)
		
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect strFromURL
	End If
	
 	If request.form("SaveType")="LeagueEditMatch" then
		If not(IsSysAdmin() OR IsLeagueAdminByID(intLeagueID)) Then
			oConn.Close
			Set oConn = Nothing
			Set oRS = Nothing
			response.clear
			response.redirect "errorpage.asp?error=3"
		End If
		intLeagueID = Request.Form("LeagueID")
		strMatchDate = Request.Form("HMatchDate")
		strHomeName = Request.Form("HomeTeamName")
		strVisitorName = Request.Form("VisitorTeamName")
		strLeagueName = Request.Form("LeagueName")
		intLeagueMatchID = Request.Form("LeagueMatchID")
		strFromURL = Request.Form("FromURL")
		
		strSQL = "UPDATE tbl_league_matches SET "
		strSQL = strSQL & " MatchDate = '" & cDate(strMatchDate) & "'"
		strSQL = strSQL & " WHERE LeagueMatchID = '" & intLeagueMatchID & "'"

		oConn.Execute(strSQL)
		
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect strFromURL
	End If

%>
 		
 		