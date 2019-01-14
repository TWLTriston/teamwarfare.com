<%' Option Explicit %>
<%
Server.ScriptTimeout = 1000
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
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	' Require login to perform action.
	Response.Clear
	Response.Redirect "errorpage.asp?error=2"
end if

Response.Write "<font color=#ffffff>Querystring: " & Request.QueryString
Response.Write "<br>Form data: " & Request.Form & "</font></br>"
'-----------------------------------------------
' Save team data
'-----------------------------------------------
if request.form("SaveType") = "team" then 
	oritname = request.form("TeamName")
	tName = Trim(Request.Form("TeamName"))
	tName = Replace(tName, chr(0160), "") ' REMOVING THE NULL ABILITY
	if tname = "" then 
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		response.redirect "errorpage.asp?error=7"
	end if
	tEmailName	= Request.Form("TeamName")
	tTag		= CheckString(Request.Form("TeamTags"))
	tURL		= CheckString(Request.Form("TeamURL"))
	tEmail		= CheckString(Request.Form("TeamEmail"))
	tIRC		= CheckString(Request.Form("TeamIRC"))
	tIRCServer	= CheckString(Request.Form("TeamIRCServer"))
	tLogo		= CheckString(Request.Form("TeamLOGO"))
	tJoinPass	=  CheckString(request.Form("TeamJoinPassword"))
	tconfirmjoinpass = CheckString(request.Form("TeamConfirmJoinPassword"))
	tDesc			= CheckString(ForumEncode(request.Form("TeamDesc")))
	tnewFounderID	= CheckString(request.Form("newFounderID"))
	tOldFounderID	= CheckString(request.Form("OldFounderID"))
	sMethod			= Request.Form("SaveMethod")
	tid = Request.Form("TeamID")
	strOldTeamName = Request.Form("OldTeamName")
	
	if tjoinpass <> tconfirmjoinpass then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		if smethod="Edit" then
			Response.Clear
			response.redirect "addteam.asp?isedit=true&team=" & server.urlencode(request.form("TeamName")) & "&error=2"
		else
			Response.Clear
			response.redirect "addteam.asp?error=2"
		end if
	end if
	
	if smethod="New" then
		strSQL = "EXECUTE usp_TeamAdd '" & CheckString(tName) & "', '" & tTag & "', '" & tURL & "', "
		strSQL = strSQL & " '" & tEmail & "', '" & Session("PlayerID") & "','" & tJoinPass & "','" & tDesc & "','" & tirc & "','" & tIRCServer & "', '" & tLogo & "'"
		oRS.Open strSQL, oConn
		If Not(oRs.EOF AND oRS.BOF) Then
			strReturnCode = oRs.Fields("RETURN_CODE").Value
		End If
		oRS.NextRecordset 
		Select Case strReturnCode
			Case "TEAM_EXIST"
				oConn.Close 
				Set oConn = Nothing
				Set oRs = Nothing
				Response.Clear
				Response.Redirect "/addteam.asp?error=1"
			Case "OK"
				' All is good
			Case Else
				' ERROR!!!
		End Select
	elseif smethod="Edit" then 
		if IsTeamFounder(oritname) or IsSysAdmin() then
			If bSysAdmin Then 
					strSQL = "SELECT TeamName FROM tbl_teams WHERE TeamName= '" & CheckString(tName) & "' AND TeamID <> '" & tid & "'"
					oRs.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						oRS.Close 
						Set oConn = Nothing
						Set oRs = Nothing
						Response.Clear
						Response.Redirect "addteam.asp?isedit=true&team=" & Server.URLEncode(strOldTeamname) & "&error=4"
					End If
					oRs.Close
					if strOldTeamname <> tName THen
						strSQL = "UPDATE tbl_teams SET teamName = '" & CheckString(tname) & "' WHERE TeamID = '" & tID & "'"
						oConn.Execute (strSQL)
						strSQL = "INSERT INTO tbl_transaction (TransactionDetails, TransactionTime) VALUES ('Team name change - " & CheckString(Session("uName")) & " - " & Request.ServerVariables("REMOTE_ADDR") & " - changed " & strOldTeamname & " to " & Tname & "', GetDate())"
						oConn.Execute (strSQL)
					End If
			End If
			If Len(Trim(tNewFounderID & "")) > 0 Then
				strSQL="update tbl_teams set teamfounderid = '" & tNewFounderID & "', TeamTag='" & ttag & "', TeamURL='" & turl & "', TeamEmail='" & temail& "', TeamJoinPassword='" & tjoinpass & "', TeamDesc='" & tDesc & "', TeamIRC='" & tirc & "', TeamIRCServer='" & tircserver & "', TeamLogoURL='" & tlogo & "' where TeamName='" & CheckString(tName) & "'"
			Else
				strSQL="update tbl_teams set TeamTag='" & ttag & "', TeamURL='" & turl & "', TeamEmail='" & temail& "', TeamJoinPassword='" & tjoinpass & "', TeamDesc='" & tDesc & "', TeamIRC='" & tirc & "', TeamIRCServer='" & tircserver & "', TeamLogoURL='" & tlogo & "' where TeamName='" & CheckString(tName) & "'"
			End If
			oConn.Execute (strSQL)
		else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=3"
		end if
	'response.write "Success"
	end if

'	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
'	Mailer.RemoteHost  = "127.0.0.1"
	'Mailer.FromName    = "TWL: Team Information"
	'mailer.FromAddress = "automailer@teamwarfare.com"
	'Mailer.AddRecipient tName, tEmail
	'Mailer.Subject     = "TWL: " & temailName & " Profile Information"
	'Text = temailName & ", the TWL has been updated with new information see below and file for your records."
	'text = text & "Please do not reply to this message." & vbcrlf & vbcrlf
	'text = text & "Team Name: " & temailName & vbcrlf
	'text = text & "Team Tag: " & ttag & vbcrlf
	'text = text & "Team URL: " & tURL & vbcrlf
	'text = text & "Team IRC Channel: " & tirc & vbcrlf
	'text = text & "Team IRC Server: " & tircserver & vbcrlf
	'text = text & "Team Logo Link: " & tlogo & vbcrlf
	'text = text & "Team Join Password: " & tJoinPass & vbcrlf
	'text = text & "Team Email: " & TEmail  & vbcrlf
	'text = text & "Team Description: " & tDesc & vbcrlf
	'text = text & vbcrlf & vbcrlf & "If this information is in error please login to the TWL and change appropriately." & vbcrlf
	'Mailer.BodyText    = text
	'on error resume next
	'Mailer.SendMail
	'on error goto 0
	'set Mailer = nothing

	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "viewteam.asp?team=" & server.urlencode(request.form("TeamName"))
end if
'-----------------------------------------------
' Save player data
'-----------------------------------------------

if request.form("SaveType") = "player" then 
	pName = trim(Request.Form("PlayerName")) ' 
	pName = Replace(pName, chr(0160), "") ' REMOVING THE NULL ABILITY
	bSendActivation = True
	if pname = "" then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=7"
	end if
	pID			= Request.Form("PlayerID")
	pOldName	= Request.Form("OldPlayerName")
	pICQ		= Request.Form("PlayerICQ") 
	pEmail		= Trim(Request.Form("PlayerEmail"))
	pPass		= CheckString(Trim(Request.Form("PlayerPassword")))
	pconfirm	= CheckString(Trim(Request.Form("PlayerConfirmPassword")))
	sMethod		= Request.Form("SaveMethod")
	hideemail	= Request.Form("HideEmail")
	playersignature = CheckString(ForumEncode(request.form("signature")))
	PlayerTitle	= CheckString(request.form("PlayerTitle"))
	OldPlayerTitle = CheckString(request.Form("OldPlayerTitle"))
	pOldEmail = Request.Form("oldEMail")
	PlayerActive = CheckString(Request.Form("PlayerActive"))
	ForumAccess = CheckString(Request.Form("ForumAccess"))
	Contributor = CheckString(Request.Form("Contributor"))
	ContributorAmount = CheckString(Request.Form("ContributorAmount"))
	CanActivate = CheckString(Request.Form("CanActivate"))
	IsSuspended = CheckString(Request.Form("Suspension"))
	ForumBanLiftDate = CheckString(Request.Form("ForumBanLiftDate"))
	SiteBanLiftDate = CheckString(Request.Form("SiteBanLiftDate"))
	SuspensionLiftDate = CheckString(Request.Form("SuspensionLiftDate"))
	Comment = CheckString(Request.Form("Comment"))
	AccessID = CheckString(Session("PlayerID"))
	'' Redacted
	if smethod="Edit" then
		if PlayerTitle <> OldPlayerTitle then
			strSQL = "INSERT INTO tbl_transaction (TransactionDetails, TransactionTime) VALUES ('Title change - " & CheckString(Session("uName")) & " - " & Request.ServerVariables("REMOTE_ADDR") & " - changed " & pName & " title to " & PlayerTitle & "', GetDate())"
			oConn.Execute (strSQL)
		end if
	end if
	if PlayerTitle = "" then 
		PlayerTitle = "TWL Member"
	end if
	
	if hideemail = "on" then
		phideemail=1
	else
		phideemail=0
	end if 
	if ppass <> pconfirm then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "addplayer.asp?error=1"
	end if
	strActivationCode = Session.SessionID & Int((6000 - 0 + 1) * Rnd + 0)
	if sMethod="New" then 
		strSQL= "select Playerhandle from tbl_players where playerhandle = '" & CHeckString(pName) & "' OR PlayerEmail = '" & CheckString(pEmail) & "'"
		oRs.Open strSQL, oConn
		isUsed = false
		if not (ors.eof and ors.bof) then
			isUsed=true
		end if
		ors.NextRecordset 
		if isUsed then
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.redirect "addplayer.asp?error=2"
		else
			jDate=date
			strSQL = "insert into tbl_Players (PlayerHandle, PlayerJoinDate, PlayerPassword, PlayerLastVisit, PlayerEmail, PlayerICQ, PlayerHideEmail, PlayerTitle, PlayerSignature, PlayerActive, ActivationCode) values "
			strSQL = strSQL & "('" & checkstring(pName) & "','" & jDate & "','" & enc2pass & "', '" & jDate & "','" & checkstring(pEmail) & "','" & pICQ & "', '" & pHideEmail & "', '" & playerTitle & "', '" & playersignature & "', 'N', '" & strActivationCode & "')"
			oconn.Execute (strsql)
		end if
	elseif smethod="Edit" then
		if not(Session("LoggedIn")) then
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=2"
		else
			if (pname = session("uName") or SysAdmin) then
				If pOldEmail <> pEmail Then
					strSQL = "SELECT PlayerHandle FROM tbl_players WHERE PlayerEmail = '" & CheckString(pEmail) & "'"
					oRs.Open strSQL, oConn
					If Not(ors.EOF and oRS.BOF) Then
						Response.Clear 
						Response.Redirect "/addplayer.asp?isedit=true&PlayerName=" & Server.URLEncode(pName) & "&error=3"
					End If
					oRS.NextRecordset
				Else
					bSendActivation = False
				End If
				'-------------------------------------------------
				' Change name code.... clean up this section later
				'-------------------------------------------------
				If bSysAdmin Then 
					strSQL = "SELECT PlayerHandle FROM tbl_players WHERE PlayerHandle = '" & CheckString(pName) & "' AND PlayerID <> '" & pID & "'"
					oRs.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						oRS.Close 
						Set oConn = Nothing
						Set oRs = Nothing
						Response.Clear
						Response.Redirect "addplayer.asp?isedit=true&PlayerName=" & Server.URLEncode(pOldName) & "&error=4"
					End If
					oRs.Close
					if pOldName <> pName THen
						strSQL = "UPDATE tbl_players SET PlayerHandle = '" & CheckString(pName) & "' WHERE PlayerID = '" & pID & "'"
						oConn.Execute (strSQL)
						strSQL = "INSERT INTO tbl_transaction (TransactionDetails, TransactionTime) VALUES ('Player name change - " & CheckString(Session("uName")) & " - " & Request.ServerVariables("REMOTE_ADDR") & " - changed " & pOldName & " to " & pName & "', GetDate())"
						oConn.Execute (strSQL)
					End If
				End If
				
				If bSysAdmin Then
					strSQL = "SELECT Suspension, PlayerCanActivate, ForumAccess FROM tbl_players WHERE PlayerID = '" & pID & "'"
					oRs.Open strSQL, oConn
					intComment = 0
					strCommentAuto = ""
					If Not(CStr(oRs.Fields("Suspension").Value & "") = CStr(IsSuspended)) Then
						intComment = 1
						If (IsSuspended = 1) Then
							strCommentAuto = strCommentAuto & "<b>Added Suspension</b> "
						Else
							strCommentAuto = strCommentAuto & "<b>Removed Suspension</b> "
						End If
					End If
					
					If Not(CStr(oRs.Fields("PlayerCanActivate").Value & "") = CStr(CanActivate)) Then
						intComment = 1
						If (CanActivate = 1) Then
							strCommentAuto = strCommentAuto & "<b>Removed Site Ban</b> "
						Else
							strCommentAuto = strCommentAuto & "<b>Added Site Ban</b> "
						End If
					End If
					
					If Not(CStr(oRs.Fields("ForumAccess").Value & "") = CStr(ForumAccess)) Then
						intComment = 1
						If (ForumAccess = 1) Then
							strCommentAuto = strCommentAuto & "<b>Removed Forum Ban</b> "
						Else
							strCommentAuto = strCommentAuto & "<b>Added Forum Ban</b> "
						End If
					End If
					oRs.Close
					If Len(strCommentAuto) > 0 Then
						Comment = strCommentAuto & "<br />" & Comment
					End If
				End If
				
				If IsSysAdmin() And intComment = 1 Then
					strSQL = "INSERT INTO tbl_player_comments (PlayerID, Comment, AdminID) values "
					strSQL = strSQL & "('" & PID & "','" & Comment & "','" & AccessID & "')"
					oConn.Execute (strSQL)
				End If
				
				strSQL="update tbl_players set "
				if len(pPass) > 0 then
					strSQL = strSQL & " PlayerPassword='" & enc2pass & "', "
				end if
				If (pOldEmail <> pEmail) Then
					strSQL = strSQL & " PlayerEmail='" & CheckString(pEmail) & "', "
					If Not(SysAdmin) Then
						strSQL = strSQL & " PlayerActive='N', "
					End If
					randomize
					strSQL = strSQL & " ActivationCode='" & strActivationCode & "', "
				End If
				
				strSQL = strSQL & " PlayerICQ='" & CheckString(pICQ) & "', "
				strSQL = strSQL & " PlayerHideEmail=" & pHideEMail & ", "
				If IsSysAdmin() Then
					strSQL = strSQL & " PlayerTitle='" & playertitle & "', "
					strSQL = strSQL & " ForumAccess='" & ForumAccess & "', "
					strSQL = strSQL & " PlayerActive='" & PlayerActive & "', "
					strSQL = strSQL & " Contributor='" & Contributor & "', "
					strSQL = strSQL & " ContributorAmount='" & ContributorAmount & "', "
					strSQL = strSQL & " PlayerCanActivate='" & CanActivate & "', "
					strSQL = strSQL & " Suspension='" & IsSuspended & "', "
					strSQL = strSQL & " SuspensionLiftDate='" & SuspensionLiftDate & "', "
					strSQL = strSQL & " ForumBanLiftDate='" & ForumBanLiftDate & "', "
					strSQL = strSQL & " SiteBanLiftDate='" & SiteBanLiftDate & "', "
					
				End If
				strSQL = strSQL & " PlayerSignature='" & playersignature & "' "
				strSQL = strSQL & " where playerhandle='" & replace(pname, "'", "''") & "'"
				oConn.Execute (strSQL)
				
			else
				oConn.Close 
				Set oConn = Nothing
				Set oRs = Nothing
				Response.Clear
				response.redirect "errorpage.asp?error=3"
			end if
		end if
	end if
	If Not(bSysAdmin) And bSendActivation Then
		if smethod="New" then
			Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			Mailer.RemoteHost  = "127.0.0.1"
			Mailer.FromName    = "TWL: Member Information"
			mailer.FromAddress = "automailer@web.teamwarfare.com"
			Mailer.AddRecipient pName, pEmail
			Mailer.Subject     = "TWL: Welcome, " & pname & ", to the TWL"
			Text = pName & ", welcome to TeamWarfare. We hope you enjoy our system."
			text = text & "This email serves as confirmation that you have registered, and contains activation instructions for your new account." & vbcrlf & vbcrlf
			text = text & "Summary of member information: " & vbcrlf
			text = text & "Login Name: " & pName & vbcrlf
			text = text & "Password: " & pPass & vbcrlf
			text = text & "Activation Code: " & strActivationCode & vbCrLf & vbCrLf
			text = text & "Click the link below to activate your account: " & vbCrLf
			Text = Text & "http://www.teamwarfare.com/ActivateAccount.asp?playername=" & Server.URLEncode(pname) & "&actcode=" & Server.URLEncode(strActivationCode) & vbCrLf & vbCrLf
			text = text & "If you need to create a team, click the link below:" & vbCrLf
			Text = Text & "http://www.teamwarfare.com/addteam.asp" & vbCrLf
			Mailer.BodyText    = text
			' Keep out the bad guys
			If instr(1,pEmail,"cjb.net") = 0 Then
			  If Not(Mailer.SendMail) Then
			    if Mailer.Response <> "" then
			      strError = Mailer.Response
			    else
			      strError = "Unknown"
			    end if
			    Response.Write "Mail failure occured. Reason: " & strError
			  end if
			End If
			' End keeping out the bad guys
			set mailer = nothing
			Session.Abandon()
			Response.Cookies("User")("UName")=""
			Response.Cookies("User")("UserInfo")=""
		ElseIf sMethod = "Edit" Then
			Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			Mailer.RemoteHost  = "127.0.0.1"
			Mailer.FromName    = "TWL: Member Information"
			mailer.FromAddress = "automailer@web.teamwarfare.com"
			Mailer.AddRecipient pName, pEmail
			Mailer.Subject     = "TWL: Welcome, " & pname & ", to the TWL"
			Text = pName & ", the email address on your account was changed. You must re-activate your account before logging in again." & vbCrLf
			text = text & "Click the link below to activate your account: " & vbCrLf
			Text = Text & "http://www.teamwarfare.com/activateaccount.asp?playername=" & Server.URLEncode(pname) & "&actcode=" & strActivationCode & vbCrLf
			text = text & "Your activation code is: " & strActivationCode & vbCrLf
			Mailer.BodyText    = text
			If Not(Mailer.SendMail) Then
			  if Mailer.Response <> "" then
			    strError = Mailer.Response
			  else
			    strError = "Unknown"
			  end if
			  Response.Write "Mail failure occured. Reason: " & strError
			end if
			set mailer = nothing
			Session.Abandon()
			Response.Cookies("User")("UName")=""
			Response.Cookies("User")("UserInfo")=""
		End If
	End If
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	
	Response.Clear
	Response.Redirect "accountupdated.asp"
end if
'-----------------------------------------------
' Save ladder data
'-----------------------------------------------
if request.form("SaveType") = "ladder" then 
	if not IsSysAdmin() then 
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
%>
	Ladder Name: <% =Server.HTMLEncode( Request.Form("LadderName") )%> <br>
	Ladder Admin: <% =Server.HTMLEncode(Request.Form("LadderAdmin")) %> <br>
	Current Ladder Count: 
<%
	lName = Request.Form("LadderName")
	lAdmin =  Request.Form("LadderAdmin") 
	lGame = Request.Form("LadderGame")
	iGame = Request.Form("GameID")
	lactive = Request.Form("LadderActive")
	labbr = Request.Form("LadderAbbreviation")
	llocked = Request.Form("LadderLocked") 
	lchallenge = Request.Form("LadderChallenge")
	lrules = Request.Form("LadderRules")
	MinPlayer = Request.Form("MinPlayer")
	maxroster = Request.Form("MaxRoster")
	maps = Request.Form("maps")
	mapconfiguration = Request.Form("mapconfiguration")
	timezone = Request.Form("timezone") 
	timeoptions = Request.Form("timeoptions")
	Scoring = Request.Form("Scoring")
	restrank = Request.Form("restrank")
	chrRequireDistinctMaps = Request.Form("RequireDistinctMaps")
	intIdentifierID = Request.Form("selIdentifierID")
	intChallengeDays = Cint(0 & Request.Form("chkChallengeSunday")) + _
						Cint(0 & Request.Form("chkChallengeMonday")) + _
						Cint(0 & Request.Form("chkChallengeTuesday")) + _
						Cint(0 & Request.Form("chkChallengeWednesday")) + _
						Cint(0 & Request.Form("chkChallengeThursday")) + _
						Cint(0 & Request.Form("chkChallengeFriday")) + _
						Cint(0 & Request.Form("chkChallengeSaturday"))
	intMatchDays = Cint(0 & Request.Form("chkMatchSunday")) + _
					Cint(0 & Request.Form("chkMatchMonday")) + _
					Cint(0 & Request.Form("chkMatchTuesday")) + _
					Cint(0 & Request.Form("chkMatchWednesday")) + _
					Cint(0 & Request.Form("chkMatchThursday")) + _
					Cint(0 & Request.Form("chkMatchFriday")) + _
					Cint(0 & Request.Form("chkMatchSaturday"))
	
	
	if request.form("savemethod")="Edit" then
		strSql = "select * from tbl_Ladders where LadderName ='" & CheckString(Request.Form("OldName")) & "'"
	else
		strSql = "select * from tbl_Ladders where LadderName ='" & CheckString(Request.Form("LadderName")) & "'"
	end if
	oRs.Open strSQL, oConn
	if oRs.EOF and oRs.BOF then
		oRs.Close 
		jDate=date
		strSQL = "insert into tbl_Ladders ( LadderName, LadderActive, "
		strSQL = strSQL & " LadderAdmin, LadderGame, LadderAbbreviation, "
		strSQL = strSQL & " LadderLocked, LadderChallenge, LadderRules, "
		strSQL = strSQL & " MinPlayer, RosterLimit, RestRank, Maps, MapConfiguration, TimeZone, "
		strSQL = strSQL & " TimeOptions, Scoring, GameID, RequireDistinctMaps, ChallengeDays, MatchDays, IdentifierID "
		strSQL = strSQL & " ) values ("
		strSQL = strSQL & " '" & CheckString(lName) & "','" & lactive & "', "
		strSQL = strSQL & " '" & lAdmin & "','" & lGame & "', '" & labbr & "',"
		strSQL = strSQL & " '" & llocked & "', '" & lChallenge & "', '" & CheckString(lRules) & "', "
		strSQL = strSQL & " '" & MinPlayer & "', '" & maxroster & "', '" & restrank & "', '" & Maps & "', '" & MapConfiguration & "', '" & TimeZone & "', "
		strSQL = strSQL & " '" & TimeOptions & "', '" & Scoring & "', '" & iGame  & "', '" & chrRequireDistinctMaps & "', '" & intChallengeDays & "', '" & intMatchDays & "', '" & CheckString(intIdentifierID) & "'"
		strSQL = strSQL & " )"
		oRs.Open strSQL, oConn
	else
		thisladderid=ors.fields("LadderID").value
		ors.close
		strsql = "update tbl_ladders set "
		strSQL = strSQL & " LadderName='" & CheckString(lname) & "', "
		strSQL = strSQL & " LadderAdmin='" & lAdmin & "', "
		strSQL = strSQL & " LadderGame='" & lGame  & "', "
		strSQL = strSQL & " LadderActive='" & lactive & "', "
		strSQL = strSQL & " LadderAbbreviation='" & labbr & "', "
		strSQL = strSQL & " RestRank='" & restrank & "', "
		strSQL = strSQL & " LadderLocked='" & llocked & "', "
		strSQL = strSQL & " LadderChallenge='" & lchallenge & "', "
		strSQL = strSQL & " LadderRules='" & CheckString(lRules) & "', "
		strSQL = strSQL & " MinPlayer='" & MinPlayer & "', "
		strSQL = strSQL & " RosterLimit='" & maxroster & "', "
		strSQL = strSQL & " Maps='" & Maps & "', "
		strSQL = strSQL & " MapConfiguration='" & MapConfiguration & "', "
		strSQL = strSQL & " TimeZone='" & TimeZone & "', "
		strSQL = strSQL & " TimeOptions='" & TimeOptions & "', "
		strSQL = strSQL & " Scoring='" & Scoring & "', "
		strSQL = strSQL & " GameID='" & iGame & "', "
		strSQL = strSQL & " RequireDistinctMaps='" & chrRequireDistinctMaps & "', "
		strSQL = strSQL & " ChallengeDays ='" & intChallengeDays & "', "
		strSQL = strSQL & " MatchDays='" & intMatchDays & "', "
		strSQL = strSQL & " IdentifierID='" & intIdentifierID & "' "
		strSQL = strSQL & " where ladderid='" & thisladderid & "'"
		ors.open strsql, oConn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/adminops.asp?rAdmin=Ladder"
end if
'-----------------------------------------------
' Save player ladder data
'-----------------------------------------------
if request.form("SaveType") = "playerladder" then 
	if not IsSysAdmin() then 
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	lName = Request.Form("PlayerLadderName")
	labbr = Request.Form("Abbreviation")
	llocked = Request.Form("Locked") 
	lactive = Request.Form("Active") 
	lchallenge = Request.Form("Challenge")
	intGame = Request.Form("GameID")
	
	if request.form("savemethod")="Edit" then
		strSql = "select * from tbl_playerLadders where PlayerLadderName ='" & CheckString(Request.Form("OldName")) & "'"
	else
		strSql = "select * from tbl_playerLadders where PlayerLadderName ='" & CheckString(Request.Form("LadderName")) & "'"
	end if
	oRs.Open strSQL, oConn
	if oRs.EOF and oRs.BOF then
		oRs.Close 
		jDate=date
		strSQL = "insert into tbl_playerLadders (PlayerLadderName, Active, Abbreviation, Locked, Challenge, GameID) values "
		strSQL = strSQL & "('" & CheckString(lName) & "','" & lactive & "','" & labbr & "', '" & llocked & "', '" & lChallenge & "','" & intGame  & "')"
		'response.write strSQL
'		Response.End
		oConn.Execute(strSQL)
	else
		thisladderid=ors.fields("PlayerLadderID").value
		ors.close
		strsql = "update tbl_Playerladders set PlayerLadderName='" & replace(lname, "'", "''") & _
			"', Active='" & lactive & "', Abbreviation='" & labbr & "', Locked='" & llocked & _
			"', Challenge='" & lchallenge & "', GameID='" & intGame  & "' where Playerladderid='" & thisladderid & "'"
		oConn.Execute(strSQL)
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/adminmenu.asp"
end if
'-----------------------------------------------
' Save team and ladder linkage
'-----------------------------------------------
if Request.Form("SaveType")="LadderJoin" then
	if not(IsSysAdmin()) and not(IsTeamFounderByID(request.form("TeamID"))) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQL="select * from lnk_T_L order by TLLinkID"
		oRs.Open strSQL, oConn
		if not (ors.eof and ors.bof) then
			do while not ors.eof
				lastlinkID = ors.fields(0).value
				ors.movenext
			loop
		end if
		newlinkID = lastlinkID + 1
		ors.close
			
		strSQL = "select TeamID, TeamFounderID, TeamName, TeamEmail from tbl_Teams where TeamID=" & Request.form("TeamID")
		oRs.Open strSQL, oConn
		if not (ors.eof and ors.bof) then
			do while not ors.eof
				teamname=ors.fields(2).value
				tid=ors.fields(0).value
				ownerid=ors.fields(1).value
				temail=ors.Fields(3).Value
				ors.movenext
			loop
		end if
		ors.close
		strSQL = "select LadderID, LadderName from tbl_Ladders where LadderName='" & CheckString(Request.Form("LadderToJoin")) & "'"
		oRs.Open strSQL, oConn
		if not (ors.eof and ors.bof) then
			do while not ors.eof
				lid=ors.fields(0).value
				lname = ors.fields(1).Value
				ors.movenext
			loop
		end if
		ors.close
		strSQL = "select rank from lnk_T_L where ladderID=" & lid & " AND Rank IS NOT NULL order by rank"
		oRs.Open strSQL, oConn
		rnk=0
		if not (ors.eof and ors.bof) then
			do while not ors.eof
				rnk=ors.fields(0).value
				ors.movenext
			loop
		end if
		rnk=rnk+1
		ors.close
		cleartojoin=true
		strsql="select * from lnk_T_P_L where playerid=" & ownerid
		ors.open strsql, oconn
		if not (ors.bof and ors.eof) then
			do while not ors.eof
				TLLinkIDJoined=ors.fields("TLLinkID").value
				strsql="select ladderid from lnk_T_L where tllinkid=" & TLLinkIDJoined
				ors2.open strsql,oconn
				if not (ors2.bof and ors2.eof) then
					if ors2.fields(0).value = lid then
						cleartojoin=false
						' Owner on another team for the same ladder, assuming owner can't quit his own team
					end if
				end if
				ors2.close
				ors.movenext
			loop
		end if
		ors.close
		if cleartojoin then
			'-------------
			' Ladder Join
			'-------------
			strSQL = "select count(*) from lnk_T_L where teamID=" & tid & " and ladderID=" & lid
			ors.open strSQL, oConn
			if ors.fields(0).value = 0 then
				strsql= "insert into lnk_T_L (TLLinkId, TeamID, LadderID, Rank, Status) values "
				strSQl = strSQL & "('" & newlinkId & "','" & tid & "','" & lid & "','" & rnk & "', 'Available')"'
				ors.close
				oConn.Execute(strSQL)
			else
				strsql= "update lnk_T_L set isactive=1, rank=" & rnk & ", wins=0, losses=0, forfeits=0 where teamid=" & tid & " and ladderid=" & lid
				ors.close
				oConn.Execute(strSQL)
			end if
		end if
		session("CurrentTeam") = teamname
		session("CurrentLadder") = Request.Form("LadderToJoin")
		'response.write "Joinable?" & cleartojoin
		if cleartojoin then
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "security.asp?SecType=founderjoin&team=" & server.urlencode(teamname) & "&ladder=" & server.urlencode(request.form("laddertojoin"))
		else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=1"
		end if
		'response.write ownerid
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	response.redirect "default.asp"
end if
'-----------------------------------------------
' Save team and ladder linkage
'-----------------------------------------------
if Request.Form("SaveType")="LeagueJoin" then
	if not(IsSysAdmin()) and not(IsTeamFounderByID(request.form("TeamID"))) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQL = "select TeamID, TeamFounderID, TeamName, TeamEmail from tbl_Teams where TeamID=" & Request.form("TeamID")
		oRs.Open strSQL, oConn
		if not (ors.eof and ors.bof) then
			do while not ors.eof
				teamname=ors.fields(2).value
				tid=ors.fields(0).value
				ownerid=ors.fields(1).value
				temail=ors.Fields(3).Value
				ors.movenext
			loop
		Else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=7"
		End If
		ors.close
		strSQL = "select l.LeagueID, l.LeagueName, ConferenceName, LeagueConferenceid from tbl_leagues l, tbl_league_conferences c where l.LeagueID = C.LeagueID AND LeagueConferenceID='" & CheckString(Request.Form("LeagueConferenceID")) & "'"
		oRs.Open strSQL, oConn
		if not (ors.eof and ors.bof) then
			do while not ors.eof
				LeagueID = ors.fields("LeagueID").value
				LeagueName = ors.fields("LeagueName").Value
				ConferenceName = oRs.Fields("ConferenceName").Value
				ConferenceID = ors.fields("LeagueConferenceid").value
				ors.movenext
			loop
		Else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=7"
		End If
		ors.close
		
		cleartojoin=true
		strsql="select lnkLeagueTeamID from lnk_league_team_player where playerid=" & ownerid
		ors.open strsql, oconn
		if not (ors.bof and ors.eof) then
			do while not ors.eof
				JoinedLinkID=ors.fields("lnkLeagueTeamID").value
				strsql="SELECT LeagueID FROM lnk_league_team WHERE lnkLeagueTeamID=" & JoinedLinkID & " AND LeagueID = '" & LeagueID & "' AND Active=1"
				ors2.open strsql,oconn
				if not (ors2.bof and ors2.eof) then
					cleartojoin=false
					' Owner on another team for the same ladder, assuming owner can't quit his own team
				end if
				ors2.nextrecordset
				ors.movenext
			loop
		end if
		ors.close
		if cleartojoin then
			strSQL = "EXECUTE LeagueTeamJoin @LeagueID = '" & leagueID & "', @TeamID='" & tid & "', @ConferenceID = '" & ConferenceID & "'"
			oConn.Execute(strSQL)			
			
'			Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
'			Mailer.RemoteHost  = "127.0.0.1"
'			Mailer.FromName    = "TWL: League Join Notification"
	'		mailer.FromAddress = "automailer@teamwarfare.com"
	'		Mailer.AddRecipient teamName, tEmail
	'		Mailer.Subject     = "TWL: " & teamName & " joins queue for " & LeagueName & " League in " & ConferenceName & " Conference"
	'		Text = teamName & ", your team has joined the waiting queue for the " & LeagueName  & " League in " & ConferenceName & " Conference. You will recieve a notification when you team is either accepted or declined entry into the conference." & vbcrlf
	'		Text = text & "This is an information e-mail only, please do not reply to this message."
	'		Mailer.BodyText    = text
	'		on error resume next
	'		Mailer.SendMail
	'		on error goto 0
	'		set mailer = nothing

			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "viewteam.asp?team=" & Server.URLEncode(teamName)
		else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=29"
		end if
		'response.write ownerid
	end if
end if
'-----------------------------------------------
' Save Player and ladder linkage
'-----------------------------------------------
if Request.Form("SaveType")="PlayerLadderJoin" then
	If not(IsSysAdmin()) and not(session("uName") = request("PlayerName")) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQL = "select PlayerLadderID, PlayerLadderName from tbl_PlayerLadders where PlayerLadderName='" & CheckString(Request.Form("LadderToJoin")) & "'"
		oRs.Open strSQL, oConn
		if not (ors.eof and ors.bof) then
			do while not ors.eof
				lid=ors.fields(0).value
				lname = ors.fields(1).Value
				ors.movenext
			loop
		Else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "/errorpage.asp?error=7"
		ENd If
		ors.close
		strSQL = "select Top 1 rank from lnk_p_pl where PlayerladderID=" & lid & " AND IsActive = 1 order by rank desc"
		oRs.Open strSQL, oConn
		rnk=0
		if not (ors.eof and ors.bof) then
			rnk=ors.fields(0).value
		end if
		If IsNull(Rnk) Then 
			rnk = 0
		End If
		rnk=rnk+1
		ors.close
		cleartojoin=true
		if cleartojoin then
			'-------------
			' Ladder Join
			'-------------
			strSQL = "select count(*) from lnk_p_pl where PlayerID=" & request("playerID") & " and PlayerladderID=" & lid
			ors.open strSQL, oConn
			if ors.fields(0).value = 0 then
				strsql= "insert into lnk_p_pl (PlayerID, PlayerLadderID, Rank, Status) values "
				strSQl = strSQL & "('" & Request("PlayerID") & "','" & lid & "','" & rnk & "', 'Available')"'
				ors.close
				ors.open strSQL, oConn
			else
				strsql= "update lnk_p_pl set isactive=1, rank=" & rnk & ", wins=0, losses=0, forfeits=0 where PlayerID=" & request("PlayerID") & " and Playerladderid=" & lid
				ors.close
				ors.Open strsql, oconn
			end if
		end if
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	response.redirect "viewplayer.asp?player=" & request("Playername")
end if
'-----------------------------------------------
' Add Comms
'-----------------------------------------------
if Request.form("SaveType") = "Add_Communications" then
	strSQL="INSERT INTO tbl_Comms ( matchID, CommDate, CommAuthor, Comms ) values ('" 
	strSQL = strSQL & Request.form("matchID") & "',GetDate(),'"  & replace(Request.Form("commauthor"), "'", "''") & "','"
	strSQL = strSQL & replace(Request.Form("comms"), "'", "''") & "')"
	oConn.Execute(strSQL)
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/TeamLadderAdmin.asp?ladder=" & server.urlencode(request("Ladder")) & "&team=" & server.urlencode(request("Team"))
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
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL= "update tbl_comms set Comms='" & replace(Request.Form("comms"), "'", "''") &  "' where commid=" & Request.Form("commid")
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/TeamLadderAdmin.asp?ladder=" & server.urlencode(request("Ladder")) & "&team=" & server.urlencode(request("Team"))
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
		response.redirect "errorpage.asp?error=3"
	end if
	if Request.QueryString("commid") <> "" then
		strSQL= "delete from tbl_comms where commid=" & Request.QueryString("commid")
		'Response.Write strSQl
		oRs.Open strSQL, oConn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect"TeamLadderAdmin.asp?ladder=" & server.urlencode(request("ladder")) & "&team=" & server.urlencode(request("team"))
end if
'-----------------------------------------------
' Add player Comms
'-----------------------------------------------
if Request.form("SaveType") = "playerAdd_Communications" then

	strSQL="INSERT INTO tbl_playerComms ( playermatchID, CommDate, CommAuthor, Comms, CommTime ) values ('" 
	strSQL = strSQL & Request.form("matchID") & "','" & Request.Form("commdate") & "','"  & replace(Request.Form("commauthor"), "'", "''") & "','"
	strSQL = strSQL & replace(Request.Form("comms"), "'", "''") & "',GetDate())"
	oConn.EXECUTE (strSQL)
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/PlayerLadderAdmin.asp?ladder=" & server.urlencode(request("Ladder")) & "&player=" & server.urlencode(request("player"))
end if
'-----------------------------------------------
' Edit player Comms
'-----------------------------------------------
if Request.form("SaveType") = "playerEdit_Communications" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL= "update tbl_playercomms set Comms='" & replace(Request.Form("comms"), "'", "''") &  "', CommDate='" & date & "', CommTime='" & time & "' where commid=" & Request.Form("commid")
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/PlayerLadderAdmin.asp?ladder=" & server.urlencode(request("Ladder")) & "&player=" & server.urlencode(request("player"))
end if
'-----------------------------------------------
' Delete player Comms
'-----------------------------------------------
if Request.QueryString("SaveType") = "playerDelete_Communications" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	if Request.QueryString("commid") <> "" then
		strSQL= "delete from tbl_playercomms where commid='" & Request.QueryString("commid") & "'"
		'Response.Write strSQl
		oRs.Open strSQL, oConn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/PlayerLadderAdmin.asp?ladder=" & server.urlencode(request("Ladder")) & "&player=" & server.urlencode(request("player"))
end if
'-----------------------------------------------
' Submit Challenge
'-----------------------------------------------
if Request.QueryString("SaveType") = "challenge" then
	
	if not(IsSysAdmin()) and not(IsLadderAdmin(request.querystring("ladder"))) and not(IsTeamFounder(request.querystring("team"))) and not(IsTeamCaptain(request.querystring("team"), request.querystring("ladder"))) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	
	' first verify there is no other challenge
	strSQL = "SELECT lnk_T_L.Status FROM tbl_Teams INNER JOIN (tbl_Ladders INNER JOIN lnk_T_L ON "
	strSQL = strSQL & "tbl_Ladders.LadderID = lnk_T_L.LadderID) ON tbl_Teams.TeamID = lnk_T_L.TeamID WHERE "
	strSQL = strSQL & "(((tbl_Ladders.LadderName)='" & CheckString(Request.QueryString("ladder")) & "') AND lnk_t_l.IsActive=1 and ((tbl_Teams.TeamName)='" & replace(Request.QueryString("opponent"), "'", "''")  & "'));"
	ors.open strSQL, oConn
	if not (ors.eof and ors.bof) then
		if (ors.fields(0).value = "Available" or left(ors.Fields(0).Value, 8) = "Defeated") then
			cValid = true
		end if
	end if
	ors.close
	
	map1 = Request.QueryString("map1")
	map2 = Request.QueryString("map2")
	map3 = Request.QueryString("map3")
	map4 = Request.QueryString("map4")
	map5 = Request.QueryString("map5")
	If Len(map1) = 0 Then
		map1 = "TBD"
	End if
	If Len(map2) = 0 Then
		map2 = "TBD"
	End if
	If Len(map3) = 0 Then
		map3 = "TBD"
	End if
	If Len(map4) = 0 Then
		map4 = "TBD"
	End if
	If Len(map5) = 0 Then
		map5 = "TBD"
	End if
	
	if cValid then
		cValid=False
		strSQL = "SELECT lnk_T_L.Status FROM tbl_Teams INNER JOIN (tbl_Ladders INNER JOIN lnk_T_L ON "
		strSQL = strSQL & "tbl_Ladders.LadderID = lnk_T_L.LadderID) ON tbl_Teams.TeamID = lnk_T_L.TeamID WHERE "
		strSQL = strSQL & "(((tbl_Ladders.LadderName)='" & CheckString(Request.QueryString("ladder")) & "') AND lnk_t_l.IsActive=1 and ((tbl_Teams.TeamName)='" & replace(Request.QueryString("team"), "'", "''")  & "'));"
		ors.open strSQL, oConn
		if not (ors.eof and ors.bof) then
			if (ors.fields(0).value = "Available" or left(ors.Fields(0).Value,6) = "Immune") then
				cValid = true
			end if
		end if
		ors.close
		if cValid = true then 
			strSQL = "select LadderID from tbl_Ladders where LadderName='" & checkstring(Request.QueryString("ladder")) & "'"
			ors.open strSQL, oconn
			if not (ors.eof and ors.bof) then
				mLadder = ors.fields(0).value
			end if
			ors.close
			'strSQL = "select MatchID from tbl_Matches order by matchID desc"
			'ors.open strSQL, oconn
			dim matchID
			'if not (ors.eof and ors.bof) then
'				do while not ors.eof
			'		matchID = ors.fields(0).value + 1
'					ors.movenext
'				loop
			'end if
			'ors.close
			strSQL = "SELECT [lnk_T_L].[TLLinkID] FROM tbl_Ladders INNER JOIN "
			strSQL = strSQL & "(tbl_Teams INNER JOIN lnk_T_L ON [tbl_Teams].[TeamID]=[lnk_T_L].[TeamID]) ON [tbl_Ladders].[LadderID]=[lnk_T_L].[LadderID] "
			strSQL = strSQL & "WHERE ((([tbl_Teams].[TeamName])='" & replace(Request.QueryString("opponent"), "'", "''") & "') And (([tbl_Ladders].[LadderName])='" & checkstring(Request.QueryString("ladder")) & "'));"
			ors.open strSQL, oconn
			if not (ors.eof and ors.bof) then
				oID = ors.fields(0).value
			end if
			ors.close
			strSQL = "SELECT [lnk_T_L].[TLLinkID] FROM tbl_Ladders INNER JOIN "
			strSQL = strSQL & "(tbl_Teams INNER JOIN lnk_T_L ON [tbl_Teams].[TeamID]=[lnk_T_L].[TeamID]) ON [tbl_Ladders].[LadderID]=[lnk_T_L].[LadderID] "
			strSQL = strSQL & "WHERE ((([tbl_Teams].[TeamName])='" & replace(Request.QueryString("team"), "'", "''") & "') And (([tbl_Ladders].[LadderName])='" & checkstring(Request.QueryString("ladder")) & "'));"
			ors.open strSQL, oconn
			if not (ors.eof and ors.bof) then
				tID = ors.fields(0).value
			end if
			ors.close
			iMap = "TBD"
			strSQL = "insert into tbl_Matches ( "
			strSQL = strSQL & " MatchDefenderID, MatchAttackerID,MatchDate, MatchChallengeDate, "
			strSQL = strSQL & " MatchMap1ID, MatchMap2ID, MatchMap3ID, MatchMap4ID, MatchMap5ID, MatchLadderID "
			strSQL = strSQL & " ) VALUES ( "
			strSQL = strSQL & " '" & oID & "','" & tID & "','TBD','" & now & "', "
			strSQL = strSQL & " '" & map1 & "','" & map2 & "','" & map3 &  "','" & map4 &  "','" & map5 &  "', "
			strSQL = strSQL & " '" & mLadder & "')"
			oConn.Execute(strSQL)
			strSQL = "update lnk_T_L set Status='Defending', LadderMode=0, ModeFlagTime=null where TLLinkID=" & oID
			oConn.Execute(strSQL)
			strSQL = "update lnk_T_L set Status='Attacking', LadderMode=0, ModeFlagTime=null where TLLinkID=" & tID
			oConn.Execute(strSQL)
			
			strsql = "Select TeamID, teamName, TeamEmail from tbl_teams where teamname='" & replace( Request.querystring("team"), "'", "''") & "'"
			ors.Open strsql, oconn
				aname = ors.Fields(1).Value
				aemail = ors.Fields(2).Value
				aid = ors.Fields(0).Value
			ors.Close
	
			strsql = "Select TeamID, teamName, TeamEmail from tbl_teams where teamname='" & replace( Request.querystring("opponent"), "'", "''") & "'"
			ors.Open strsql, oconn
				dname = ors.Fields(1).Value
				demail = ors.Fields(2).Value
				did = ors.Fields(0).Value
			ors.Close
			lname = Request.QueryString("ladder") & " Ladder"
'			Subject     = "TWL: " & lname & ": " & aName & " attacks " & dname
'			Text = aName & ", your team has attacked " & dname & " on the " & lname & " Ladder." & vbcrlf
'			text = text & "Your opponent should respond within 48 hours of this e-mail." & vbcrlf
'			Text = text & "This is a confirmation e-mail, please do not reply to this message."
'			if not(MailTeamCaptains(tid, text, subject, true, false, mladder)) then
'				'Response.Write "Mail not sent to attacker"
'			end if
			'Response.Write "Mail Sent"
			Subject     = "TWL: " & lname & ": " & dName & " defends against " & aname
			Text = dName & ", your team has been attacked by " & aname & " on the " & lname & " Ladder." & vbcrlf
			test = text & "You have 48 hours to respond to this challenge by logging into the Teamwarfare page and offering two different dates and times." & vbcrlf			
			Text = text & "This is a confirmation e-mail, please do not reply to this message."
			if not(MailTeamCaptains(oid, text, subject, false, false, mladder)) then
				'Response.Write "Mail not sent to defender"
			end if
			'Response.Write "Mail Sent"			
			'Response.write text
		else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=6"
		end if
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=4"
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "TeamLadderAdmin.asp?ladder=" & server.urlencode(request.querystring("ladder")) & "&team=" & server.urlencode(request.querystring("team"))
end if
'-----------------------------------------------
' Delete News
'-----------------------------------------------
if Request.QueryString("SaveType") = "DeleteNews" then
	if not(IsSysAdmin()) and not(IsAnyLadderAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQL =  "delete from tbl_News where newsID=" & Request.QueryString("newsid")
		ors.open strSQL, oconn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "newsdesk.asp"
end if
'-----------------------------------------------
' Edit News
'-----------------------------------------------
if Request.Form("SaveType") = "EditNews" then
	if not(IsSysAdmin()) and not(IsAnyLadderAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQL =  "update tbl_News set NewsType='" & Request.Form("NewsType") & "', NewsContent='" & replace(Request.Form("newscontent"), "'", "''") & "', NewsHeadline = '" & CheckString(Request.Form("Headline")) & "' where newsID=" & Request.form("newsid")
		oconn.execute(strSQL)
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "newsdesk.asp"
end if
'-----------------------------------------------
' Add News
'-----------------------------------------------
if Request.Form("SaveType") = "AddNews" then
	if not(IsSysAdmin()) and not(IsAnyLadderAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQL = "select NewsID from tbl_news order by NewsID desc"
		ors.open strSQL, oconn
		if not (ors.eof and ors.bof) then
			newidnum=ors.fields(0).value + 1
		end if
		strSQL =  "insert into tbl_news(NewsID, NewsDate, NewsHeadline, NewsAuthor, NewsType, NewsContent) values "
		strSQL = strSQL & "('" & newidnum & "','" & now & "','" & replace(Request.Form("headline"), "'", "''") & "','" & replace(session("uName"), "'", "''") & "','" & Request.Form("NewsType") & "','" & replace(Request.Form("content"), "'", "''") & "')"
		ors.close
		ors.open strSQL, oconn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "newsdesk.asp"
end if
'-----------------------------------------------
' Accept Match
'-----------------------------------------------
if Request.Form("SaveType") = "AcceptMatch" then
	if not(IsSysAdmin()) and not(IsLadderAdmin(request.form("ladder"))) and not(IsTeamFounder(request.form("team"))) and not(IsTeamCaptain(request.form("team"), request.form("ladder"))) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		day1 = Request.Form("Day1")
		day2 = Request.Form("day2")
		intMatchDays = Request.Form("MD")
		If IsNumeric(intMatchDays) Then
			intMatchDays = cint(intMatchDays)
		Else
			intMatchDays = -1
		End If
		bump=0
		if Request.Form("scApproved") = "true" then
			bump=1
		end if
		mdate1 = day1
		mDate1=left(mDate1, len(mDate1)-4)
		loc=instr(1,mDate1,",")+1
		mdate1=mid(mdate1, loc+1, len(mdate1)-loc)
		pubdate1=datevalue(mDate1)
		pubtime1=timevalue(mDate1)
		mdate2 = day2
		mdate2=left(mdate2, len(mdate2)-4)
		loc=instr(1,mdate2,",")+1
		mdate2=mid(mdate2, loc+1, len(mdate2)-loc)
		pubdate2=datevalue(mdate2)
		pubtime2=timevalue(mdate2)
'		Response.Write "<BR>Time1 - " & pubtime1
'		Response.Write "<BR>Date1 - " & pubdate1
'		Response.Write "<BR>Time2 - " & pubtime2
'		Response.Write "<BR>Date2 - " & pubdate2
'		Response.End 
		if Not((intMatchDays = 2^vbSunday) _
					OR (intMatchDays = 2^vbMonday) _
					OR (intMatchDays = 2^vbTuesday) _
					OR (intMatchDays = 2^vbWednesday) _
					OR (intMatchDays = 2^vbThursday) _
					OR (intMatchDays = 2^vbFriday) _
					OR (intMatchDays = 2^vbSaturday)) Then 
					
			if (pubdate1 = pubdate2) then
				oConn.Close 
				Set oConn = Nothing
				Set oRs = Nothing
				Response.Clear
				Response.Redirect "errorpage.asp?error=10"
				'error
			end if
		End If
		if (pubtime1 = pubtime2) then
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "errorpage.asp?error=12"
			'error
		end if

		' Update the record, and mark the maps that need marking
		strMapConfiguration = Request.Form("MC")
		strSQL =  "update tbl_Matches set MatchAwaitingForfeit=0, "
		strSQL = strSQL & " forfeitreason='xxx', "
		strSQL = strSQL & " MatchSelDate1='" & Request.Form("Day1") & "', "
		strSQL = strSQL & " MatchSelDate2='" & Request.Form("day2") & "', "
		For i = 1 TO Len(strMapConfiguration)
			If Mid(strMapConfiguration, i, 1) = "D" Then
				strSQL = strSQL & " MatchMap" & i & "ID='" & CheckString(Request.Form("Map" & i)) & "', "
			End IF
		Next
		strSQL = strSQL & " MatchAcceptanceDate='" & now & "', "
		strSQL = strSQL & " Cast=Cast + " & bump & " "
		strSQL = strSQL & " where MatchID=" & Request.Form("matchid")
		ors.open strSQL, oconn

		strsql = "SELECT tbl_Ladders.LadderName, tbl_Teams.TeamName, lnk_T_L.TLLinkID, tbl_Matches.MatchID, tbl_ladders.LadderID"
		strsql = strsql & " FROM tbl_Teams INNER JOIN ((lnk_T_L INNER JOIN tbl_Ladders ON lnk_T_L.LadderID = tbl_Ladders.LadderID) INNER JOIN tbl_Matches ON lnk_T_L.TLLinkID = tbl_Matches.MatchDefenderID) ON tbl_Teams.TeamID = lnk_T_L.TeamID"
		strsql = strsql & " WHERE (((tbl_Matches.MatchID)=" & request.form("matchID") & "));"
		ors.Open strsql, oconn
		if not(ors.EOF and ors.BOF) then
			lname = ors.Fields(0).Value
			dname = ors.Fields(1).Value
			did = ors.Fields(2).Value
			lid = ors.Fields(3).Value
		end if
		ors.Close
		strsql = "SELECT tbl_Ladders.LadderName, tbl_Teams.TeamName, lnk_T_L.TLLinkID, tbl_Matches.MatchID"
		strsql = strsql & " FROM tbl_Teams INNER JOIN ((lnk_T_L INNER JOIN tbl_Ladders ON lnk_T_L.LadderID = tbl_Ladders.LadderID) INNER JOIN tbl_Matches ON lnk_T_L.TLLinkID = tbl_Matches.MatchAttackerID) ON tbl_Teams.TeamID = lnk_T_L.TeamID"
		strsql = strsql & " WHERE (((tbl_Matches.MatchID)=" & request.form("matchID") & "));"
		ors.Open strsql, oconn
		if not(ors.EOF and ors.BOF) then
			aname = ors.Fields(1).Value
			aid = ors.Fields(2).Value
		end if
		ors.Close
'		Subject     = "TWL: " & lname & ": " & dName & " choose dates..."
'		Text = aName & ", your team has attacked " & dname & " on the " & lname & " Ladder." & vbcrlf
'		Text = text & "Please login to the teamwarfare admin panel for your team and accept a proposed date." & vbcrlf
'		Text = text & "This is a confirmation e-mail, please do not reply to this message." & vbcrlf
'		if not(MailTeamCaptains(aid, text, subject, true, false, lid)) then
'			'Response.Write "Mail not sent to attacker"
'		end if
'		'Response.Write "Mail Sent"
'		Subject     = "TWL: " & lname & ": " & dName & ", you chose dates against " & aname
'		dtext = dName & ", your team has been attacked by " & aname & " on the " & lname & " Ladder." & vbcrlf
'		dtext = dtext & "You have done all you need to do until the opposing team chooses a date. " & vbcrlf
'		dtext = dtext & "This is a confirmation e-mail, please do not reply to this message." & vbcflr
'		if not(MailTeamCaptains(did, dtext, subject, true, false, lid)) then
'			'Response.Write "Mail not sent to defender"
'		end if
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "TeamLadderAdmin.asp?ladder=" & Server.URLEncode(Request.form("ladder")) & "&team=" & Server.URLEncode(Request.Form("team"))
end if
'-----------------------------------------------
' Accept Match Date
'-----------------------------------------------
if Request.FORM("SaveType") = "AcceptMatchDate" then
	if not(IsSysAdmin()) and not(IsLadderAdmin(Request("Ladder"))) and not(IsTeamFounder(Request("Team"))) and not(IsTeamCaptain(Request("Team"), Request("Ladder"))) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQL = "SELECT MatchDate FROM tbl_matches WHERE MatchID = '" & Request.Form("MatchID") & "'"
		oRs.Open strSQl, oConn
		If Not(oRs.EOF AND oRS.BOF) Then
			mDate = oRs.Fields("MatchDate").Value
		End If
		oRs.NextRecordSet
		If mDate = "TBD" Then
			bump=-1
			if Request.Form("scApproved") = "true" then
				bump=1
			end if
			mDate=Request.Form("matchdate")
	
			strMapConfiguration = Request.Form("MC")
			strSQL =  "update tbl_Matches set MatchAwaitingForfeit=0, "
			strSQL = strSQL & " forfeitreason='xxx', "
			For i = 1 TO Len(strMapConfiguration)
				If Mid(strMapConfiguration, i, 1) = "A" Then
					strSQL = strSQL & " MatchMap" & i & "ID='" & Request.Form("Map" & i) & "', "
				End IF
			Next
			strSQL = strSQL & " MatchDate='" & mDate & "', "
			strSQL = strSQL & " MatchLockDate=GetDate(), "
			strSQL = strSQL & " Cast=Cast + " & bump & " "
			strSQL = strSQL & " where MatchID=" & Request.form("matchid")
			oConn.Execute (strSQL)
			
			' Do random map selection here
			Dim MapList
			randomize
			For i = 1 TO Len(strMapConfiguration)
				If Mid(strMapConfiguration, i, 1) = "R" Then
					' Here is where we select random maps
					MapList = ""
					strSQL = "EXEC GetMapList '" & Request.Form("MatchID") & "', " & i
					ors.open strSQL, oconn
					If Not(oRS.EOF AND oRS.BOF) Then
						Do While Not(oRs.EOF)
							If Len(MapList) > 0 Then
								MapList = MapList & "Þ" & oRS.Fields("MapName").Value
							Else
								MapList = oRS.Fields("MapName").Value
							End If
							oRS.MoveNext 
						Loop
					End If
					oRs.NextRecordset 
					MapArray = Split(MapList, "Þ")		
	
					rNum1 = Int((uBound(MapArray) - 0 + 1) * Rnd + 0)
					If rNum1 > uBound(MapArray) Then
						' Too High
						rNum1 = uBound(MapArray)
					ElseIf rNum1 < 0 Then
						' Too Low
						rNum1 = 0
					End If
					
					Map = MapArray(rNum1)
					
					strSQL = "UPDATE tbl_matches SET MatchMap" & i & "ID='" & CheckString(Map) & "' WHERE MatchID = " & Request.Form("MatchID")
					oConn.Execute(strSQL)
				End If
			Next
	
			strSQL="select matchladderid from tbl_matches where MatchID=" & Request.form("matchid")
			ors.Open strSQL, oconn
			if not (ors.EOF and ors.BOF) then
				matchladderid=ors.Fields(0).Value 
			end if
			ors.NextRecordset 
			
			Dim arrArrays(20)
			For iCounter = 0 To uBound(arrArrays)
				arrArrays(iCounter) = ""
			Next
			iCounter = -1
			CurrentItem = ""
			strSQL = "SELECT * FROM vLadderOptions WHERE SelectedBy = 'R' AND LadderID='" & matchladderid & "'"
			oRs.Open strSQL, oConn
			If Not(oRS.BOF AND ors.EOF) Then
				Do While Not(oRs.EOF)
					If CurrentItem <> oRS.Fields("OptionID").Value  Then
						iCounter = iCounter + 1
						CurrentItem = oRS.Fields("OptionID").Value 
						arrArrays(iCounter) = CurrentItem
					End If
					arrArrays(iCounter) = arrArrays(iCounter) & "|" & oRS.Fields("OptionValueID").Value
					oRS.MoveNext 
				Loop
				strSQL = ""
				For iCounter = 0 To 20
					If Len(arrArrays(iCounter)) > 0 Then
						Randomize
						currentArray = Split(arrArrays(iCounter), "|")
					
						currentID = currentArray(0)
						rNum = Int((uBound(currentArray) - 1 + 1) * Rnd + 1)
						If rNum > uBound(CurrentArray) Then
							' Too High
							rNum = uBound(CurrentArray)
						ElseIf rNum < 1 Then
							' Too Low
							rNum = 1
						End If
						strSQL = strSQL & "EXEC usp_AddMatchOption '" & Request.Form("MatchID") & "', '" & currentID & "', '" & currentArray(rNum) & "';"
					End If
				Next
				If Len(strSQL) > 0 Then
					oConn.Execute (strSQL)
				End If
			End IF
			oRS.NextRecordset 
			
			strSQL="select MatchAttackerID, MatchDefenderID, MatchLadderID from tbl_matches where MatchID=" & Request.form("matchid")
			ors.Open strSQL, oconn
			if not (ors.EOF and ors.BOF) then
				mAID=ors.Fields(0).Value 
				mDID=ors.Fields(1).Value
				mLID=ors.Fields(2).Value
			end if
			ors.Close
			if (mAID <> "" and mDID <> "" and mLID <> "") then
				strSQL="select teamname, teamtag from tbl_teams inner join lnk_T_L on lnk_T_L.teamid=tbl_teams.teamid where TLLinkID=" & mAID & " and lnk_T_L.ladderid=" & mLID
				ors.Open strsql, oconn
				if not (ors.EOF and ors.BOF) then
					mAName=ors.Fields(0).Value
					mATag=ors.Fields(1).Value 
				end if
				ors.Close 
				strSQL="select teamname, teamtag from tbl_teams inner join lnk_T_L on lnk_T_L.teamid=tbl_teams.teamid where TLLinkID=" & mDID & " and lnk_T_L.ladderid=" & mLID
				ors.Open strsql, oconn
				if not (ors.EOF and ors.BOF) then
					mDName=ors.Fields(0).Value
					mDTag=ors.Fields(1).Value 
				end if
				ors.Close 
				strSQL="select laddername, ladderabbreviation from tbl_ladders where ladderid=" & mLID
				ors.Open strsql, oconn
				if not (ors.EOF and ors.BOF) then
					mLName=ors.Fields(0).Value
					mLAbbr=ors.fields(1).value
				end if
				ors.Close
				mattrank=-1
				mdefrank=-1
				strsql="Select rank from lnk_T_L where tllinkID=" & maid
				ors.Open strsql, oconn
				if not (ors.EOF and ors.BOF) then
					mattrank=ors.Fields(0).Value
				end if
				ors.Close
				strsql="Select rank from lnk_T_L where tllinkID=" & mDID
				ors.Open strsql, oconn
				if not (ors.EOF and ors.BOF) then
					mdefrakn=ors.Fields(0).Value
				end if
				ors.Close
				
				mDate=left(mdate, len(mDate)-4)
				loc=instr(1,mdate,",")+1
				mdate=mid(mdate, loc+1, len(mdate)-loc)
				pubdate=datevalue(mdate)
				pubtime=timevalue(mdate)
				dTime="" & pubtime
				if left(ltrim(dtime),1) <> "1" then
					dtime="0" & dtime
				end if
				strsql="select count(*) from tbl_disp_pending where mmID='" & request.form("matchID") & "'"
				ors.Open strsql,oconn
				if ors.Fields(0).Value = 0 then
					strSQL="insert into tbl_disp_pending(mDate, mTime, mAtt, mDef, mLadder, mATag, mDTag, mmID, mdefrank, mattrank, mLadderAbbreviation) values ('" & PubDate & "','" & dtime & "','" & replace(mAName,"'","''") & "','" & replace(mDName,"'","''") & "','" & replace(mLName, "'", "''") & "','" & replace(mATag,"'","''") & "','" & replace(mDTag,"'","''") & "'," & Request.form("matchid") & ", " & mdefrank & ", " & mattrank & ", '" & mLAbbr & "')"
					oconn.Execute strsql
				end if
				ors.Close 
				'Response.Write strsql
			
			lname = MLName
			dname = mDName
			aname = mAname
			aid=mAid
			did=mDID
			lid=MLID
			Subject     = "TWL: " & lname & ": " & dName & " match scheduled ."
	'		Text = aName & ", your team has attacked " & dname & " on the " & lname & " Ladder." & vbcrlf
	'		Text = text & "Please login to the teamwarfare admin panel for your team to confirm side choices." & vbcrlf
	'		Text = text & "This is a confirmation e-mail, please do not reply to this message." & vbcrlf
	'		if not(MailTeamCaptains(aid, text, subject, true, false, lid)) then
	'			'Response.Write "Mail not sent to attacker"
	'		end if
			'Response.Write "Mail Sent"
	'		Subject     = "TWL: " & lname & ": " & aName & " scheduled match."
	'		Text = dName & ", your team has been attacked by " & aname & " on the " & lname & " Ladder." & vbcrlf
	'		Text = text & "Please login to the teamwarfare admin panel for your team to confirm side choices." & vbcrlf
	'		Text = text & "This is a confirmation e-mail, please do not reply to this message." & vbcrlf
	'		if not(MailTeamCaptains(did, text, subject, true, false, lid)) then
	'			'Response.Write "Mail not sent to defender"
	'		end if
	
			end if
		End If
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "TeamLadderAdmin.asp?ladder=" & server.urlencode(Request("Ladder")) & "&team=" & server.urlencode(Request("Team"))
		
end if
'-----------------------------------------------
' Match Reporting
'-----------------------------------------------
if Request.Form("SaveType") = "ReportMatch" then
	matchid=Request.Form("matchid")
	matchwinnername=Request.Form("matchwinnername")
	matchlosername=Request.Form("matchlosername")
	matchwinnerid=Request.Form("matchwinnerid")
	matchloserid=Request.Form("matchloserid")
	matchdate=Request.Form("matchdate")
	matchwinnerdefending=Request.Form("matchwinnerdefending")
	ladderid=Request.Form("ladderid")

'	if not(IsSysAdmin() or IsLadderAdminByID(ladderid) or IsTeamFounderByID(matchwinnerID) or IsTeamCaptainByID(matchwinnerID, ladderID) or IsTeamFounderByID(matchloserID) or IsTeamCaptainByID(matchloserID, ladderID))
	if not(IsSysAdmin() or IsLadderAdminByID(ladderid) or IsTeamFounder(matchlosername) or IsTeamCaptainByLinkID(matchloserID)) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
		'response.write (IsSysAdmin() or IsLadderAdminByID(ladderid) or IsTeamFounderByID(matchloserID) or IsTeamCaptainByID(matchloserID, ladderID))
	else
		strSQL = "select matchid from tbl_matches where matchid=" & matchid
		ors.open strSQL, oconn
		if (ors.eof and ors.bof) Then
			Response.Redirect "viewteam.asp?team=" & server.URLEncode(matchlosername)
		End If
		oRs.NextRecordSet

		strSQL = "select lnk_T_L.TLLinkID, lnk_T_L.rank from lnk_T_L inner join tbl_Teams on tbl_teams.TeamID=lnk_T_L.TeamID where (tbl_Teams.TeamName='" & replace(matchwinnername, "'", "''") & "') and (lnk_T_L.ladderID=" & ladderID & ")"
		ors.open strSQL, oconn
			'Response.Write strsql & "<br><br>"
		if not (ors.eof and ors.bof) then
			winnerid=ors.fields(0).value
			winnerrank=ors.fields(1).value
		end if
		ors.close
		strSQL = "select lnk_T_L.TLLinkID, lnk_T_L.rank from lnk_T_L inner join tbl_Teams on tbl_teams.TeamID=lnk_T_L.TeamID where (tbl_Teams.TeamName='" & replace(matchlosername, "'", "''") & "') and (lnk_T_L.ladderID=" & ladderID & ")"
		ors.open strSQL, oconn
			'Response.Write strsql & "<br><br>"
		if not (ors.eof and ors.bof) then
			loserid=ors.fields(0).value
			loserrank=ors.fields(1).value
		end if
		ors.close
		strSQL = "insert into tbl_History (MatchID, MatchWinnerID, MatchLoserID, MatchMap1, MatchMap2, MatchMap3, MatchMap4, MatchMap5, "
		strSQL = strSQL & "MatchMap1DefenderScore, MatchMap2DefenderScore, MatchMap3DefenderScore, MatchMap4DefenderScore, MatchMap5DefenderScore, "
		strSQL = strSQL & "MatchMap1AttackerScore,MatchMap2AttackerScore, MatchMap3AttackerScore, MatchMap4AttackerScore, MatchMap5AttackerScore, "
		strSQL = strSQL & "MatchMap1OT, MatchMap2OT, MatchMap3OT, MatchMap4OT, MatchMap5OT, "
		strSQL = strSQL & "MatchMap1forfeit, MatchMap2forfeit, MatchMap3forfeit, MatchMap4forfeit, MatchMap5forfeit, "
		strSQL = strSQL & "MatchDate, MatchWinnerDefending, MatchLadderID) values "
		strSQL = strSQL & "(" & matchid & ", " & winnerid & "," & loserid & ",'" & CheckString(Request.Form("Map1")) & "','" & CheckString(Request.Form("Map2")) & "','" & CheckString(Request.Form("Map3")) & "','" & CheckString(Request.Form("Map4")) & "','" & CheckString(Request.Form("Map5"))
		strSQL = strSQL & "','" & CheckString(Request.Form("Map1DefScore")) & "','" & CheckString(Request.Form("Map2DefScore")) & "','" & CheckString(Request.Form("Map3DefScore")) & "','" & CheckString(Request.Form("Map4DefScore")) & "','" & CheckString(Request.Form("Map5DefScore"))
		strSQL = strSQL & "','" & CheckString(Request.Form("Map1AttScore")) & "','" & CheckString(Request.Form("Map2AttScore")) & "','" & CheckString(Request.Form("Map3AttScore")) & "','" & CheckString(Request.Form("Map4AttScore")) & "','" & CheckString(Request.Form("Map5AttScore"))
		strSQL = strSQL & "','" & CheckString(Request.Form("intMap1OT")) & "','" & CheckString(Request.Form("intMap2OT")) & "','" & CheckString(Request.Form("intMap3OT")) & "','" & CheckString(Request.Form("intMap4OT")) & "','" & CheckString(Request.Form("intMap5OT"))
		strSQL = strSQL & "','" & CheckString(Request.Form("intMap1Forfeit")) & "','" & CheckString(Request.Form("intMap2Forfeit")) & "','" & CheckString(Request.Form("intMap3Forfeit")) & "','" & CheckString(Request.Form("intMap4Forfeit")) & "','" & CheckString(Request.Form("intMap5Forfeit"))
		strSQL = strSQL  & "','" & CheckString(matchdate) & "','" & CheckString(matchwinnerdefending) & "','" & CheckString(ladderID) & "')" 
'		Response.Write strSQL & "<br><br>"
'		Response.End 
		ors.open strSQL, oconn
			'Response.Write strsql & "<br><br>"
		strSQL = "delete from tbl_matches where matchid=" & matchid
		ors.open strSQL, oconn
				'Response.Write strsql & "<br><br>"
		strSQL = "update lnk_T_L set status='Defeated by " & replace(matchwinnername, "'", "''") & "', WinStreak = 0, LossStreak = LossStreak + 1, Losses=(Losses+1), LadderMode='1', ModeFlagTime='" & now & "' where TLLinkID=" & loserid
		ors.open strSQL, oconn
				'Response.Write strsql & "<br><br>"
		pstat="Immune until " & formatdatetime(dateadd("h",1,now),3)
			'Response.Write pstat
		strSQL = "update lnk_T_L set Status='" & pstat & "', wins=(wins+1), LadderMode=3, WinStreak = WinStreak + 1, LossStreak = 0, ModeFlagtime='" & now & "' where TLLinkID=" & winnerid
		ors.open strSQL, oconn
	
		'Response.Write strsql & "<br><br>"
		if matchwinnerdefending = "True" then
			'defender wins
		else
			strSQL = "update lnk_T_L set rank=(rank + 1) where (rank >" & (loserrank - 1) & " and rank < " & winnerrank & ") and isactive=1 and ladderID = " & ladderID
			'Response.Write strsql & "<br><br>"
			ors.open strSQL, oconn
			strSQL = "update lnk_T_L set rank=" & loserrank & " where TLLinkID=" & winnerID & " and ladderID=" & ladderID
			ors.open strSQL, oconn
			'Response.Write strsql
		end if
		strSQL="update tbl_Comms set CommDead=1 where matchid=" & matchid
		ors.open strSQL, oconn
		strSQl="delete from tbl_disp_pending where mmID=" & matchid
		'Response.Write strsql
		oconn.Execute strSQL
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "viewteam.asp?team=" & server.URLEncode(matchlosername)
end if
'-----------------------------------------------
' Player Submit Challenge
'-----------------------------------------------
if Request.QueryString("SaveType") = "playerchallenge" then
	PlayerName = request("Player")
	if Not(IsSysAdmin()) AND Not(IsPlayerLadderAdmin(request("Ladder"))) AND NOT(Session("uName") = playerName) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	
	' first verify there is no other challenge
	strSQL = "SELECT lnk.Status, lnk.PPLLinkID, p.PlayerHandle, p.PlayerEmail "
	strSQL = strSQL & " FROM tbl_players p, lnk_p_pl lnk, tbl_playerladders l "
	strSQL = strSQL & " WHERE l.PlayerLadderID = lnk.PlayerLadderID "
	strSQL = strSQL & " AND p.PlayerID = lnk.PlayerID "
	strSQL = strSQL & " AND l.PlayerLadderName = '"	& Checkstring(Request.QueryString("ladder")) & "' "
	strSQL = strSQL & " AND lnk.IsActive = 1 "
	strSQL = strSQL & " AND p.PlayerHandle ='" & replace(Request.QueryString("opponent"), "'", "''") & "'"
	ors.open strSQL, oConn
	if not (ors.eof and ors.bof) then
		dID = ors.Fields("PPLLinkID").Value 
		dEmail = ors.Fields ("PlayerEmail").Value
		dName = ors.Fields ("PlayerHandle").Value 
		if (ors.fields(0).value = "Available" or left(ors.Fields(0).Value, 8) = "Defeated") then
			cValid = true
		end if
	end if
	ors.close
	

	if cValid then
		cValid=False

		
		strSQL = "SELECT lnk.Status, lnk.PPLLinkID, p.PlayerHandle, p.PlayerEmail "
		strSQL = strSQL & " FROM tbl_players p, lnk_p_pl lnk, tbl_playerladders l "
		strSQL = strSQL & " WHERE l.PlayerLadderID = lnk.PlayerLadderID "
		strSQL = strSQL & " AND p.PlayerID = lnk.PlayerID "
		strSQL = strSQL & " AND l.PlayerLadderName = '"	& CheckString(Request.QueryString("ladder")) & "' "
		strSQL = strSQL & " AND lnk.IsActive = 1 "
		If (PlayerName <> "") Then 
			strSQL = strSQL & " AND p.PlayerHandle ='" & replace(PlayerName, "'", "''") & "'"
		Else
			strSQL = strSQL & " AND p.PlayerHandle ='" & replace(Session("uName"), "'", "''") & "'"
		End If
		ors.open strSQL, oConn
		if not (ors.eof and ors.bof) then
			aID = ors.Fields("PPLLinkID").Value 
			aEmail = ors.Fields ("PlayerEmail").Value
			aName = ors.Fields ("PlayerHandle").Value 
			if (ors.fields(0).value = "Available" or left(ors.Fields(0).Value,6) = "Immune") then
				cValid = true
			end if
		end if
		ors.close
		if cValid = true then 
			strSQL = "select PlayerLadderID from tbl_playerLadders where playerLadderName='" & CheckString(Request.QueryString("ladder")) & "'"
			ors.open strSQL, oconn
			if not (ors.eof and ors.bof) then
				mLadder = ors.fields(0).value
			end if
			ors.close

			iMap = "TBD"
			strSQL = "insert into tbl_PlayerMatches(MatchDefenderID, MatchAttackerID, MatchDate, MatchChallengeDate, MatchMap1ID, MatchLadderID) values ('" & dID & "','" & aID & "','TBD','" & now & "','" & iMap & "','" & mLadder & "')"
			ors.open strSQL, oconn
			strSQL = "update lnk_p_pl set Status='Defending', LadderMode=0, ModeFlagTime=null where PPLLinkID=" & dID
			ors.open strSQL, oconn
			strSQL = "update lnk_p_pl set Status='Attacking', LadderMode=0, ModeFlagTime=null where PPLLinkID=" & aID
			ors.open strSQL, oconn
			
			lName = Request.QueryString ("Ladder")
'			Subject     = "TWL: " & lname & ": " & aName & " attacks " & dname
'			Text = aName & ", you have attacked " & dname & " on the " & lname & " Ladder." & vbcrlf
'			text = text & "Your opponent should respond within 48 hours of this e-mail." & vbcrlf
'			Text = text & "This is a confirmation e-mail, please do not reply to this message."
'			if not(MailPlayersOnLadder(aid, aName, aEmail, text, subject, true, false, mladder)) then
'				'Response.Write "Mail not sent to attacker"
'			end if
			'Response.Write "Mail Sent"
			Subject     = "TWL: " & lname & ": " & dName & " defends against " & aname
			Text = dName & ", you have been attacked by " & aname & " on the " & lname & " Ladder." & vbcrlf
			test = text & "You have 48 hours to respond to this challenge by logging into the Teamwarfare page and offering two different dates and times." & vbcrlf			
			Text = text & "This is a confirmation e-mail, please do not reply to this message."
			if not(MailPlayersOnLadder(did, dName, dEmail, text, subject, true, false, mladder)) then
				'Response.Write "Mail not sent to defender"
			end if
			'Response.Write "Mail Sent"			
			'Response.write text
		else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=6"
		end if
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=4"
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "PlayerLadderAdmin.asp?ladder=" & server.urlencode(lName) & "&player=" & server.urlencode(playerName)
end if
'-----------------------------------------------
' Player Accept Match
'-----------------------------------------------
if Request.Form("SaveType") = "PlayerAcceptMatch" then
	If Not(IsSysAdmin()) AND NOT(IsPlayerLadderAdmin(request("Ladder"))) AND NOT(Session("uName") = request("playerName")) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		day1 = Request.Form("Day1")
		day2 = Request.Form("day2")
		mdate1 = day1
		mDate1=left(mDate1, len(mDate1)-4)
		loc=instr(1,mDate1,",")+1
		mdate1=mid(mdate1, loc+1, len(mdate1)-loc)
		pubdate1=datevalue(mDate1)
		pubtime1=timevalue(mDate1)
		
		mdate2 = day2
		mdate2=left(mdate2, len(mdate2)-4)
		loc=instr(1,mdate2,",")+1
		mdate2=mid(mdate2, loc+1, len(mdate2)-loc)
		pubdate2=datevalue(mdate2)
		pubtime2=timevalue(mdate2)
		if (pubdate1 = pubdate2) then
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "errorpage.asp?error=10"
			'error
		end if
		if (pubtime1 = pubtime2) then
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "errorpage.asp?error=12"
			'error
		end if
		strSQL = "update tbl_playerMatches set MatchAwaitingForfeit=0, forfeitreason='xxx', MatchSelDate1='" & Request.Form("Day1") & "', MatchSelDate2='" & Request.Form("day2") & "', MatchAcceptanceDate='" & now & "' where PlayerMatchID=" & Request.Form("matchid")
		oConn.Execute(strSQL)

		strsql = "SELECT l.PlayerLadderName, p.PlayerHandle, lnk.PPLLinkID, pm.PlayerMatchID, l.PlayerLadderID, p.PlayerEMail"
		strsql = strsql & " FROM tbl_players p, tbl_PlayerLadders l, lnk_p_pl lnk, tbl_playerMatches pm WHERE "
		strsql = strsql & " l.Playerladderid = lnk.Playerladderid AND p.PlayerID = lnk.PlayerID "
		strSQL = strSQL & " AND pm.MatchDefenderID = lnk.PPLLinkID "
		strsql = strsql & " AND pm.PlayerMatchID=" & request.form("matchID")
		ors.Open strsql, oconn
		if not(ors.EOF and ors.BOF) then
			lname = ors.Fields(0).Value
			dname = ors.Fields(1).Value
			did = ors.Fields(2).Value
			lid = ors.Fields(3).Value
			demail = ors.Fields("PlayerEmail").Value 
		end if
		ors.Close
		strsql = "SELECT l.PlayerLadderName, p.PlayerHandle, lnk.PPLLinkID, pm.PlayerMatchID, l.PlayerLadderID, p.PlayerEMail"
		strsql = strsql & " FROM tbl_players p, tbl_PlayerLadders l, lnk_p_pl lnk, tbl_playerMatches pm WHERE "
		strsql = strsql & " l.Playerladderid = lnk.Playerladderid AND p.PlayerID = lnk.PlayerID "
		strSQL = strSQL & " AND pm.MatchAttackerID = lnk.PPLLinkID "
		strsql = strsql & " AND pm.PlayerMatchID=" & request.form("matchID")
		ors.Open strsql, oconn
		if not(ors.EOF and ors.BOF) then
			aname = ors.Fields(1).Value
			aid = ors.Fields(2).Value
			aemail = ors.Fields("PlayerEmail").Value 
		end if
		ors.Close

'		Subject     = "TWL: " & lname & ": " & dName & " choose dates..."
'		Text = aName & ", you have attacked " & dname & " on the " & lname & " Ladder." & vbcrlf
'		Text = text & "Please login to the teamwarfare admin panel and accept a proposed date." & vbcrlf
'		Text = text & "This is a confirmation e-mail, please do not reply to this message." & vbcrlf
'		if not(MailPlayersOnLadder(aid, aname, aemail, text, subject, true, false, lid)) then
'			'Response.Write "Mail not sent to attacker"
'		end if
'		
'		Subject     = "TWL: " & lname & ": " & dName & ", you chose dates against " & aname
'		dtext = dName & ", you have been attacked by " & aname & " on the " & lname & " Ladder." & vbcrlf
'		dtext = dtext & "You have done all you need to do until the opposing player chooses a date. " & vbcrlf
'		dtext = dtext & "This is a confirmation e-mail, please do not reply to this message." & vbcflr
'		if not(MailPlayersOnLadder(did, dname, demail, dtext, subject, true, false, lid)) then
'			'Response.Write "Mail not sent to defender"
'		end if
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "PlayerLadderAdmin.asp?ladder=" & Request.form("ladder") & "&player=" & Request("PlayerName")
end if
'-----------------------------------------------
' Player Accept Match Date
'-----------------------------------------------
if Request.FORM("SaveType") = "PlayerAcceptMatchDate" then
	If Not(IsSysAdmin()) AND NOT(IsPlayerLadderAdmin(Request("LadderName"))) AND Not(session("uName") = request("PlayerName")) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		mDate=Request.Form("matchdate")
		' Here is where we select random maps
		randomize
		LadderName = Request("LadderName")
		PlayerName = Request("PlayerName")
		strSQL = "select map.mapname FROM tbl_maps map, lnk_pl_m lnk, tbl_playerLadders l "
		strSQL = strSQL & " WHERE l.Playerladderid = lnk.Playerladderid "
		strSQL = strSQL & " AND map.mapid = lnk.mapid AND l.PlayerLadderName='" & CheckString(LadderName) & "'"
		ors.open strSQL, oconn
		i=3
		if not (ors.eof and ors.bof) then
			do while not ors.eof
				i = i + 1
				ors.movenext
			loop
		end if
		ors.close
		ors.open strSQL, oconn
		dim parr(256)
		parr(0) = ""
		parr(1) = ""
		parr(2) = ""
		if not (ors.eof and ors.bof) then
			j=3
			do while not ors.eof
				parr(j) = ors.fields(0).value
				j=j+1
				ors.movenext
			loop
		end if		
		parr(j) = ""
		parr(j+1) = ""
		parr(j+2) = ""
		nummaps = j + 2
		k=0
		for p = 0 to 2
			rNum = Int((nummaps - 0 + 1) * Rnd + 0)
			map1 = parr(rNum)
			if len(trim(map1)) > 0 then
				p=2
			else
				p=1
			end if
		next
		v=0
		'Response.Write "Map1 : " & Server.HTMLEncode(map1)
		strSQL =  "update tbl_playerMatches set MatchAwaitingForfeit=0, forfeitreason='xxx', MatchMap1ID='" & replace(map1, "'", "''") & "', MatchDate='" & mDate & "', MatchLockDate='" & now & "' where PlayerMatchID=" & Request.form("matchid")
		ors.close
		OCONN.Execute STRsql
			
		

		strsql = "SELECT l.PlayerLadderName, p.PlayerHandle, lnk.PPLLinkID, pm.PlayerMatchID, l.PlayerLadderID, p.PlayerEMail"
		strsql = strsql & " FROM tbl_players p, tbl_PlayerLadders l, lnk_p_pl lnk, tbl_playerMatches pm WHERE "
		strsql = strsql & " l.Playerladderid = lnk.Playerladderid AND p.PlayerID = lnk.PlayerID "
		strSQL = strSQL & " AND pm.MatchDefenderID = lnk.PPLLinkID "
		strsql = strsql & " AND pm.PlayerMatchID=" & request.form("matchID")
		ors.Open strsql, oconn
		if not(ors.EOF and ors.BOF) then
			lname = ors.Fields(0).Value
			dname = ors.Fields(1).Value
			did = ors.Fields(2).Value
			lid = ors.Fields(3).Value
			demail = ors.Fields("PlayerEmail").Value 
		end if
		ors.Close
		strsql = "SELECT l.PlayerLadderName, p.PlayerHandle, lnk.PPLLinkID, pm.PlayerMatchID, l.PlayerLadderID, p.PlayerEMail"
		strsql = strsql & " FROM tbl_players p, tbl_PlayerLadders l, lnk_p_pl lnk, tbl_playerMatches pm WHERE "
		strsql = strsql & " l.Playerladderid = lnk.Playerladderid AND p.PlayerID = lnk.PlayerID "
		strSQL = strSQL & " AND pm.MatchAttackerID = lnk.PPLLinkID "
		strsql = strsql & " AND pm.PlayerMatchID=" & request.form("matchID")
		ors.Open strsql, oconn
		if not(ors.EOF and ors.BOF) then
			aname = ors.Fields(1).Value
			aid = ors.Fields(2).Value
			aemail = ors.Fields("PlayerEmail").Value 
		end if
		ors.Close

'		Subject     = "TWL: " & lname & ": " & dName & " match scheduled ."
'		Text = aName & ", you have attacked " & dname & " on the " & lname & " Ladder." & vbcrlf
'		Text = text & "Please login to the teamwarfare admin panel to confirm side choices." & vbcrlf
'		Text = text & "This is a confirmation e-mail, please do not reply to this message." & vbcrlf
'		if not(MailPlayersOnLadder(aid, aname, aemail, text, subject, true, false, lid)) then
'			'Response.Write "Mail not sent to attacker"
'		end if
'		'Response.Write "Mail Sent"
'		Subject     = "TWL: " & lname & ": " & aName & " scheduled match."
'		Text = dName & ", you have been attacked by " & aname & " on the " & lname & " Ladder." & vbcrlf
'		Text = text & "Please login to the teamwarfare admin panel to confirm side choices." & vbcrlf
'		Text = text & "This is a confirmation e-mail, please do not reply to this message." & vbcrlf
'		if not(MailPlayersOnLadder(did, dname, demail, text, subject, true, false, lid)) then
'			'Response.Write "Mail not sent to defender"
'		end if
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "PlayerLadderAdmin.asp?ladder=" & server.urlencode(LadderName) & "&player=" & server.urlencode(PlayerName)
		
end if
'-----------------------------------------------
' Player Match Reporting
'-----------------------------------------------
if Request.QueryString("SaveType") = "PlayerReportMatch" then
	matchid=Request.QueryString("matchid")
	matchwinnername=Request.QueryString("matchwinner")
	matchlosername=Request.QueryString("matchloser")
	matchwinnerid=Request.QueryString("matchwinnerid")
	matchloserid=Request.QueryString("matchloserid")
	map1=Request.QueryString("map1")
	map1defenderscore=Request.QueryString("map1defenderscore")
	map1attackerscore=Request.QueryString("map1attackerscore")
	map1forfeit=Request.QueryString("map1forfeit")
	matchdate=Request.QueryString("matchdate")
	matchwinnerdefending=Request.QueryString("matchwinnerdefending")
	ladderid=Request.QueryString("ladderid")

	if not(IsSysAdmin()) AND Not(IsPlayerLadderAdminByID(ladderID)) AND Not(Session("uName") = matchlosername) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQL = "select lnk.PPLLinkID, lnk.rank from lnk_p_pl lnk, tbl_players p WHERE p.PlayerHandle = '" & Replace(matchwinnername, "'", "''") & "' AND p.PlayerID = lnk.PlayerID AND lnk.PlayerLadderID=" & ladderID
		ors.open strSQL, oconn
'		Response.Write strsql & "<br><br>"
		if not (ors.eof and ors.bof) then
			winnerid=ors.fields(0).value
			winnerrank=ors.fields(1).value
		end if
		ors.close
		strSQL = "select lnk.PPLLinkID, lnk.rank from lnk_p_pl lnk, tbl_players p WHERE p.PlayerHandle = '" & Replace(matchlosername, "'", "''") & "' AND p.PlayerID = lnk.PlayerID AND lnk.PlayerLadderID=" & ladderID
		ors.open strSQL, oconn
'		Response.Write strsql & "<br><br>"
		if not (ors.eof and ors.bof) then
			loserid=ors.fields(0).value
			loserrank=ors.fields(1).value
		end if
		ors.close
		strSQL = "insert into tbl_playerHistory (PlayerMatchID, MatchWinnerID, MatchLoserID, MatchMap1, "
		strSQL = strSQL & "MatchMap1DefenderScore,  "
		strSQL = strSQL & "MatchMap1AttackerScore, "
		strSQL = strSQL & "MatchMap1forfeit, "
		strSQL = strSQL & "MatchDate, MatchWinnerDefending, MatchLadderID,MatchAttackerRank,MatchDefenderRank) values "
		strSQL = strSQL & "(" & matchid & ", " & winnerid & "," & loserid & ",'" & replace(map1, "'", "''")
		strSQL = strSQL & "','" & map1defenderscore
		strSQL = strSQL & "','" & map1attackerscore
		strSQL = strSQL & "','" & map1forfeit & "','" & matchdate & "','" & matchwinnerdefending & "','" & ladderID & "',"
		
		'// added by fission 9/30/04
		If matchwinnerdefending = "True" Then
			strSQL = strSQL & "'" & loserrank & "','" & winnerrank & "')"
		Else
			strSQL = strSQL & "'" & winnerrank & "','" & loserrank & "')"
		End If
		
'		Response.Write strSQL & "<br><br>"
		ors.open strSQL, oconn
'			Response.Write strsql & "<br><br>"
		strSQL = "delete from tbl_playerMatches where PlayerMatchID=" & matchid
		ors.open strSQL, oconn
'				Response.Write strsql & "<br><br>"
		strSQL = "update lnk_p_pL set status='Defeated by " & replace(matchwinnername, "'", "''") & "', Losses=(Losses+1), LadderMode='1', ModeFlagTime='" & now & "' where PPLLinkID=" & loserid
		ors.open strSQL, oconn
'				Response.Write strsql & "<br><br>"
		pstat="Immune until " & formatdatetime(dateadd("h",1,now),3)
'			Response.Write pstat
		strSQL = "update lnk_p_pL set Status='" & pstat & "', wins=(wins+1), LadderMode=3, ModeFlagtime='" & now & "' where PPLLinkID=" & winnerid
		ors.open strSQL, oconn
	
'		Response.Write strsql & "<br><br>"
		if matchwinnerdefending = "True" then
			'defender wins
		else
			strSQL = "update lnk_p_pL set rank=(rank + 1) where (rank >" & (loserrank - 1) & " and rank < " & winnerrank & ") and isactive=1 and PlayerladderID = " & ladderID
'			Response.Write strsql & "<br><br>"
			ors.open strSQL, oconn
			strSQL = "update lnk_p_pL set rank=" & loserrank & " where PPLLinkID=" & winnerID & " and PlayerladderID=" & ladderID
			ors.open strSQL, oconn
'			Response.Write strsql
		end if
		strSQL="update tbl_playerComms set CommDead=1 where Playermatchid=" & matchid
		ors.open strSQL, oconn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "viewplayer.asp?player=" & server.URLEncode(matchlosername)
end if
'-----------------------------------------------
' Delete Player
'-----------------------------------------------
if Request.Form("SaveType") = "DeletePlayer" then
	if not (IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	Dim sUpper
	sUpper = UCase(Request.Form("player"))
	If ( sUpper = "TRISTON" Or sUpper = "POLARIS" Or sUpper = "TOTALCARNAGE" ) Then
		Response.Clear
		Response.Redirect "adminmenu.asp"
	End If
		
	strSQL = "Select playerID from tbl_players where playerhandle='" & replace(Request.Form("player"), "'","''") & "'"
	ors.open strSQL, oconn
	if not (ors.eof and ors.bof) then
		playerid=ors.fields(0).value
	end if
	ors.close
	if playerid <> "" then
		strSQL="delete from lnk_T_P_L where playerID=" & playerid
		oConn.Execute(strSQL)
		strSQL="delete from tbl_players where playerid=" & playerid
		oConn.Execute(strSQL)
		strSQL="delete from lnk_player_identifier where playerid=" & playerid
		oConn.Execute(strSQL)
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminmenu.asp"
end if
'-----------------------------------------------
' Delete Team
'-----------------------------------------------
if Request.Form("SaveType") = "DeleteTeam" then
	if not (IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if

' add delete all link id stuff
	teamid=request.form("teamid")
	strsql = "select teamname from tbl_teams where teamid='" & teamid & "'"
	ors.open strsql, oconn
	teamname = ors.fields(0).value
	ors.close
	cleartodelete=true
	strsql="select * from lnk_T_L where TeamID = " & teamID
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		' check all ladders
		do while not ors.eof
			if ors.fields("Status").value <> "Available" then
				cleartodelete=false
			end if
			ors.movenext			
		loop
	end if
	'cleartodelete=true  I had to change this to true for a minute to allow the deletion of a team.  
	'It looks like it's possible for the status to not update properly when a team leaves a competition.
	if cleartodelete then
		ors.requery

		if not (ors.eof and ors.bof) then
			do while not ors.eof
				linkID = ors.fields(0).value
				strSQL="delete from lnk_T_P_L where TLlinkid=" & linkid
				ors2.open strSQL, oconn
				strSQL="select * from lnk_t_l where teamid=" & teamid & " AND ladderid=" & ors.fields("LadderID").value & " AND isactive=1"
				ors2.open strsql, oconn
				if not (ors2.bof and ors2.eof) then
					ladderid=ors2.fields("LadderId").value
					oldrank=ors2.fields("Rank").value				
					strSQL="update lnk_T_L set isactive=0, rank=null where TeamID=" & teamid & " and ladderid=" & ladderid
					oConn.Execute(strSQL)

					strsql="update lnk_T_L set rank=(rank - 1) where (rank > " & oldrank & ") and ladderID = " & ladderID
					oConn.Execute(strSQL)
				end if
				ors2.close
				'response.write "found one <br>"
				ors.movenext
			loop
		end if
		ors.close
'		strSQL="delete from lnk_T_L where TeamID=" & teamid
'		ors3.open strsql, oconn		
		strSQL="update tbl_teams set teamactive=0, teamname='" & replace(teamname, "'","''") & " code_deleted' where teamid=" & teamid
		ors.open strSQL, oconn
	else
		ors.close
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "adminops.asp?error=1"
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminmenu.asp"
end if
'-----------------------------------------------
' Restore Team
'-----------------------------------------------
if Request.Form("SaveType") = "RestoreTeam" then
	if not (IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strsql = "select teamname from tbl_teams where teamid='" & request.form("teamid") & "'"
	ors.open strsql, oconn
	teamname = ors.fields(0).value
	ors.close
	teamname=replace(teamname, " code_deleted", "")
	strSQL="update tbl_teams set teamactive=1, teamname='" & replace(teamname, "'","''") & "' where teamid=" & request.form("teamid")
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminmenu.asp"
end if

'-----------------------------------------------
' Admin Kill Player Match
'-----------------------------------------------
if Request.Form("SaveType") = "PlayeradmKillMatch" then
	rAction=Request.Form("rAction")
	ladderid = Request.Form("LadderID")
	ladderName = Request.Form("LadderName")
	rAdmin = Request.Form("rAdmin")
	If Not(bSysAdmin OR IsPlayerLadderAdminByID(LadderID)) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	End If
	For Each FormItem In Request.Form
		If Left(Formitem, 11) = "AdminMatch_" Then
			MatchID = Right(FormItem, Len(FormItem) - 11)
			Action = Request.Form(FormItem)
			Select Case Action
				Case "OVERRIDE"
					strSQL="UPDATE tbl_PlayerMatches set ForfeitReason='Admin Override by " & CheckString(session("uName")) & " on " & now & "' where Playermatchid='" & matchid & "'"
					oConn.Execute(strSQL)
				Case "KILL"
					strSQL = "EXEC usp_PlayerMatchKill '" & MatchID & "'"
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						DefenderLinkID	= oRS.Fields("DefenderLinkID").Value
						DefenderName	= oRS.Fields("DefenderName").Value
						DefenderEmail	= oRs.Fields("DefenderEmail").Value 
						AttackerLinkID	= oRS.Fields("AttackerLinkID").Value	
						AttackerName	= oRS.Fields("AttackerName").Value
						AttackerEmail	= oRs.Fields("AttackerEmail").Value 
						LadderName		= oRS.Fields("LadderName").Value
					End If
					oRS.NextRecordset 
					If DefenderLinkID <> 0 Then
						subjectd = "TWL: " & LadderName & ": Match Killed"
						textd = DefenderName & ", your current match has been killed. If this is in error, please e-mail your ladder admin."
						subjecta = "TWL: " & LadderName & ": Match Killed"
						texta = AttackerName & ", your current match has been killed. If this is in error, please e-mail your ladder admin."

						if not(MailPlayersOnLadder(AttackerLinkID, AttackerName, AttackerEmail, texta, subjecta, true, false, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
						if not(MailPlayersOnLadder(DefenderLinkID, DefenderName, DefenderEmail, textd, subjectd, true, false, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
					End If
				Case "FORFEITD"
					strSQL = "EXEC usp_PlayerMatchForfeit_Defender '" & MatchID & "'"
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						DefenderLinkID	= oRS.Fields("DefenderLinkID").Value
						DefenderName	= oRS.Fields("DefenderName").Value
						DefenderEmail	= oRs.Fields("DefenderEmail").Value 
						AttackerLinkID	= oRS.Fields("AttackerLinkID").Value	
						AttackerName	= oRS.Fields("AttackerName").Value
						AttackerEmail	= oRs.Fields("AttackerEmail").Value 
						LadderName		= oRS.Fields("LadderName").Value
					End If
					oRS.NextRecordset 
					If DefenderLinkID <> 0 Then
						subjectd = "TWL: " & LadderName & ": Forfeit loss awarded"
						textd = DefenderName & ", you have been penalized a forfeit by an administrator. This can be due to a lack of response, or by request. If this is in error, please e-mail your admin."
						subjecta="TWL: " & LadderName & ": Forfeit win awarded"
						texta=AttackerName & ", you have been awarded a forfeit win in your current match, please contact an administrator if this is incorrect."

						if not(MailPlayersOnLadder(AttackerLinkID, AttackerName, AttackerEmail, texta, subjecta, true, false, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
						if not(MailPlayersOnLadder(DefenderLinkID, DefenderName, DefenderEmail, textd, subjectd, true, false, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
					End If
				Case "FORFEITA"
					strSQL = "EXEC usp_PlayerMatchForfeit_Attacker '" & MatchID & "'"
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						DefenderLinkID	= oRS.Fields("DefenderLinkID").Value
						DefenderName	= oRS.Fields("DefenderName").Value
						DefenderEmail	= oRs.Fields("DefenderEmail").Value 
						AttackerLinkID	= oRS.Fields("AttackerLinkID").Value	
						AttackerName	= oRS.Fields("AttackerName").Value
						AttackerEmail	= oRs.Fields("AttackerEmail").Value 
						LadderName		= oRS.Fields("LadderName").Value
					End If
					oRS.NextRecordset 
					If DefenderLinkID <> 0 Then
						subjecta = "TWL: " & LadderName & ": Forfeit loss awarded"
						texta = AttackerName & ", you have been penalized a forfeit by an administrator. This can be due to a lack of response, or by request. If this is in error, please e-mail your admin."
						subjectd="TWL: " & LadderName & ": Forfeit win awarded"
						textd=DefenderName & ", you have been awarded a forfeit win in your current match, please contact an administrator if this is incorrect."

						if not(MailPlayersOnLadder(AttackerLinkID, AttackerName, AttackerEmail, texta, subjecta, true, false, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
						if not(MailPlayersOnLadder(DefenderLinkID, DefenderName, DefenderEmail, textd, subjectd, true, false, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
					End If
				Case "NOTHING"
					' Do nothing, since we didnt want to do anything
			End Select
		End If
	Next
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminops.asp?rAdmin=" & rAdmin & "&LadderID=" & LadderID & "&LadderName=" & Server.URLEncode(LadderName)
end if
'--------------------------
' Admin Kill Match
'--------------------------
if Request.Form("SaveType") = "admKillMatch" then
	rAction=Request.Form("rAction")
	ladderid = Request.Form("LadderID")
	ladderName = Request.Form("LadderName")
	rAdmin = Request.Form("rAdmin")
	If Not(bSysAdmin OR IsLadderAdminbyID(LadderID)) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	End If
	For Each FormItem In Request.Form
		If Left(Formitem, 11) = "AdminMatch_" Then
			MatchID = Right(FormItem, Len(FormItem) - 11)
			Action = Request.Form(FormItem)
			Select Case Action
				Case "OVERRIDE"
					strSQL="UPDATE tbl_matches set ForfeitReason='Admin Override by " & CheckString(session("uName")) & " on " & now & "' where matchid='" & matchid & "'"
					oConn.Execute(strSQL)
				Case "KILL"
					strSQL = "EXEC usp_MatchKill '" & MatchID & "', '" & Session("PlayerID") & "'"
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						DefenderLinkID	= oRS.Fields("DefenderLinkID").Value
						DefenderName	= oRS.Fields("DefenderName").Value
						AttackerLinkID	= oRS.Fields("AttackerLinkID").Value	
						AttackerName	= oRS.Fields("AttackerName").Value
						LadderName		= oRS.Fields("LadderName").Value
					End If
					oRS.NextRecordset 
					If DefenderLinkID <> 0 Then
						subjectd = "TWL: " & LadderName & ": Match Killed"
						textd = DefenderName & ", your current match has been killed. If this is in error, please e-mail your ladder admin."
						subjecta = "TWL: " & LadderName & ": Match Killed"
						texta = AttackerName & ", your current match has been killed. If this is in error, please e-mail your ladder admin."

						if not(MailTeamCaptains(AttackerLinkID, texta, subjecta, true, False, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
						if not(MailTeamCaptains(DefenderLinkID, textd, subjectd, true, False, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
					End If
				Case "FORFEITD"
					strSQL = "EXEC usp_MatchForfeit_Defender '" & MatchID & "', '" & Session("PlayerID") & "'"
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						DefenderLinkID	= oRS.Fields("DefenderLinkID").Value
						DefenderName	= oRS.Fields("DefenderName").Value
						AttackerLinkID	= oRS.Fields("AttackerLinkID").Value	
						AttackerName	= oRS.Fields("AttackerName").Value
						LadderName		= oRS.Fields("LadderName").Value
					End If
					oRS.NextRecordset 
					If DefenderLinkID <> 0 Then
						subjectd = "TWL: " & LadderName & ": Forfeit loss awarded"
						textd = DefenderName & ", your team has been penalized a forfeit by an administrator. This can be due to a lack of response, or by request. If this is in error, please e-mail your admin."
						subjecta="TWL: " & LadderName & ": Forfeit win awarded"
						texta=AttackerName & ", your team has been awarded a forfeit win in your current match, please contact an administrator if this is incorrect."

						if not(MailTeamCaptains(AttackerLinkID, texta, subjecta, true, False, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
						if not(MailTeamCaptains(DefenderLinkID, textd, subjectd, true, False, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
					End If
				Case "FORFEITA"
					strSQL = "EXEC usp_MatchForfeit_Attacker '" & MatchID & "', '" & Session("PlayerID") & "'"
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						DefenderLinkID	= oRS.Fields("DefenderLinkID").Value
						DefenderName	= oRS.Fields("DefenderName").Value
						AttackerLinkID	= oRS.Fields("AttackerLinkID").Value	
						AttackerName	= oRS.Fields("AttackerName").Value
						LadderName		= oRS.Fields("LadderName").Value
					End If
					oRS.NextRecordset 
					If DefenderLinkID <> 0 Then
						subjecta = "TWL: " & LadderName & ": Forfeit loss awarded"
						texta = AttackerName & ", your team has been penalized a forfeit by an administrator. This can be due to a lack of response, or by request. If this is in error, please e-mail your admin."
						subjectd="TWL: " & LadderName & ": Forfeit win awarded"
						textd=DefenderName & ", your team has been awarded a forfeit win in your current match, please contact an administrator if this is incorrect."

						if not(MailTeamCaptains(AttackerLinkID, texta, subjecta, true, False, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
						if not(MailTeamCaptains(DefenderLinkID, textd, subjectd, true, False, ladderid)) then
							'Response.Write "Mail not sent to attacker"
						end if
					End If
				Case "NOTHING"
					' Do nothing, since we didnt want to do anything
			End Select
		End If
	Next
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminops.asp?rAdmin=" & rAdmin & "&LadderID=" & LadderID & "&LadderName=" & Server.URLEncode(LadderName)
end if
'-----------------------------------------------
' Halt Ladder
'-----------------------------------------------
if Request.QueryString("SaveType") = "HaltLadder" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		if Request.QueryString("ladder") <> "" then
			strSQL="update tbl_ladders set LadderActive=0 where LadderID=" & Request.QueryString("ladder")
			ors.open strSQL, oconn
		end if
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminops.asp?rAdmin=Ladder"
end if
'-----------------------------------------------
' Start Ladder
'-----------------------------------------------
if Request.QueryString("SaveType") = "StartLadder" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		if Request.QueryString("ladder") <> "" then
			strSQL="update tbl_ladders set LadderActive=1 where LadderID=" & Request.QueryString("ladder")
			ors.open strSQL, oconn
		end if
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminops.asp?rAdmin=Ladder"
end if
'-----------------------------------------------
' PHalt Ladder
'-----------------------------------------------
if Request.QueryString("SaveType") = "PHaltLadder" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		if Request.QueryString("ladder") <> "" then
			strSQL="update tbl_playerladders set Active=0 where PlayerLadderID=" & Request.QueryString("ladder")
			ors.open strSQL, oconn
		end if
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminops.asp?rAdmin=PLadder"
end if
'-----------------------------------------------
' PStart Ladder
'-----------------------------------------------
if Request.QueryString("SaveType") = "PStartLadder" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		if Request.QueryString("ladder") <> "" then
			strSQL="update tbl_playerladders set Active=1 where PlayerLadderID=" & Request.QueryString("ladder")
			ors.open strSQL, oconn
		end if
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminops.asp?rAdmin=PLadder"
end if
'-----------------------------------------------
' New Map
'-----------------------------------------------
if Request.Form("SaveType") = "NewMap" then
	if not(IsSysAdmin()) and not(IsAnyLadderAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQL="insert into tbl_maps (MapName, MapAbbreviation, MapType, MapTerrain, Generators, Inventories, BaseTurrets, VehiclePad, Description, MapImage) "
		strSQL=strSQL & "values ('" & replace(Request.form("mapname"),"'", "''") & "','" & Request.Form("mapabbr") &  "','" & Request.Form("maptype") & "','" 
		strSQL=strSQL & Request.Form("mapterr") & "','" & Request.Form("mapgens") & "','" & Request.Form("mapinv") & "','"  
		strSQL=strSQL & Request.Form("mapturr") & "','" & Request.Form("mapv") & "','" & replace(Request.Form("mapdesc"), "'", "''") & "','" & Request.Form("mapimage") & "')"
		'Response.write strSQL
		ors.open strSQL, oconn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminops.asp?rAdmin=Ladder"
end if
'-----------------------------------------------
' Kill Map
'-----------------------------------------------
if Request.Form("SaveType") = "KillMap" then
	if not (IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL="delete from tbl_maps where mapname='" & replace(Request.Form("mapname"), "'", "''") & "'"
	'Response.write strSQL
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "adminops.asp?rAdmin=Ladder"
end if
'-----------------------------------------------
' Edit Map
'-----------------------------------------------
if Request.Form("SaveType") = "SaveMap" then
	if not(IsSysAdmin()) and not(IsAnyLadderAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQL="update tbl_maps set mapname='" & replace(Request.Form("mapname"), "'", "''") & "', "
		strSQL=strSQL & " mapabbreviation='" & Request.form("mapabbr") & "', "
		strSQL=strSQL & " maptype='" & Request.form("maptype") & "', "
		strSQL=strSQL & " mapterrain='" & Request.form("mapterr") & "', "
		strSQL=strSQL & " Generators='" & Request.form("mapgens") & "', "
		strSQL=strSQL & " Inventories='" & Request.form("mapinv") & "', "
		strSQL=strSQL & " BaseTurrets='" & Request.form("mapturr") & "', "
		strSQL=strSQL & " VehiclePad='" & Request.form("mapV") & "', "
		strSQL=strSQL & " Description='" & replace(Request.form("mapdesc"), "'", "''") & "', "
		strSQL=strSQL & " MapImage='" & Request.form("mapimage") & "' where mapid=" & Request.Form("mapid")
		'Response.write strSQL
		ors.open strSQL, oconn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "mapguide.asp"
end if
'-----------------------------------------------
' Promote to Captain
'-----------------------------------------------
if Request.Form("SaveType") = "PromoteCaptain" then
	if not(IsSysAdmin()) and not(IsTeamCaptain(request.form("team"), request.form("ladder"))) and not(IsTeamFounder(request.form("team"))) and not(IsLadderAdmin(request.form("ladder"))) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	
	strSQL="UPDATE lnk_T_P_L set isadmin=1 where tpllinkid=" & Request.Form("playerlist")
	'Response.write strSQL
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "TeamLadderAdmin.asp?ladder=" & server.urlencode(Request.form("ladder")) & "&team=" & server.urlencode(Request.Form("team"))
end if
'-----------------------------------------------
' Demote Captain
'-----------------------------------------------
if Request.Form("SaveType") = "DemoteCaptain" then
	if not(IsSysAdmin()) and not(IsTeamCaptain(request.form("team"), request.form("ladder"))) and not(IsTeamFounder(request.form("team"))) and not(IsLadderAdmin(request.form("ladder"))) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL="select lnk_T_L.TLLinkID, playerid from lnk_T_L inner join lnk_T_P_L on lnk_T_P_L.TLLinkID=lnk_T_L.TLLinkID where lnk_T_P_L.TPLLinkID='" & Request.Form("playerlist") & "'"
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		plid=ors.Fields(0).Value
		pid=ors.Fields(1).Value 
	end if
	ors.Close
	strSQL="select teamfounderid from tbl_teams inner join lnk_T_L on lnk_T_L.teamid=tbl_teams.teamid where lnk_T_L.TLLinkID='" & plid & "'"
	ors.open strsql, oconn
	if not (ors.EOF and ors.BOF) then
		tfid=ors.Fields(0).Value
	end if
	ors.Close 
	if tfid <> pid then
		strSQL="UPDATE lnk_T_P_L set isadmin=0 where tpllinkid='" & Request.Form("playerlist") & "'"
		ors.open strSQL, oconn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "TeamLadderAdmin.asp?ladder=" & server.urlencode(Request.form("ladder")) & "&team=" & server.urlencode(Request.Form("team"))
end if
'-----------------------------------------------
' Make Sysadmin
'-----------------------------------------------
if Request.form("SaveType") = "AddSysadmin" then
	if IsSysAdmin() then
		playerid=request.form("SysadminAdd")
		strsql = "insert into sysadmins (AdminID, SendEmail) values (" & playerid & ",1)"
		ors.open strsql, oconn
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	response.redirect "adminops.asp"
end if
'-----------------------------------------------
' Delete Sysadmin
'-----------------------------------------------
if Request.form("SaveType") = "DeleteSysadmin" then
	if IsSysAdmin() then
		playerid=request.form("SysadminDel")
		strsql = "delete from sysadmins where adminid=" & playerid
		ors.open strsql, oconn
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	response.redirect "adminops.asp"
end if
'-----------------------------------------------
' Team Drop From Ladder
'-----------------------------------------------
if request.form("SaveType") = "QuitLadder" then
	tname=request.form("teamname")
	lname=request.form("laddername")
	if not(IsSysAdmin()) and not(IsLadderAdmin(lname)) and not (IsTeamFounder(tname)) and not(IsTeamCaptain(tname, lname)) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		' then check for active challenge
		teamid=request.form("TeamID")	
		strsql = "select ladderID from tbl_ladders where laddername='" & CheckString(lname) & "'"
		ors.open strsql, oconn
		if not (ors.eof and ors.bof) then
			ladderid=ors.fields("LadderID").value
		end if
		ors.close
	
		strsql = "Select TLLinkID from lnk_T_L where TeamID=" & teamid & " and ladderID=" & ladderid
		ors.open strSQL, oconn
		if not (ors.eof and ors.bof) then
			tID = ors.fields(0).value
		end if
		ors.close
		strSQL = "select * from lnk_T_L where TLLinkID=" & tID
		ors.open strSQL, oconn
		if not (ors.bof and ors.eof) then
			status=ors.fields(4).value
		end if
		'response.write status
		if status = "Available" then 
			'response.write "Clear"
		else
			'response.write "Team getting attacked"
		end if
		ors.close
		if status= "Available" or left(status,6)="Immune" or left(status,6)="Defeat" then
			'Who dropped the team
			strSQL = "INSERT INTO tbl_transaction (TransactionDetails, TransactionTime) VALUES ('Ladder dropped - " & CheckString(Session("uName")) & " - " & Request.ServerVariables("REMOTE_ADDR") & " - dropped team " & CheckString(tname) & " from " & CheckString(lname) & "', GetDate())"
	        oConn.Execute (strSQL)
			'---------------
			' Delete Roster
			'---------------
			strsql="delete from lnk_T_p_L where TLLinkID=" & tid
			ors.open strsql, oconn

			'---------------
			' Rerank Teams
			'---------------
			strSQL="select * from lnk_t_l where teamid=" & teamid & " AND ladderid=" & ladderid
			ors.open strsql, oconn
			
			if not (ors.bof and ors.eof) and ors.fields("ranK") <> 0 then
				oldrank=ors.fields("Rank").value				
				strsql="update lnk_T_L set rank=(rank - 1) where (rank > " & oldrank & ") and ladderID = " & ladderID
				ors2.open strsql, oconn
			end if
			ors.close
			'---------------
			' Update Ladder Link Table
			'---------------
			
			strSQL="update lnk_T_L set isactive=0, rank=null where TeamID=" & teamid & " and ladderid=" & ladderid 
			ors.open strsql, oconn
			strsql="select Teamname from tbl_teams where TeamID=" & teamid
			ors.open strsql, oconn
			teamname = ors.fields(0).value
			ors.close		
			Response.clear
			%>
			<script language="javascript">
				window.opener.location.href='viewteam.asp?team=<%=Server.urlencode(teamname)%>';
				window.close();
			</script>
		<%
			Response.End 
		else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "quitladder.asp?error=1&ladder=" & server.urlencode(lname) & "&team=" & server.urlencode(tname)
		end if
	end if
	
set ors = nothing
set oConn = nothing	
set ors2 = nothing	
%>
			<script language="javascript">
				window.opener.location.href='viewteam.asp?team=<%=Server.urlencode(teamname)%>';
				window.close();
			</script>
<%	
	Response.End 
end if
'-----------------------------------------------
' Player Drop From Plyaer Ladder
'-----------------------------------------------
if request.form("SaveType") = "PlayerQuitLadder" then
	pName=CheckString(request.form("playername"))
	lname=replace(request.form("laddername"), "'", "''")
	playerID = Request.Form("playerid")
	if not(IsSysAdmin()) and not(IsPlayerLadderAdmin(lname)) and not(cstr(Session("PlayerID")) = cstr(PlayerID)) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		' then check for active challenge
		playerID=request.form("playerID")	
		strsql = "select PlayerladderID from tbl_Playerladders where Playerladdername='" & lname & "'"
		ors.open strsql, oconn
		if not (ors.eof and ors.bof) then
			ladderid=ors.fields("PlayerladderID").value
		end if
		ors.close
	
		strsql = "Select PPLLinkID, Status, rank from lnk_p_pL where PlayerID=" & playerID & " and PlayerladderID=" & ladderid
		ors.open strSQL, oconn
		if not (ors.eof and ors.bof) then
			tID = ors.fields(0).value
			status = ors.Fields (1).Value
			oldrank = ors.Fields("rank").Value 
		end if
		ors.close
		'response.write status
		if status = "Available" then 
			'response.write "Clear"
		else
			'response.write "Team getting attacked"
		end if
		if status= "Available" or left(status,6)="Immune" or left(status,6)="Defeat" then
			'---------------
			' Rerank Teams
			'---------------
			if rank <> 0 Then
				strsql="update lnk_T_L set rank=(rank - 1) where (rank > " & oldrank & ") and ladderID = " & ladderID
				ors2.open strsql, oconn
			end if

			'---------------
			' Update Ladder Link Table
			'---------------
			strSQL="update lnk_p_pL set isactive=0, rank=null where PPLLinkID=" & tID
			oConn.Execute(strSQL)
			%>
			<script language="javascript">
				window.opener.location.href='<%=Request("FromURL")%>';
				this.window.close();
			</script>
			<%
			oConn.Close
			Set oConn = Nothing
			Set oRS = Nothing
			Response.End
		else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect Request("fromurl")
		end if
	end if
set ors = nothing
set oConn = nothing	
set ors2 = nothing	
%>
			<script language="javascript">
				window.opener.location.href='<%=Request("FromURL")%>';
				this.window.close();
			</script>
<%	
end if
'-----------------------------------------------
' Kick Player from Roster
'-----------------------------------------------
if request.form("savetype")="DropPlayer" then
	' need ladder id, playerid, teamid, and preferably TLLinkID
	' Check for captain status and remove it. Make sure owner is not deletable.
	' check for admin/captain/owner.
	
	tllinkid=request.form("link")
	playerid=request.form("playerid")
	strsql="select ladderid, teamid from lnk_t_l where TLLinkID=" & TLLinkID
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		tid = ors.fields(1).value
		lid = ors.fields(0).value
	end if
	ors.close
	if not(IsSysAdmin()) and not(IsTeamFounderByID(tid)) and not(IsLadderAdminByID(lid)) and not(IsTeamCaptainByID(tid, lid)) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		if tid <> "" and lid <> "" then
			strsql="select teamname from tbl_teams where teamid=" & tid
			ors.open strsql, oconn
			tname=ors.fields(0).value
			ors.close
			strsql="select laddername from tbl_ladders where ladderid=" & lid
			ors.open strsql, oconn
			lname=ors.fields(0).value
			ors.close
		end if
		if playerid <> "" then
			strsql="delete from lnk_t_p_l where tllinkid=" & tllinkid & " and playerid=" & playerid
			ors.open strsql, oconn
		end if
		'response.write Server.HTMLEncode(tname) & Server.HTMLEncode(lname)
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "teamladderadmin.asp?team=" & server.urlencode(tname) & "&ladder=" & server.urlencode(lname)
	end if
end if
'-----------------------------------------------
' Set Ladder Admin
'-----------------------------------------------
if Request.Form("SaveType") = "SetLadderAdmin" then
	if IsSysAdmin() then
		strSQl="insert into lnk_L_A (LadderID, PlayerID) values ('" & Request.Form("ladder") & "','" & Request.Form("player") & "')"
		ors.Open strsql, oconn
		strSQL = "SELECT * FROM tbl_players WHERE PlayerID = " & Request.Form("player")
		ors.Open strSQL, oconn
		tmpName = ors.Fields("PlayerHandle")
    oRs.Close
		strSQL = "SELECT * FROM tbl_ladders WHERE LadderID = " & Request.Form("ladder")
		ors.Open strSQL, oconn
		tmpLadder = ors.Fields("LadderName")
		'oRs.Close
		strSQL = "INSERT INTO tbl_transaction (TransactionDetails, TransactionTime) VALUES ('Added Ladder Admin - " & CheckString(Session("uName")) & " - " & Request.ServerVariables("REMOTE_ADDR") & " added " & CheckString(tmpName) & " to " & CheckString(tmpLadder) & "', GetDate())"
		oConn.Execute (strSQL)
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "adminmenu.asp"
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
end if
'-----------------------------------------------
' Remove Ladder Admin
'-----------------------------------------------
if Request.Form("SaveType") = "UnSetLadderAdmin" then
	if IsSysAdmin() then
		id=request.form("lnkID")
		strsql="select * from lnk_L_A where LALinkID='" & id & "'"
		ors.open strsql, oconn
		if not (ors.bof and ors.eof) then
			'response.write "Found it"
			strSQL= "delete from lnk_L_A where LALinkID=" & id
			ors.close
			ors.open strsql, oconn
		end if
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "adminmenu.asp"
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
end if
'-----------------------------------------------
' Set Ladder Admin
'-----------------------------------------------
if Request.Form("SaveType") = "PlayerSetLadderAdmin" then
	if IsSysAdmin() then
		strSQl="insert into lnk_pL_A (PlayerLadderID, PlayerID) values ('" & Request.Form("ladder") & "','" & Request.Form("player") & "')"
		ors.Open strsql, oconn
		strSQL = "SELECT * FROM tbl_players WHERE PlayerID = " & Request.Form("player")
		ors.Open strSQL, oconn
		tmpName = ors.Fields("PlayerHandle")
    oRs.Close
		strSQL = "SELECT * FROM tbl_playerladders WHERE PlayerLadderID = " & Request.Form("ladder")
		ors.Open strSQL, oconn
		tmpLadder = ors.Fields("PlayerLadderName")
		'oRs.Close
		strSQL = "INSERT INTO tbl_transaction (TransactionDetails, TransactionTime) VALUES ('Added Ladder Admin - " & CheckString(Session("uName")) & " - " & Request.ServerVariables("REMOTE_ADDR") & " added " & CheckString(tmpName) & " to " & CheckString(tmpLadder) & "', GetDate())"
		oConn.Execute (strSQL)
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "adminmenu.asp"
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
end if
'-----------------------------------------------
' Remove Ladder Admin
'-----------------------------------------------
if Request.Form("SaveType") = "PlayerUnSetLadderAdmin" then
	if IsSysAdmin() then
		id=request.form("lnkID")
		strSQL= "delete from lnk_pL_A where PLAdminID=" & id
		ors.open strsql, oconn
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "adminmenu.asp"
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
end if
'-----------------------------------------------
' Put Team on Rest
'-----------------------------------------------
if Request.QueryString("SaveType") = "Rest" then
	if not(IsSysAdmin()) and not(IsLadderAdmin(request.querystring("ladder"))) and not(IsTeamFounder(request.querystring("team"))) and not(IsTeamCaptain(request.querystring("team"), request.querystring("ladder"))) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQl="update lnk_T_L set Status='Resting', LadderMode=2, ModeFlagTime='" & now & "', Restdays=(RestDays + 1) where LadderID=" & Request.QueryString("ladderid") & " and teamid=" & Request.QueryString("teamid")
		ors.Open strsql, oconn
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "teamladderadmin.asp?ladder=" & Request.QueryString("ladder") & "&team=" & Request.QueryString("team")
	end if
end if
'-----------------------------------------------
' Take Team Off Rest
'-----------------------------------------------
if Request.QueryString("SaveType") = "UnRest" then
	if not(IsSysAdmin()) and not(IsLadderAdmin(request.querystring("ladder"))) and not(IsTeamFounder(request.querystring("team"))) and not(IsTeamCaptain(request.querystring("team"), request.querystring("ladder"))) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		strSQl="update lnk_T_L set Status='Available', LadderMode=0, ModeFlagTime=null where LadderID='" & Request.QueryString("ladderid") & "' and teamid=" & Request.QueryString("teamid")
		ors.Open strsql, oconn
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "teamladderadmin.asp?ladder=" & Request.QueryString("ladder") & "&team=" & Request.QueryString("team")
	end if
end if
'-----------------------------------------------
' Save the Map Listing for a Ladder
'-----------------------------------------------
if Request.Form("SaveType") = "MapList" then
	lid=Request.Form("ladder")
	if lid<> "" then 
		if not(IsSysAdmin()) and not (IsLadderAdminByID(lid)) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
			response.redirect "errorpage.asp?error=3"
		else
			ors.CursorType=adOpenKeyset
			'oconn.BeginTrans 
			strsql="delete from lnk_L_M where ladderid=" & lid
			oconn.Execute strSQL

			strSQL = ""
			For i = 1 To (Request.Form("frm_current_maplist_map0").Count)
				strSQL = strSQL & "INSERT INTO lnk_l_m (MapID, LadderID, MapNumber) VALUES ('" & Request.Form("frm_current_maplist_map0")(i) & "', '" & lid & "', 0); "
			Next
			For i = 1 To (Request.Form("frm_current_maplist_map1").Count)
				strSQL = strSQL & "INSERT INTO lnk_l_m (MapID, LadderID, MapNumber) VALUES ('" & Request.Form("frm_current_maplist_map1")(i) & "', '" & lid & "', 1); "
			Next
			For i = 1 To (Request.Form("frm_current_maplist_map2").Count)
				strSQL = strSQL & "INSERT INTO lnk_l_m (MapID, LadderID, MapNumber) VALUES ('" & Request.Form("frm_current_maplist_map2")(i) & "', '" & lid & "', 2); "
			Next
			For i = 1 To (Request.Form("frm_current_maplist_map3").Count)
				strSQL = strSQL & "INSERT INTO lnk_l_m (MapID, LadderID, MapNumber) VALUES ('" & Request.Form("frm_current_maplist_map3")(i) & "', '" & lid & "', 3); "
			Next
			For i = 1 To (Request.Form("frm_current_maplist_map4").Count)
				strSQL = strSQL & "INSERT INTO lnk_l_m (MapID, LadderID, MapNumber) VALUES ('" & Request.Form("frm_current_maplist_map4")(i) & "', '" & lid & "', 4); "
			Next
			For i = 1 To (Request.Form("frm_current_maplist_map5").Count)
				strSQL = strSQL & "INSERT INTO lnk_l_m (MapID, LadderID, MapNumber) VALUES ('" & Request.Form("frm_current_maplist_map5")(i) & "', '" & lid & "', 5); "
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
			Response.Redirect "/adminops.asp?rAdmin=Ladder"
		end if
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "default.asp"
	end if
end if
'-----------------------------------------------
' Save the Map Listing for a Player Ladder
'-----------------------------------------------
if Request.Form("SaveType") = "MapListPlayer" then
	lid=Request.Form("ladder")
	if lid<> "" then 
		if not(IsSysAdmin()) and not (IsPlayerLadderAdminByID(lid)) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
			response.redirect "errorpage.asp?error=3"
		else
			ors.CursorType=adOpenKeyset
			'oconn.BeginTrans 
			strsql="delete from lnk_pL_M where Playerladderid=" & lid
			oconn.Execute strSQL

			strSQL = ""
			For i = 1 To (Request.Form("frm_current_maplist_map0").Count)
				strSQL = strSQL & "INSERT INTO lnk_pl_m (MapID, PlayerLadderID) VALUES ('" & Request.Form("frm_current_maplist_map0")(i) & "', '" & lid & "'); "
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
			Response.Redirect "adminmenu.asp"
		end if
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "default.asp"
	end if
end if
'-----------------------------------------------
' Change Reported Match
'-----------------------------------------------
if request.form("SaveType") = "ChangeMatch" then 
	if not (IsSysAdmin() or IsLadderAdminByID(request.form("LadderID"))) then 
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	
	LadderID = Request.Form("LadderID")
	HistoryID = Request.Form("HistoryID")
	DefID = Request.Form("DefenderID")
	AttID = Request.Form("AttackerID")
	DefOldRank = Request.Form("DefOldRank")
	DefNewRank = Request.Form("DefNewRank")
	AttOldRank = Request.Form("AttOldRank")
	AttNewRank = Request.Form("AttNewRank")
	Maps = Request.Form("Maps")
	
	Map1 = Request.Form("Map1Name")
	DefMap1Score = Request.Form("Map1DefScore")
	AttMap1Score = Request.Form("Map1AttScore")
	Map1OT = Request.Form("Map1OT")
	Map1FT = Request.Form("Map1FT")

	Map2 = Request.Form("Map2Name")
	DefMap2Score = Request.Form("Map2DefScore")
	AttMap2Score = Request.Form("Map2AttScore")
	Map2OT = Request.Form("Map2OT")
	Map2FT = Request.Form("Map2FT")

	Map3 = Request.Form("Map3Name")
	DefMap3Score = Request.Form("Map3DefScore")
	AttMap3Score = Request.Form("Map3AttScore")
	Map3OT = Request.Form("Map3OT")
	Map3FT = Request.Form("Map3FT")
	
	Map4 = Request.Form("Map4Name")
	DefMap4Score = Request.Form("Map4DefScore")
	AttMap4Score = Request.Form("Map4AttScore")
	Map4OT = Request.Form("Map4OT")
	Map4FT = Request.Form("Map4FT")

	Map5 = Request.Form("Map5Name")
	DefMap5Score = Request.Form("Map5DefScore")
	AttMap5Score = Request.Form("Map5AttScore")
	Map5OT = Request.Form("Map5OT")
	Map5FT = Request.Form("Map5FT")

	OverAllForfeit = Request.Form("OverAllForfeit")
	DefenderWin = Request.Form("DefWin")
	OldDefenderWin = Request.form("DefOldWin")
	
	matchdate = Request.Form("MatchDate")

	If DefenderWin then
		winid = defid
		losid = attid
	else
		winid = attid
		losid = defid
	end if
	if defenderwin <> olddefenderwin then
		strSQL = "update lnk_T_L set Losses=(Losses-1), wins=(wins+1) where TLLinkID=" & winid
		'response.write strsql
		ors.open strsql, oconn
		strSQL = "update lnk_T_L set Losses=(Losses+1), wins=(wins-1) where TLLinkID=" & losid
		'response.write strsql
		ors.open strsql, oconn
							
	end if

	strSQL = "update tbl_History set MatchWinnerID='" & WinID
	strsql = strsql & "', MatchLoserID='" & losID
	strsql = strsql & "', MatchMap1='" & replace(map1, "'", "''")
	strsql = strsql & "', MatchMap2='" & replace(map2, "'", "''")
	strsql = strsql & "', MatchMap3='" & replace(map3, "'", "''")
	strsql = strsql & "', MatchMap4='" & replace(map4, "'", "''")
	strsql = strsql & "', MatchMap5='" & replace(map5, "'", "''")
	strsql = strsql & "', MatchMap1DefenderScore='" & DefMap1Score
	strsql = strsql & "', MatchMap1AttackerScore='" & AttMap1Score
	strsql = strsql & "', MatchMap2DefenderScore='" & DefMap2Score
	strsql = strsql & "', MatchMap2AttackerScore='" & AttMap2Score
	strsql = strsql & "', MatchMap3DefenderScore='" & DefMap3Score
	strsql = strsql & "', MatchMap3AttackerScore='" & AttMap3Score
	strsql = strsql & "', MatchMap4DefenderScore='" & DefMap4Score
	strsql = strsql & "', MatchMap4AttackerScore='" & AttMap4Score
	strsql = strsql & "', MatchMap5DefenderScore='" & DefMap5Score
	strsql = strsql & "', MatchMap5AttackerScore='" & AttMap5Score
	strsql = strsql & "', MatchMap1OT='" & Map1OT
	strsql = strsql & "', MatchMap2OT='" & Map2OT
	strsql = strsql & "', MatchMap3OT='" & Map3OT
	strsql = strsql & "', MatchMap4OT='" & Map4OT
	strsql = strsql & "', MatchMap5OT='" & Map5OT
	strsql = strsql & "', MatchMap1Forfeit='" & Map1FT
	strsql = strsql & "', MatchMap2Forfeit='" & Map2FT
	strsql = strsql & "', MatchMap3Forfeit='" & Map3FT
	strsql = strsql & "', MatchMap4Forfeit='" & Map4FT
	strsql = strsql & "', MatchMap5Forfeit='" & Map5FT
	strsql = strsql & "', MatchDate='" & replace(MatchDate, "'", "''")
	strsql = strsql & "', MatchWinnerDefending='" & DefenderWin
	strsql = strsql & "' where HistoryID='" & HistoryID & "'"
		
	ors.open strsql, oConn
	'Response.Write "<font color=white>" & strsql
	if defnewrank <> defoldrank then
		'response.write "<br>Defender getting a new rank."
		if defnewrank < defoldrank then 
			'response.write "Defender moving up in the world"
			strsql="update lnk_t_L set rank=(rank+1) where (rank > " & defnewrank - 1 & "  and rank < " & defoldrank & ") and isactive=1 and ladderid = " & ladderid
		else
			'response.write "Defender gets a demotion"
			strsql="update lnk_t_L set rank=(rank-1) where (rank < " & defnewrank + 1 & "  and rank > " & defoldrank & ") and isactive=1 and ladderid = " & ladderid
		end if
		ors.open strsql, oconn
		'response.write "<br>" & strsql			
		strsql = "update lnk_t_l set rank=" & defnewrank & " where TLLinkID=" & defid
		ors.open strsql, oconn
		'response.write "<br>" & strsql
	end if
	if attnewrank <> attoldrank then
		'response.write "<br>Attacker getting a new rank."
		if attnewrank < attoldrank then 
			'response.write "Attacker moving up in the world"
			strsql="update lnk_t_L set rank=(rank+1) where (rank > " & attnewrank - 1 & "  and rank < " & attoldrank & ") and isactive=1 and ladderid = " & ladderid
		else
			'response.write "Attacker gets a demotion"
			strsql="update lnk_t_L set rank=(rank-1) where (rank < " & attnewrank + 1 & "  and rank > " & attoldrank & ") and isactive=1 and ladderid = " & ladderid
		end if
		ors.open strsql, oconn
		'response.write "<br>" & strsql
		strsql = "update lnk_t_l set rank=" & attnewrank & " where TLLinkID=" & attid
		ors.open strsql, oconn
		'response.write "<br>" & strsql
	
	end if

	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "AdminMenu.asp"
end if
'-----------------------------------------------
' Rerank
'-----------------------------------------------
if request.form("SaveType") = "ChangeRank" then 
	if not (IsSysAdmin() or IsLadderAdminByID(request.form("LadderID"))) then 
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if

	tllinkid = request.form("tllinkid")
	ladderid = request.form("LadderID")
	newrank = cInt(request.form("NewRank"))

	if newrank <> "" then
		strsql = "select rank from lnk_T_L where TLLinkID='" & tllinkid	& "'"
		'response.write strsql
		ors.open strsql, oconn
		if not (ors.eof and ors.bof) then
			oldrank = cint(ors.fields(0).value)
			ors.close
			if newrank <> oldrank then
				'response.write "<br>Team getting a new rank."
				if newrank < oldrank then 
					'response.write "Team moving up in the world"
					strsql="update lnk_t_L set rank=(rank+1) where (rank >= " & newrank & "  and rank < " & oldrank & ") and isactive=1 and ladderid = " & ladderid
'					response.write strsql
'					response.end
				else
					'response.write "Team gets a demotion"
					strsql="update lnk_t_L set rank=(rank-1) where (rank <= " & newrank & "  and rank > " & oldrank & ") and isactive=1 and ladderid = " & ladderid
				end if
			end if
			ors.open strsql, oconn
			'response.write "<br>" & strsql
			strsql = "update lnk_t_l set rank=" & newrank & " where TLLinkID=" & tllinkid
			ors.open strsql, oconn
			'response.write "<br>" & strsql
		else
			ors.close
		end if	
	end if

	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "AdminMenu.asp"
end if
'-----------------------------------------------
' User Viewing Cookies
'-----------------------------------------------
if request.form("SaveType") = "SetCookie" then 
	lview = request.form("LadderView")
	tview = request.form("TeamPage")
	mview = request.form("MemberPage")
	sig = Request.Form("showsig")
	sigs = Request.Form("ShowSigs")
	if Len(sigs) > 0 Then
		strSQL = "UPDATE tbl_Players SET ShowForumSigs = 1, StyleID = '" & CheckString(Request.Form("radStyle")) & "' WHERE PlayerID = '" & Session("PlayerID") & "'"
		Session("ShowSigs") = 1
	Else
		strSQL = "UPDATE tbl_Players SET ShowForumSigs = 0, StyleID = '" & CheckString(Request.Form("radStyle")) & "'  WHERE PlayerID = '" & Session("PlayerID") & "'"
		Session("ShowSigs") = 0
	End If
	oConn.Execute(strSQL)
	
	Session("StyleID") = Request.Form("radStyle")
	if sig = "ShowSigAuto" then
		sigverb = "y"
	else
		sigverb = "n"
	end if
	if lview <= 0 or lview = "" then
		lview= 25
	end if
	if tview <= 0 or tview = "" then
		tview= 40
	end if
	if mview <= 0 or mview = "" then
		mview = 40
	end if

	Response.Cookies("PerPage")("ShowSig") = sigverb
	Response.Cookies("PerPage")("LadderView") = lview
	Response.Cookies("PerPage")("MemberView") = mview
	Response.Cookies("PerPage")("TeamView") = tview
	Response.Cookies("PerPage").expires = "1/1/2015"
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	response.redirect "preferences.asp"
end if
'-----------------------------------------------
' E-mail People
'-----------------------------------------------
if request.form("Savetype") = "SendMail" then
	if not(IsSysAdmin() or IsAnyLadderAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=3"
	end if
	
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.RemoteHost  = "127.0.0.1"
	Mailer.FromName    = "TeamWarfare Staff"
	mailer.FromAddress = "automailer@web.teamwarfare.com"
	mailer.AddRecipient "TWL", "automailer@web.teamwarfare.com"
'Response.Write "<BR>"127.0.0.1"" & "127.0.0.1"
'Response.Write "<BR>FromName" & "TeamWarfare Staff"
'Response.Write "<BR>FromAddress" & request.form("email")
'---
' Email SysAdmins Regardless
'---
	strsql = "Select p.PlayerHandle, p.PlayerEmail FROM sysadmins s, tbl_players p where p.playerid = s.adminid and s.sendemail=1"
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		do while not(ors.eof)
			Mailer.AddBCC ors.fields(0).value, ors.fields(1).value
			'Response.Write ors.fields(0).value & "-" & ors.fields(1).value & "<BR>"
			ors.movenext
		loop
	end if
	ors.close
'---
' Decide who else to email
'---	
	select case request.form("mailto")
		case "lad"
			strsql = "select * from lnk_l_a"
			ors.open strsql, oconn
			if not (ors.eof and ors.bof) then
				do while not (ors.eof) 
					strsql = "select PlayerHandle, PlayerEmail from tbl_players where playerid='" & ors.fields("PlayerID").value & "'"
					ors2.open strsql, oconn
					if not (ors2.eof and ors2.bof) then
						Mailer.AddBCC ors2.fields(0).value, ors2.fields(1).value
					end if
					ors2.close
					ors.movenext
				loop
			end if
			ors.close
		case else
			ladderid = request.form("mailto")
			If left(ladderID, 1) = "p" then
				strSQL = "select PlayerHandle, PlayerEmail FROM lnk_p_pl lnk, tbl_players p WHERE p.PlayerID = lnk.PlayerID and lnk.PlayerLadderID = '" & Right(ladderid, len(ladderid) - 1) & "'"
				ors.open strsql, oconn
				if not (ors.eof and ors.bof) then
					do while not (ors.eof) 
						Mailer.AddBCC ors.fields(0).value, ors.fields(1).value
						ors.movenext
					loop
				end if
				ors.close
			ElseIf left(ladderID, 1) = "t" then
				ladderid = Right(ladderid, len(ladderid) - 1) 
				if len(trim(ladderid)) > 0 and IsNumeric(ladderID) Then
					strsql = "SELECT p.PlayerHandle, p.PlayerEmail FROM lnk_t_l ltl, lnk_t_p_l ltpl, tbl_players p "
					strsql = strSQL & " WHERE ltl.LadderID = '" & ladderid & "' "
					strsql = strSQL & " AND isActive = 1 "
					strsql = strSQL & " AND ltl.TLLinkID = ltpl.TLLinkID "
					strsql = strSQL & " AND ltpl.PlayerID = p.PlayerID "
					strsql = strSQL & " AND ltpl.IsAdmin = 1"
					ors.open strsql, oconn
					if not (ors.eof and ors.bof) then
						do while not (ors.eof)
							Mailer.AddBCC ors.fields(0).value, ors.fields(1).value
							'Response.Write ors.fields(0).value & "-" & ors.fields(1).value & "<BR>"
							ors.movenext
						loop
					end if
					ors.close
				end if
			ElseIf left(ladderID, 1) = "l" then
				ladderid = Right(ladderid, len(ladderid) - 1) 
				if len(trim(ladderid)) > 0 and IsNumeric(ladderID) Then
					strsql = "SELECT p.PlayerHandle, p.PlayerEmail FROM lnk_league_team lnkT, lnk_league_team_player lnkP, tbl_players p "
					strsql = strSQL & " WHERE lnkT.LeagueID = '" & ladderid & "' "
					strsql = strSQL & " AND lnkT.Active = 1 "
					strsql = strSQL & " AND lnkT.lnkLeagueTeamID = lnkP.lnkLeagueTeamID "
					strsql = strSQL & " AND lnkP.PlayerID = p.PlayerID "
					strsql = strSQL & " AND lnkP.IsAdmin = 1 "
					response.write "<br>" & strSQL & "<br>"
					ors.open strsql, oconn
					if not (ors.eof and ors.bof) then
						do while not (ors.eof)
							Mailer.AddBCC ors.fields(0).value, ors.fields(1).value
							'Response.Write ors.fields(0).value & "-" & ors.fields(1).value & "<BR>"
							ors.movenext
						loop
					end if
					ors.close
				end if
			ElseIf left(ladderID, 1) = "a" then
				ladderid = Right(ladderid, len(ladderid) - 1) 
				if len(trim(ladderid)) > 0 and IsNumeric(ladderID) Then
					strsql = "SELECT p.PlayerHandle, p.PlayerEmail FROM lnk_t_m lnkT, lnk_t_m_p lnkP, tbl_players p "
					strsql = strSQL & " WHERE lnkT.TournamentID = '" & ladderid & "' "
'					strsql = strSQL & " AND lnkT.Active = 1 "
					strsql = strSQL & " AND lnkT.TMLinkID = lnkP.TMLinkID "
					strsql = strSQL & " AND lnkP.PlayerID = p.PlayerID "
					strsql = strSQL & " AND lnkP.IsAdmin = 1 "
					'response.write "<br>" & strSQL & "<br>"
					ors.open strsql, oconn
					if not (ors.eof and ors.bof) then
						do while not (ors.eof)
							Mailer.AddBCC ors.fields(0).value, ors.fields(1).value
							'Response.Write ors.fields(0).value & "-" & ors.fields(1).value & "<BR>"
							ors.movenext
						loop
					end if
					ors.close
				end if
			End if
	end select
	Mailer.Subject     = request.form("Subject")
	Text = Request.form("MailBody")
	Mailer.BodyText    = text
	on error resume next
	If Mailer.SendMail Then
		'Response.Write "Sent"
	Else
		'Response.Write "problem"
	End If
	on error goto 0
	set mailer = nothing
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	response.redirect "/adminmenu.asp"

end if
'-----------------------------------------------
' Rerank
'-----------------------------------------------
if request.form("SaveType") = "ChangePlayerRank" then 
	if not (IsSysAdmin() or IsPlayerLadderAdminByID(request.form("LadderID"))) then 
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if

	PPLLinkID = request.form("PPLLinkID")
	ladderid = request.form("LadderID")
	newrank = request.form("NewRank")

	if newrank <> "" then
		strsql = "select rank from lnk_p_pL where PPLLinkID='" & PPllinkid	& "'"
		'response.write strsql
		ors.open strsql, oconn
		if not (ors.eof and ors.bof) then
			oldrank = ors.fields(0).value
			ors.close
			if newrank <> oldrank then
'				response.write "<br>Player getting a new rank."
				if newrank < oldrank then 
'					response.write "Team moving up in the world"
'					strsql="update lnk_t_L set rank=(rank+1) where (rank > " & newrank - 1 & "  and rank < " & oldrank & ") and isactive=1 and ladderid = " & ladderid
					strsql="update lnk_p_pl set rank=(rank+1) where (rank > " & newrank - 1 & "  and rank < " & oldrank & ") and isactive=1 and playerladderid = " & ladderid
				else
'					response.write "Team gets a demotion"
'					strsql="update lnk_t_L set rank=(rank-1) where (rank < " & newrank + 1 & "  and rank > " & oldrank & ") and isactive=1 and ladderid = " & ladderid
					strsql="update lnk_p_pl set rank=(rank-1) where (rank < " & newrank + 1 & "  and rank > " & oldrank & ") and isactive=1 and playerladderid = " & ladderid
				end if
			end if
			ors.open strsql, oconn
			'response.write "<br>" & strsql
			strsql = "update lnk_p_pl set rank=" & newrank & " where PPllinkid=" & PPllinkid
			ors.open strsql, oconn
'			response.write "<br>" & strsql
		else
			ors.close
		end if	
	end if

	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "AdminMenu.asp"
end if
'-----------------------------------------------
' Add a staff member / edit a staff member
'-----------------------------------------------
if request.form("SaveType") = "StaffMember" then 
	if not IsSysAdmin() then 
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	intStaffGroupID = Request.Form("StaffGroup")
	strDisplayName =Request.Form("DisplayName")
	strTitle = Request.Form("Title")
	strDescription = Request.Form("Description")
	strEmail = Request.Form("email")
	intSeqNum = Request.Form("seqNum")

	if request.form("savemethod")="Edit" then
		strSQL = "UPDATE tbl_staff SET "
		strSQL = strSQL & "StaffGroupID = '" & CheckString(intStaffGroupID) & "', "
		strSQL = strSQL & "DisplayName = '" & CheckString(strDisplayName) & "', "
		strSQL = strSQL & "Title = '" & CheckString(strTitle) & "', "
		strSQL = strSQL & "Description = '" & CheckString(strDescription) & "', "
		strSQL = strSQL & "Email = '" & CheckString(strEmail) & "', "
		strSQL = strSQL & "SeqNum = '" & CheckString(intSeqNum) & "' "
		strSQL = strSQL & " WHERE StaffID = " & Request.Form("StaffID")
		
		oConn.Execute (strSQL)
	else
		strSQL = "INSERT INTO tbl_staff ( StaffGroupID, DisplayName, Title, Description, Email, SeqNum ) VALUES ( "
		strSQL = strSQL & "'" & CheckString(intStaffGroupID) & "', "
		strSQL = strSQL & "'" & CheckString(strDisplayName) & "', "
		strSQL = strSQL & "'" & CheckString(strTitle) & "', "
		strSQL = strSQL & "'" & CheckString(strDescription) & "', "
		strSQL = strSQL & "'" & CheckString(strEmail) & "', "
		strSQL = strSQL & "'" & CheckString(intSeqNum) & "') "
		oConn.Execute (strSQL)
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "staff.asp"
End If

'-----------------------------------------------
' Delete a match from history
'-----------------------------------------------
if Request.Form("SaveType") = "DeleteHistory" Then
	HistoryID = Request.Form("HistoryID")
	If Not(bSysAdmin) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	End If
	strSQL = "SELECT MatchWinnerID, MatchLoserID, MatchForfeit FROM vHistory WHERE HistoryID='" & HistoryID & "'"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		WinnerLinkID = oRS.Fields("MatchWinnerID").Value 
		LoserLinkID = oRS.Fields("MatchLoserID").Value 
		Forfeit = oRS.Fields("MatchForfeit").Value 
		oRs.Close
		If cBool(forfeit) Then
			strSQL = "UPDATE lnk_t_l SET Forfeits = ForFeits - 1 WHERE TLLinkID='" & LoserLinkID & "';DELETE FROM tbl_history WHERE HistoryID='" & HistoryID & "'"
			oConn.Execute strSQL
		Else
			strSQL = "UPDATE lnk_t_l SET Losses = Losses - 1 WHERE TLLinkID='" & LoserLinkID & "'"
			strSQL = strSQL & "UPDATE lnk_t_l SET Wins = Wins - 1 WHERE TLLinkID='" & WinnerLinkID & "';"
			strSQL = strSQL & "DELETE FROM tbl_history WHERE HistoryID='" & HistoryID & "'"
			oConn.Execute strSQL
		End If
	End If
	Response.Redirect "/edithistory.asp?ladder=" & Server.URLEncode(Request.Form("Ladder"))
End If

'-----------------------------------------------
' Remove Ladder Admin
'-----------------------------------------------
If Request.QueryString("SaveType") = "DeleteAdmin" Then
	intLALinkID = Request.Querystring("LALinkID")
	strSQL = "DELETE FROM lnk_l_a WHERE LALinkID = '" & intLALinkID & "'"
	oConn.Execute(strSQL)
	Response.Clear
	Response.Redirect "/assignadmin.asp#" & Request.Querystring("LadderID")
End If

'-----------------------------------------------
' Make Primary Admin
'-----------------------------------------------
If Request.QueryString("SaveType") = "PrimaryAdmin" Then
	intLALinkID = Request.Querystring("LALinkID")
	intLadderID = Request.Querystring("LadderID")
	strSQL = "UPDATE lnk_l_a SET PrimaryAdmin = 0 WHERE LadderID = '" & intLadderID & "'"
	oConn.Execute(strSQL)
	strSQL = "UPDATE lnk_l_a SET PrimaryAdmin = 1 WHERE LALinkID = '" & intLALinkID & "'"
	oConn.Execute(strSQL)
	Response.Clear
	Response.Redirect "/assignadmin.asp#" & Request.Querystring("LadderID")
End If

'-----------------------------------------------
' IP Ban Code!
'-----------------------------------------------
If Request.QueryString ("SaveType") = "AddIPBan" Then
	If Len(Request.QueryString ("IP")) > 0 Then
		strSQL = "EXECUTE AddIPBan '" & CheckString(Request.QueryString("IP")) & "'"
		oConn.Execute(strSQL)
	End If
	oConn.Close
	On Error Goto 0
	set ors = nothing
	set oConn = nothing	
	set ors2 = nothing	
	Response.Clear
	Response.Redirect "/ipban.asp"
End If
'-----------------------------------------------
' IP Ban Code!
'-----------------------------------------------
If Request.QueryString ("SaveType") = "DeleteIPBan" Then
	If Len(Request.QueryString ("IP")) > 0 Then
		strSQL = "EXECUTE DeleteIPBan '" & CheckString(Request.QueryString("IP")) & "'"
		oConn.Execute(strSQL)
	End If
	oConn.Close
	On Error Goto 0
	set ors = nothing
	set oConn = nothing	
	set ors2 = nothing	
	Response.Clear
	Response.Redirect "/ipban.asp"
End If

'-----------------------------------------------
' new / edit Games
'-----------------------------------------------
if request.form("SaveType") = "Games" then 
	if not IsSysAdmin() then 
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strGameName = Request.Form("GameName")
	strGameAbbr = Request.Form("GameAbbreviation")
	intGameID = Request.Form("GameID")
	intForumID = Request.Form("ForumID")
	intDisputeForumID = Request.Form("DisputeForumID")
	if request.form("savemethod")="Edit" then
		strSQL = "UPDATE tbl_games SET "
		strSQL = strSQL & " GameName = '" & CheckString(strGameName) & "', "
		strSQL = strSQL & " GameAbbreviation = '" & CheckString(strGameAbbr) & "', "
		strSQL = strSQL & " ForumID = '" & CheckString(intForumID) & "', "
		strSQL = strSQL & " DisputeForumID = '" & CheckString(intDisputeForumID) & "' "
		strSQL = strSQL & " WHERE GameID = '" & intGameID & "'"
		oConn.Execute(strSQL)
	Else
		strSQL = "INSERT INTO tbl_games (GameName, GameAbbreviation, ForumID, DisputeForumID) VALUES ("
		strSQL = strSQL & "'" & CheckString(strGameName) & "', "
		strSQL = strSQL & "'" & CheckString(strGameAbbr) & "', "
		strSQL = strSQL & "'" & CheckString(intForumID) & "', "
		strSQL = strSQL & "'" & CheckString(DisputeForumID) & "') "
		
		oConn.Execute(strSQL)
	end if
	oConn.Close
	Set oRS = Nothing
	Set oConn = Nothing	
	Response.Clear
	Response.Redirect "gamelist.asp"
end if
'-----------------------------------------------
' AssignDivision on League
'-----------------------------------------------
if Request.Form("SaveType")="LeagueAssignDivision" then
	intLeagueID = Request.Form("LeagueID")
	if not(bSysAdmin or IsLeagueAdminById(intLeagueID)) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		intLeagueTeamID = Request.Form("lnkLeagueTeamID")
		intDivisionID = Request.Form("selDivisionID")
		intConferenceID = Request.Form("selConferenceID")
		
		strSQL = "SELECT tbl_teams.TeamID, TeamFounderID, TeamName, TeamEmail "
		strSQL = strSQL & " FROM tbl_Teams "
		strSQL = strSQL & " INNER JOIN lnk_league_team "
		strSQL = strSQL & " ON lnk_league_team.TeamID = tbl_teams.TeamID "
		strSQL = strSQL & " WHERE lnk_league_team.lnkLeagueTeamID='" & intLeagueTeamID & "'"
		oRs.Open strSQL, oConn
		If Not(ors.eof and ors.bof) Then
			teamname=ors.fields(2).value
			tid=ors.fields(0).value
			ownerid=ors.fields(1).value
			temail=ors.Fields(3).Value
		Else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=7"
		End If
		ors.close
		strSQL = "SELECT l.LeagueID, l.LeagueName, ConferenceName, d.LeagueConferenceID, DivisionName "
		strSQL = strSQL & "FROM tbl_leagues l, tbl_league_conferences c, tbl_league_divIsions d "
		strSQL = strSQL & "WHERE l.LeagueID = C.LeagueID AND d.LeagueConferenceID=c.LeagueConferenceID "
		strSQL = strSQL & " AND LeagueDivisionID='" & CheckString(intDivisionID) & "'"
		oRs.Open strSQL, oConn
		if not (ors.eof and ors.bof) then
			intLeagueID = ors.fields("LeagueID").value
			strLeagueName = ors.fields("LeagueName").Value
			strConferenceName = oRs.Fields("ConferenceName").Value
			intConferenceID = ors.fields("LeagueConferenceid").value
			strDivisionName = oRS.Fields("DivisionName").Value
		Else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=7"
		End If
		ors.close
		
		strSQL = "EXECUTE LeagueAssignDivision @LeagueDivisionID='" & intDivisionID & "', @lnkLeagueTeamID='" & intLeagueTeamID & "'"
		oConn.Execute(strSQL)
		
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.RemoteHost  = "127.0.0.1"
		Mailer.FromName    = "TWL: League Notification"
		mailer.FromAddress = "automailer@teamwarfare.com"
		Mailer.AddRecipient teamName, tEmail
		Mailer.Subject     = "TWL: " & teamName & " accepted into " & strLeagueName & " League in " & strConferenceName & " Conference on " & strDivisionName & " Division"
		Text = teamName & ", your team been accepted into the " & strLeagueName  & " League in " & strConferenceName & " Conference on " & strDivisionName & " Division. You should take this time to review the rules for this league on www.teamwarfare.com." & vbcrlf
		Text = text & "This is an information e-mail only, please do not reply to this message."
		Mailer.BodyText    = text
		on error resume next
		Mailer.SendMail
		on error goto 0
		set mailer = nothing

		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "leagueassign.asp?league=" & Server.URLEncode(strLeagueName)
	end if
end if
'-----------------------------------------------
' LeagueDecline
'-----------------------------------------------
if Request.Form("SaveType")="LeagueDecline" then
	intLeagueID = Request.Form("LeagueID")
	if not(bSysAdmin or IsLeagueAdminById(intLeagueID)) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		intLeagueTeamID = Request.Form("lnkLeagueTeamID")
		intDivisionID = Request.Form("selDivisionID")
		intConferenceID = Request.Form("selConferenceID")
		
		strSQL = "SELECT tbl_teams.TeamID, TeamFounderID, TeamName, TeamEmail "
		strSQL = strSQL & " FROM tbl_Teams "
		strSQL = strSQL & " INNER JOIN lnk_league_team "
		strSQL = strSQL & " ON lnk_league_team.TeamID = tbl_teams.TeamID "
		strSQL = strSQL & " WHERE lnk_league_team.lnkLeagueTeamID='" & intLeagueTeamID & "'"
		oRs.Open strSQL, oConn
		If Not(ors.eof and ors.bof) Then
			teamname=ors.fields(2).value
			tid=ors.fields(0).value
			ownerid=ors.fields(1).value
			temail=ors.Fields(3).Value
		Else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=7"
		End If
		ors.close
		strSQL = "SELECT l.LeagueID, l.LeagueName "
		strSQL = strSQL & "FROM tbl_leagues l "
		strSQL = strSQL & "WHERE LeagueID='" & CheckString(intLeagueID) & "'"
		oRs.Open strSQL, oConn
		if not (ors.eof and ors.bof) then
			intLeagueID = ors.fields("LeagueID").value
			strLeagueName = ors.fields("LeagueName").Value
		Else
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=7"
		End If
		ors.close
		
		strSQL = "EXECUTE LeagueDeclineAdmittance @lnkLeagueTeamID='" & intLeagueTeamID & "'"
		oConn.Execute(strSQL)
		
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.RemoteHost  = "127.0.0.1"
		Mailer.FromName    = "TWL: League Notification"
		mailer.FromAddress = "automailer@web.teamwarfare.com"
		Mailer.AddRecipient teamName, tEmail
		Mailer.Subject     = "TWL: " & teamName & " league application declined for " & strLeagueName & " League"
		Text = teamName & ", your team been declined admittance into the " & LeagueName  & " League on www.teamwarfare.com. Contact your admin for details." & vbcrlf
		Text = text & "This is an information e-mail only, please do not reply to this message."
		Mailer.BodyText    = text
		on error resume next
		Mailer.SendMail
		on error goto 0
		set mailer = nothing

		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "leagueassign.asp?league=" & Server.URLEncode(strLeagueName)
	end if
end if
'-----------------------------------------------
' Save the Map Listing for a League
'-----------------------------------------------
if Request.Form("SaveType") = "LeagueMapList" then
	lid=Request.Form("LeagueID")
	if lid<> "" then 
		if not(bSysAdmin or IsLeagueAdminById(lid)) then
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "errorpage.asp?error=3"
		else
			ors.CursorType=adOpenKeyset
			'oconn.BeginTrans 
			strsql="delete from lnk_league_maps where LeagueID=" & lid
			oconn.Execute strSQL

			strSQL = ""
			For i = 1 To (Request.Form("frm_current_maplist_map0").Count)
				strSQL = strSQL & "INSERT INTO lnk_league_maps (MapID, LeagueID) VALUES ('" & CheckString(Request.Form("frm_current_maplist_map0")(i)) & "', '" & CheckString(lid) & "'); "
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
			Response.Redirect "leagueadmin.asp"
		end if
	else
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "default.asp"
	end if
end if
'-----------------------------------------------
' LeagueAddMatch
'-----------------------------------------------
if Request.Form("SaveType") = "LeagueAddMatch" then
	intLeagueID = Request.Form("LeagueID")
	If Not(bSysAdmin OR IsLeagueAdminById(intLeagueID)) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	Else
		intHomeLeagueTeamID = Request.Form("selHomeLeagueTeamID")
		intVisitorLeagueTeamID = Request.Form("selVisitorLeagueTeamID")
		strMatchDate = Request.Form("txtMatchDate")
		strMap1 = Request.Form("selMap1")
		strMap2 = Request.Form("selMap2")
		strMap3 = Request.Form("selMap3")
		strMap4 = Request.Form("selMap4")
		strMap5 = Request.Form("selMap5")
		If Len(intVisitorLeagueTeamID) = 0 OR Len(intHomeLeagueTeamID) = 0 Then
			oConn.Close
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			response.redirect "leaguematches.asp?league=" & Server.URLEncode(Request.Form("League"))
		End If		
		strSQL = "EXECUTE LeagueScheduleMatch "
		strSQL = strSQL & " @LeagueID = '" & CheckString(intLeagueID) & "', "
		strSQL = strSQL & " @HomeLeagueTeamID = '" & CheckString(intHomeLeagueTeamID) & "', "
		strSQL = strSQL & " @VisitorLeagueTeamID = '" & CheckString(intVisitorLeagueTeamID) & "', "
		strSQL = strSQL & " @MatchDate = '" & CheckString(strMatchDate) & "', "
		strSQL = strSQL & " @Map1 = '" & CheckString(strMap1) & "', "
		strSQL = strSQL & " @Map2 = '" & CheckString(strMap2) & "', "
		strSQL = strSQL & " @Map3 = '" & CheckString(strMap3) & "', "
		strSQL = strSQL & " @Map4 = '" & CheckString(strMap4)& "', "
		strSQL = strSQL & " @Map5 = '" & CheckString(strMap5) & "' "
		oConn.Execute(strSQL)
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "leaguematches.asp?league=" & Server.URLEncode(Request.Form("League")) & "&matchdate=" & Server.URLEncode(strMatchDate)
	End If
end if
'-----------------------------------------------
' Promote to Captain
'-----------------------------------------------
if Request.Form("SaveType") = "PromoteLeagueCaptain" then

	
	strSQL = "SELECT lnk.LeagueID, lnk.lnkLeagueTeamID FROM lnk_league_team lnk INNER JOIN lnk_league_team_player l ON l.lnkLeagueTeamID = lnk.lnkLeagueTeamID WHERE lnkLeagueTeamPlayerID='" & Request.Form("PlayerList") & "'"
	oRs.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		intLeagueID = oRs.FieldS("LeagueID").Value
		intLinkID = oRs.Fields("lnkLeagueTeamID").Value
	End If
	oRS.NextRecordSet
		
	If Not(bSysAdmin OR IsLeagueAdminByID(intLeagueID) OR IsLeagueTeamCaptainByLinkID(intLinkID) OR IsTeamFounder(Request.Form("Team"))) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	ENd If

	strSQL="UPDATE lnk_league_team_player set isadmin=1 where lnkLeagueTeamPlayerID='" & Request.Form("playerlist") & "'"
	'Response.write strSQL
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "TeamLeagueAdmin.asp?league=" & server.urlencode(Request.form("league")) & "&team=" & server.urlencode(Request.Form("team"))
end if
'-----------------------------------------------
' DemoteLeagueCaptain
'-----------------------------------------------
if Request.Form("SaveType") = "DemoteLeagueCaptain" then
	strSQL="select lnk.lnkLeagueTeamID, playerid from lnk_league_team_player lnk where lnk.lnkLeagueTeamPlayerID='" & Request.Form("playerlist") & "'"
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		intLinkID=ors.Fields(0).Value
		intPlayerID=ors.Fields(1).Value 
	end if
	ors.Close
	strSQL="select teamfounderid, lnk.LeagueID from tbl_teams inner join lnk_league_team lnk on lnk.teamid=tbl_teams.teamid where lnk.lnkLeagueTeamID='" & intLinkID & "'"
	ors.open strsql, oconn
	if not (ors.EOF and ors.BOF) then
		intFounderID=ors.Fields(0).Value
		intLeagueID = oRS.FIelds("LeagueID").Value
	end if
	ors.Close 

	If Not(bSysAdmin OR IsLeagueAdminByID(intLeagueID) OR IsLeagueTeamCaptainByLinkID(intLinkID) OR IsTeamFounder(Request.Form("Team"))) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	ENd If

	if intFounderID <> intPlayerID then
		strSQL="UPDATE lnk_league_team_player set isadmin=0 where lnkLeagueTeamPlayerID='" & Request.Form("playerlist") & "'"
		ors.open strSQL, oconn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "TeamLeagueAdmin.asp?league=" & server.urlencode(Request.form("league")) & "&team=" & server.urlencode(Request.Form("team"))
end if
'-----------------------------------------------
' Kick Player from League Roster
'-----------------------------------------------
if request.form("savetype")="DropLeaguePlayer" then
	intLinkID=request.form("link")
	playerid=request.form("playerid")
		
	strsql="select leagueid, teamid from lnk_league_team where lnkLeagueTeamID=" & intLinkID
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		tid = ors.fields(1).value
		lid = ors.fields(0).value
	end if
	ors.close

	If Not(bSysAdmin OR IsLeagueAdminByID(lid) OR IsLeagueTeamCaptainByLinkID(intLinkID)) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	else
		if playerid <> "" then
			strsql="delete from lnk_league_team_player where lnkLeagueTeamPlayerID=" & playerid
			oConn.Execute(strSQL)
		end if
		'response.write Server.HTMLEncode(tname) & Server.HTMLEncode(lname)
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
	Response.Redirect "TeamLeagueAdmin.asp?league=" & server.urlencode(Request.form("league")) & "&team=" & server.urlencode(Request.Form("team"))
	end if
end if
'-----------------------------------------------
' Add League Comms
'-----------------------------------------------
if Request.form("SaveType") = "AddLeagueCommunications" then
	strSQL = "SELECT LeagueID, HomeTeamLinkID, VisitorTeamLinkID FROM tbl_league_matches WHERE LeagueMatchID = '" & Request.Form("MatchID") & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intHomeLinkID = oRS.Fields("HomeTeamLInkID").Value
		intVisitorLinkID = oRS.Fields("VisitorTeamLinkID").Value
		intLeagueID = oRs.Fields("LeagueID").Value
	End If
	oRS.NextRecordSet
		
	If Not(bSysAdmin OR IsLeagueAdminByID(intLeagueID) OR IsLeagueTeamCaptainByLinkID(intHomeLinkID) OR IsLeagueTeamCaptainByLinkID(intVisitorLinkID)) Then
'	Response.End
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	End If

	strSQL = "INSERT INTO tbl_league_Comms ( LeaguematchID, CommDate, CommAuthor, Comms ) values ('" 
	strSQL = strSQL & Request.form("matchID") & "',GetDate(),'"  & replace(Request.Form("commauthor"), "'", "''") & "',"
	strSQL = strSQL & "'" & replace(Request.Form("comms"), "'", "''") & "')"
	oConn.Execute(strSQL)
	strSQL = "UPDATE tbl_league_matches SET CommsCount = CommsCount + 1, LastCommDate = GetDate(), LastCommAuthor = '" & CheckString(Request.Form("CommAuthor")) & "' WHERE LeagueMatchID='" & CheckString(Request.Form("MatchID")) & "'"
	oConn.Execute(strSQL)
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/TeamLeagueAdmin.asp?league=" & server.urlencode(request("league")) & "&team=" & server.urlencode(request("Team")) & "&matchid=" & request.form("matchid")
end if
'-----------------------------------------------
' Edit League Comms
'-----------------------------------------------
if Request.form("SaveType") = "EditLeagueCommunications" then
	if not(bSysAdmin) then		
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL= "update tbl_league_comms set Comms='" & replace(Request.Form("comms"), "'", "''") &  "' where leaguecommid=" & Request.Form("leaguecommid")
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/TeamLeagueAdmin.asp?league=" & server.urlencode(request("league")) & "&team=" & server.urlencode(request("Team")) & "&matchid=" & request.form("matchid")
end if
'-----------------------------------------------
' Delete League Comms
'-----------------------------------------------
if Request.QueryString("SaveType") = "DeleteLeagueCommunications" then
	if not(bSysAdmin) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	if Request.QueryString("leaguecommid") <> "" then
		strSQL= "delete from tbl_league_comms where leaguecommid=" & Request.QueryString("leaguecommid")
		'Response.Write strSQl
		oRs.Open strSQL, oConn
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/TeamLeagueAdmin.asp?league=" & server.urlencode(request("league")) & "&team=" & server.urlencode(request("Team")) & "&matchid=" & request("matchid")
end if
'-----------------------------------------------
' LeagueAssignAdmin
'-----------------------------------------------
if Request.Form("SaveType") = "LeagueAssignAdmin" then
	if not(bSysAdmin) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	If Len(Request.Form("PlayerID")) > 0 Then
		strSQL = "INSERT INTO lnk_league_admin (PlayerID, LeagueID, LeagueConferenceID, LeagueDivisionID) VALUES ('" & Request.Form("PlayerID") & "', '" & Request.Form("LeagueID") & "', 0, 0)"
		oConn.Execute(strSQL)		
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/LeagueAssignAdmin.asp"
end if
'-----------------------------------------------
' LeagueRemoveAdmin
'-----------------------------------------------
if Request.Form("SaveType") = "LeagueRemoveAdmin" Then
	if not(bSysAdmin) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	If Len(Request.Form("selLeagueAdminID")) > 0 Then
		strSQL = "DELETE FROM lnk_league_admin WHERE LeagueAdminID = '" & Request.Form("selLeagueAdminID") & "'"
		oConn.Execute(strSQL)		
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/LeagueAssignAdmin.asp"
end if
'-----------------------------------------------
' LeagueDeleteMatch
'-----------------------------------------------
if Request.QueryString("SaveType") = "LeagueDeleteMatch" then
	intMatchID = Request.QueryString("MatchID")
	strSQL = "SELECT LeagueID FROM tbl_league_matches WHERE LeagueMatchID = '" & CheckString(intMatchID) & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intLeagueID = oRS.Fields("LeagueID").Value
	End If
	oRs.Close
	if not(bSysAdmin OR IsLeagueAdminByID(intLeagueID)) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL = "DELETE FROM tbl_league_comms WHERE LeagueMatchID = '" & CheckString(intMatchID) & "';DELETE FROM tbl_league_matches WHERE LeagueMatchID = '" & CheckString(intMatchID) & "'"
	oConn.Execute(strSQL)		
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/TeamLeagueAdmin.asp?league=" & Server.URLEncode(Request.QueryString("league")) & "&team=" & Server.URLEncode(Request.QueryString("team"))
end if
'-----------------------------------------------
' LeagueDeleteMatch
'-----------------------------------------------
if Request.QueryString("SaveType") = "LeagueDeleteMatchAdmin" then
	intMatchID = Request.QueryString("MatchID")
	strSQL = "SELECT LeagueID FROM tbl_league_matches WHERE LeagueMatchID = '" & CheckString(intMatchID) & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intLeagueID = oRS.FieldS("LeagueID").Value
	End If
	oRs.Close
	if not(bSysAdmin OR IsLeagueAdminByID(intLeagueID)) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL = "DELETE FROM tbl_league_comms WHERE LeagueMatchID = '" & CheckString(intMatchID) & "';DELETE FROM tbl_league_matches WHERE LeagueMatchID = '" & CheckString(intMatchID) & "'"
	oConn.Execute(strSQL)		
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/leaguematches.asp?league=" & Server.URLEncode(Request.QueryString("league")) & "&matchdate=" & Server.URLEncode(Request.QueryString("matchdate"))
end if
'-----------------------------------------------
' LeagueReportMatch
'-----------------------------------------------
If Request.Form("SaveType") = "LeagueReportMatch" Then
	intMatchID = Request.Form("MatchID")
	strFromURL = Trim(Request.Form("FromURL"))
	If Len(strFromURL) = 0 Then
		strFromURL = "default.asp"
	End If

	If Not(IsNumeric(intMatchID)) Then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing	
		Response.Clear
		Response.Redirect "errorpage.asp?error=7"
	End if
	
	Dim strMaps(6)
	Dim strHScores(6), strVScores(6)
	Dim intHScore, intVScore
	intHScore = 0
	intVScore = 0
	strSQL = "SELECT HomeTeamLinkID, VisitorTeamLinkID, MatchDate, "
	strSQL = strSQL & " Map1, Map2, Map3, Map4, Map5, "
	strSQL = strSQL & " LeagueID, LeagueConferenceID, LeagueDivisionID "
	strSQL = strSQL & " FROM tbl_league_matches "
	strSQL = strSQL & " WHERE LeagueMatchID = '" & intMatchID & "'"
	'Response.Write strSQL & "<br />"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intHomeLinkID = oRs.Fields("HomeTeamLinkID").Value
		intVisitorLinkID = oRs.Fields("VisitorTeamLinkID").Value
		strMatchDate = oRs.Fields("MatchDate").Value
		strMaps(1) = oRs.Fields("Map1").Value
		strMaps(2) = oRs.Fields("Map2").Value
		strMaps(3) = oRs.Fields("Map3").Value
		strMaps(4) = oRs.Fields("Map4").Value
		strMaps(5) = oRs.Fields("Map5").Value
		intLeagueID = oRs.Fields("LeagueID").Value
		intConferenceID = oRs.Fields("LeagueConferenceID").Value
		intDivisionID = oRs.Fields("LeagueDivisionID").Value
	Else
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect strFromURL
	End if
	oRS.NextRecordSet
	intMaps = 5
	if Len(strMaps(2)) = 0 Then
		intMaps = 1
	Elseif Len(strMaps(3)) = 0 Then
		intMaps = 2
	Elseif Len(strMaps(4)) = 0 Then
		intMaps = 3
	Elseif Len(strMaps(5)) = 0 Then
		intMaps = 4
	Else
		intMaps = 5
	End If
	
	if not(bSysAdmin OR IsLeagueTeamCaptainByLinkID(intHomeLinkId) OR IsLeagueTeamCaptainByLinkID(intVisitorLinkId) OR IsLeagueAdminByID(intLeagueID)) Then
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		response.clear
		response.redirect "errorpage.asp?error=3"
	End If
	
	strSQL = "SELECT LeagueName, l.Scoring, l.WinPoints, l.LossPoints, l.DrawPoints, l.NoShowPoints, l.MapWinPoints, l.MapLossPoints, l.MapDrawPoints, l.MapNoShowPoints "
	strSQL = strSQL & " FROM tbl_leagues l "
	strSQL = strSQL & " WHERE l.LeagueID = '" & intLeagueID & "'"		
	oRS.Open strSQL, oCOnn
	If Not(oRS.EOF AND oRs.BOF) Then
		intWinPoints = oRS.FieldS("WinPoints").Value
		intLossPoints = oRS.FieldS("LossPoints").Value
		intDrawPoints = oRS.FieldS("DrawPoints").Value
		intNoShowPoints = oRS.FieldS("NoShowPoints").Value
		intMapWinPoints = oRS.FieldS("MapWinPoints").Value
		intMapLossPoints = oRS.FieldS("MapLossPoints").Value
		intMapDrawPoints = oRS.FieldS("MapDrawPoints").Value
		intMapNoShowPoints = oRS.FieldS("MapNoShowPoints").Value
		intScoring = oRS.FieldS("Scoring").Value
	End If
	oRS.NextRecordSet
	intHWins = 0
	intVWins = 0
	intHLosses = 0
	intVLosses = 0 
	intHDraws = 0
	intVDraws = 0 
	intHNoShows = 0
	intVNoShows = 0
	intHMapWins = 0
	intVMapWins = 0
	intHMapLosses = 0 
	intVMapLosses = 0 
	intHMapDraws = 0
	intVMapDraws = 0
	intHMapNoShows = 0
	intVMapNoShows = 0
	If intScoring = 0 Then 
		If Request.Form("HNoShow") = "1" Then
			If Request.Form("VNoShow") = "1" Then
				intHNoShows = 1
				intVNoShows = 1
				intHMapNoShows = intMaps
				intVMapNoShows = intMaps
				intHMatchPoints = intNoShowPoints + intHMapNoShows * intMapNoShowPoints
				intVMatchPoints = intNoShowPoints + intVMapNoShows * intMapNoShowPoints
				strOutcome = "Both teams forfeit for failure to show."
			Else
				intHNoShows = 1
				intVWins = 1
				intHMapNoShows = intMaps
				intVMapWins = intMaps
				intHMatchPoints = intNoShowPoints + intHMapNoShows * intMapNoShowPoints
				intVMatchPoints = intWinPoints + intVMapWins * intMapWinPoints
				strOutcome = strHomeName & " forfeits for failure to show."
			End If
		Else
			If Request.Form("VNoShow") = "1" Then
				strOutcome = strVisitorName  & " forfeits for failure to show."
				intHWins = 1
				intVNoShows = 1
				intHMapWins = intMaps
				intVMapNoShows = intMaps
				intHMatchPoints = intWinPoints + intHMapWins * intMapWinPoints
				intVMatchPoints = intNoShowPoints + intVMapNoShows * intMapNoShowPoints
			Else
				For i = 1 to 5
				   	If Len(Trim(strMaps(i))) > 0 AND Not(IsNull(strMaps(i))) Then
						strHScores(i) = Request.Form("HMapScore" & i)
						strVScores(i) = Request.Form("VMapScore" & i)
						If IsNumeric(strHScores(i)) Then
							intHScore = intHScore + strHScores(i)
						End If
						If IsNumeric(strVScores(i)) Then
							intVScore = intVScore + strVScores(i)
						End If
						If IsNumeric(strVScores(i)) AND IsNumeric(strHScores(i)) Then
							if strVScores(i) > strHScores(i) Then
								intVMapWins = intVMapWins + 1
								intHMapLosses = intHMapLosses  + 1
							ElseIf strVScores(i) < strHScores(i) Then
								intHMapWins = intHMapWins + 1
								intVMapLosses = intVMapLosses  + 1
							Else
								intHMapDraws = intHMapDraws + 1
								intVMapDraws = intVMapDraws + 1
							End If
						End If
					End If
				Next	
				intVMapPoints = intVMapDraws * intMapDrawPoints + intVMapLosses * intMapLossPoints + intVMapWins * intMapWinPoints
				intHMapPoints = intHMapDraws * intMapDrawPoints + intHMapLosses * intMapLossPoints + intHMapWins * intMapWinPoints
				If intHScore > intVScore Then
					intHWins = 1
					intVLosses = 1
					intHMatchPoints = intWinPoints + intHMapPoints
					intVMatchPoints = intLossPoints + intVMapPoints 
					strOutcome = strHomeName & " wins. "
				ElseIf intHScore = intVScore Then
					intHDraws = 1
					intVDraws = 1
					intHMatchPoints = intDrawPoints + intHMapPoints
					intVMatchPoints = intDrawPoints + intVMapPoints
					strOutcome = "Tie game."
				Else
					intVWins = 1
					intHLosses = 1
					intHMatchPoints = intLossPoints + intHMapPoints
					intVMatchPoints = intWinPoints + intVMapPoints
					strOutCome = strVisitorName & " wins."
				End If
			End If
		End If
	ElseIf intScoring = 1 Then
		' Each map is different
		intHMapWins = 0 
		intVMapWins = 0
		intVMapLosses = 0
		intHMapLosses = 0
		intHMapDraws = 0
		intVMapDraws = 0
		intHMapNoShows = 0 
		intVMapNoShows = 0
		For i = 1 to 5
			If Len(Trim(strMaps(i))) > 0 AND Not(IsNull(strMaps(i))) Then
				strHScores(i) = Request.Form("HMapScore" & i)
				strVScores(i) = Request.Form("VMapScore" & i)
				If IsNumeric(strHScores(i)) Then
					strHScores(i) = cint(strHScores(i))
				End If
				If IsNumeric(strVScores(i)) Then
					strVScores(i) = Cint(strVScores(i))
				End If
				
				If Request.Form("HMap" & i & "NoShow") = "1" AND Request.Form("VMap" & i & "NoShow") = "1" Then
					intHMapNoShows = intHMapNoShows + 1
					intVMapNoShows = intVMapNoShows + 1
					strOutCome = strOutCome & "Both teams forfeit " & strMaps(i) & "<br />"
				ElseIf Request.Form("HMap" & i & "NoShow") = "1" Then
					intHMapNoShows = intHMapNoShows + 1
					intVMapWins = intVMapWins + 1
					strOutCome = strOutCome & strHomeName & " forfeits " & strMaps(i) & "<br />"
				ElseIf Request.Form("VMap" & i & "NoShow") = "1" Then
					intHMapWins = intHMapWins + 1
					intVMapNoShows = intVMapNoShows + 1
					strOutCome = strOutCome & strVisitorName & " forfeits " & strMaps(i) & "<br />"
				ElseIf strVScores(i) > strHScores(i) Then
					intHMapLosses = intHMapLosses + 1
					intVMapWins = intVMapWins + 1
					strOutCome = strOutCome & strVisitorName & " wins " & strMaps(i) & "<br />"
				ElseIf strVScores(i) < strHScores(i) Then
					intHMapWins = intHMapWins + 1
					intVMapLosses = intVMapLosses + 1
					strOutCome = strOutCome & strHomeName & " wins " & strMaps(i) & "<br />"
				ElseIf strVScores(i) = strHScores(i) Then
					intHMapDraws = intHMapDraws + 1
					intVMapDraws = intVMapDraws + 1
					strOutCome = strOutCome & "Draw on " & strMaps(i) & "<br />"
				End If
			End If
		Next			
		if intVMapLosses = 0 AND intVMapDraws = 0 AND intVMapWins = 0 Then
			If intHMapLosses = 0 AND intHMapDraws = 0 AND intHMapWins = 0 Then
				intHNoShows = 1
			Else
				intHWins = 1
			End If
			' Visitor Forfeit?
			intVNoShows = 1
		ElseIf intHMapLosses = 0 AND intHMapDraws = 0 AND intHMapWins = 0 Then
			intHNoShows = 1
			intVWins = 1
		ElseIf intVMapWins > intHMapWins Then
			intVWins = 1
			intHLosses = 1
		ElseIf intHMapWins > intVMapWins THen
			intHWins = 1
			intVLosses = 1
		Else
			intHDraws = 1
			intVDraws = 1
		End If
		intHScore = intHMapWins
		intVScore = intVMapWins
		
		intHMatchPoints = intHMapWins * intMapWinPoints + intHMapLosses * intMapLossPoints + intHMapDraws * intMapDrawPoints + intHMapNoShows * intMapNoShowPoints
		intVMatchPoints = intVMapWins * intMapWinPoints + intVMapLosses * intMapLossPoints + intVMapDraws * intMapDrawPoints + intVMapNoShows * intMapNoShowPoints
	End If
		
	strSQL = "INSERT INTO tbl_league_history ("
	strSQL = strSQL & " LeagueMatchID, LeagueID, LeagueConferenceID, LeagueDivisionID, "
	strSQL = strSQL & " HomeTeamLinkID, VisitorTeamLinkID, MatchDate, "
	strSQL = strSQL & " Map1, Map1HomeScore, Map1VisitorScore, "
	strSQL = strSQL & " Map2, Map2HomeScore, Map2VisitorScore, "
	strSQL = strSQL & " Map3, Map3HomeScore, Map3VisitorScore, "
	strSQL = strSQL & " Map4, Map4HomeScore, Map4VisitorScore, "
	strSQL = strSQL & " Map5, Map5HomeScore, Map5VisitorScore, "
	strSQL = strSQL & " ReportingPlayerID, ReportTime, HomeTeamPoints, VisitorTeamPoints "
	strSQL = strSQL & " ) VALUES ( "
	strSQL = strSQL & "'" & CheckString(intMatchID) & "', " & intLeagueID & ", " & intConferenceID & ", " & intDivisionID & ", "
	strSQL = strSQL & intHomeLinkID & ", " & intVisitorLinkID & ", '" & strMatchDate & "', "
	For i = 1 to 5
		strSQL = strSQL & "'" & CheckString(strMaps(i)) & "', '" & CheckString(strHScores(i)) & "', '" & CheckString(strVScores(i)) & "', "
	Next
	strSQL = strSQL & "'" & Session("PlayerID") & "', GetDate(), " & intHMatchPoints & "," & intVMatchPoints
	strSQL = strSQL & ")"
	Response.Write "<br /><br />" & strSQL & "<br /><br />"
	oConn.Execute(strSQL)
	
	strSQL = "UPDATE lnk_league_team SET "
	strSQL = strSQL & " LeaguePoints = LeaguePoints + " & intHMatchPoints & ", "
	strSQL = strSQL & " Wins = Wins + " & intHWins & ", "
	strSQL = strSQL & " Losses = Losses + " & intHLosses & ", "
	strSQL = strSQL & " Draws = Draws + " & intHDraws & ", "
	strSQL = strSQL & " NoShows = NoShows + " & intHNoShows & ", "
	strSQL = strSQL & " RoundsWon = RoundsWon + " & intHScore & ", "
	strSQL = strSQL & " RoundsLost = RoundsLost + " & intVScore & " "
	strSQL = strSQL & " WHERE lnkLeagueTeamID = " & intHomeLinkID	
	Response.Write "<br /><br /><b>Home</b> " & strSQL & "<br /><br />"
	oConn.Execute(strSQL)
		
	strSQL = "UPDATE lnk_league_team SET "
	strSQL = strSQL & " LeaguePoints = LeaguePoints + " & intVMatchPoints & ", "
	strSQL = strSQL & " Wins = Wins + " & intVWins & ", "
	strSQL = strSQL & " Losses = Losses + " & intVLosses & ", "
	strSQL = strSQL & " Draws = Draws + " & intVDraws & ", "
	strSQL = strSQL & " NoShows = NoShows + " & intVNoShows & ", "
	strSQL = strSQL & " RoundsWon = RoundsWon + " & intVScore & ", "
	strSQL = strSQL & " RoundsLost = RoundsLost + " & intHScore & " "
	strSQL = strSQL & " WHERE lnkLeagueTeamID = " & intVisitorLinkID	
	Response.Write "<br /><br /><b>Visitor</b>" & strSQL & "<br /><br />"
	oConn.Execute(strSQL)
	
	strSQL = "UPDATE lnk_league_team "
	strSQL = strSQL & " SET WinPct = CASE  "
	strSQL = strSQL & "	WHEN (Wins + Losses + Draws + NoShows) = 0 THEN 0 "
	strSQL = strSQL & "	ELSE Wins * 10000.0 / (Wins + Losses + Draws + NoShows)  "
	strSQL = strSQL & "	END  "
	strSQL = strSQL & " WHERE lnkLeagueTeamID = '" & intHomeLinkID & "'"
	Response.Write "<br /><br />" & strSQL & "<br /><br />"
	 oConn.Execute(strSQL)

	strSQL = "UPDATE lnk_league_team "
	strSQL = strSQL & " SET WinPct = CASE  "
	strSQL = strSQL & "	WHEN (Wins + Losses + Draws + NoShows) = 0 THEN 0 "
	strSQL = strSQL & "	ELSE Wins * 10000.0 / (Wins + Losses + Draws + NoShows) "
	strSQL = strSQL & "	END  "
	strSQL = strSQL & " WHERE lnkLeagueTeamID = '" & intVisitorLinkID & "'"
	Response.Write "<br /><br />" & strSQL & "<br /><br />"
	oConn.Execute(strSQL)

	strSQL = "DELETE FROM tbl_league_matches WHERE LeagueMatchID = '" & CheckString(intMatchID) & "'"
	Response.Write "<br /><br />" & strSQL & "<br /><br />"
	oConn.Execute(strSQL)

'Response.End 
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect Request.Form("FromURL")
End If

'-----------------------------------------------
' Change Record
'-----------------------------------------------
If Request.Form("SaveType") = "ChangeRecord" Then
	intTLLinkID = Request.Form("TLLinkID")
	intLadderID = Request.Form("LadderID")
	If Not(IsSysAdmin()) Then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect = "/errorpage.asp?error=3"
	End If
	intLosses = "0" & Request.Form("Losses")
	intWins = "0" & Request.Form("Wins")
	intForfeits = "0" & Request.Form("Forfeits")
	if isNumeric(intLosses) AND isNumeric(intWins) AND isNumeric(intForfeits) Then
		strSQL = "UPDATE lnk_t_l SET Wins = '" & intWins & "', Losses='"  & intLosses & "', Forfeits='" & intForfeits & "' WHERE TLLinkID = '" & intTLLinkID & "'"
		oConn.Execute(strSQL)
	End If
	%>
	<script language="javascript" type="text/javascript">
	<!--
		window.opener.location = window.opener.location.href;
		window.close();
	//-->
	</script>
	<%
	Response.End
End If 

'-----------------------------------------------
' Change League Record
'-----------------------------------------------
If Request.Form("SaveType") = "ChangeLeagueRecord" Then
	intLnkLeagueTeamID = Request.Form("lnkLeagueTeamID")
	If Not(IsSysAdmin()) Then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect = "/errorpage.asp?error=3"
	End If
	intLosses = Request.Form("Losses")
	intWins = Request.Form("Wins")
	intNoShows = Request.Form("NoShows")
	intDraws = Request.Form("Draws")
	intRoundsWon = Request.Form("RoundsWon")
	intRoundsLost = Request.Form("RoundsLost")
	intLeaguePoints = Request.Form("LeaguePoints")
	intMatchesPlayed = Request.Form("MatchesPlayed")

	if isNumeric(intLosses) AND isNumeric(intWins) AND isNumeric(intNoShows)AND isNumeric(intDraws) Then
		If isNumeric(intRoundsWon) AND isNumeric(intRoundsLost) AND isNumeric(intLeaguePoints) AND isNumeric(intMatchesPlayed) Then
			strSQL = "UPDATE lnk_league_team SET Wins = '" & intWins & "', Losses='"  & intLosses & "',"
			strSQL = strSQL & " NoShows='" & intNoShows & "', "
			strSQL = strSQL & " Draws='" & intDraws & "', "
			strSQL = strSQL & " RoundsWon='" & intRoundsWon & "', "
			strSQL = strSQL & " RoundsLost='" & intRoundsLost & "', "
			strSQL = strSQL & " LeaguePoints='" & intLeaguePoints & "', "
			strSQL = strSQL & " MatchesPlayed='" & intMatchesPlayed & "' "
			strSQL = strSQL & " WHERE lnkLeagueTeamID = '" & intLnkLeagueTeamID & "'"
			oConn.Execute(strSQL)

			strSQL = "UPDATE lnk_league_team "
			strSQL = strSQL & " SET WinPct = CASE  "
			strSQL = strSQL & "	WHEN (Wins + Losses + Draws + NoShows) = 0 THEN 0 "
			strSQL = strSQL & "	ELSE Wins * 10000.0 / (Wins + Losses + Draws + NoShows)  "
			strSQL = strSQL & "	END  "
			strSQL = strSQL & " WHERE lnkLeagueTeamID = '" & intLnkLeagueTeamID & "'"
			oConn.Execute(strSQL)
'			Response.Write "<br /><br />" & strSQL & "<br /><br />"
		End If
	End If
	%>
	<script language="javascript" type="text/javascript">
	<!--
		window.opener.location = window.opener.location.href;
		window.close();
	//-->
	</script>
	<%
	Response.End
End If 
'-----------------------------------------------
' LeagueRemoveAdmin
'-----------------------------------------------
if Request.QueryString("SaveType") = "LeagueBumpBack" then
	if not(bSysAdmin Or IsLeagueAdminByID(Request("LeagueID"))) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL = "UPDATE lnk_league_team SET LeagueDivisionID = 0 WHERE lnkLeagueTeamID = '" & Request.Querystring("lnkLeagueTeamID") & "'"
	oConn.Execute(strSQL)		
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/viewleaguedivision.asp?league=" & Server.URLEncode(Request.Querystring("League") & "") & "&conference=" & Server.URLEncode(Request.Querystring("Conference") & "") & "&division=" & Server.URLEncode(Request.Querystring("Division") & "")
end if
'-----------------------------------------------
' League Match Vote
'-----------------------------------------------
if Request.QueryString("SaveType") = "LeagueMatchVote" then
	if not(LoggedIn) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=2"
	end if
	intMatchID = Request.QueryString("MatchID")
	intVoteID = Request.QueryString("VoteFor")
	strLeagueName = Request.QuerySTring("League")
	strSQL = "SELECT lnkLeagueTeamID FROM lnk_match_player_votes WHERE PlayerID = '" & Session("PlayerID") & "' AND LeagueMatchID ='" & CheckString(intLeagueMatchID) & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect = "viewleaguematch.asp?league=" & Server.URLEncode(strLeagueName & "") & "&LeagueMatchID=" & intMatchID
	Else 
		strSQL = "EXECUTE LeagueMatchVoteFor @MatchID = '" & CheckString(intMatchID) & "', @PlayerID = '" & Session("PlayerID") & "', @VoteFor = '" & intVoteID & "'"
		oConn.Execute (strSQL)
	End If

	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/viewleaguematch.asp?league=" & Server.URLEncode(Request.Querystring("League") & "") & "&LeagueMatchID=" & intMatchID
end if
'-----------------------------------------------
' Ladder Match Vote
'-----------------------------------------------
if Request.QueryString("SaveType") = "LadderMatchVote" then
	if not(LoggedIn) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=2"
	end if
	intMatchID = Request.QueryString("MatchID")
	intVoteID = Request.QueryString("VoteFor")
	strLadderName = Request.QuerySTring("Ladder")
	strSQL = "SELECT TLLinkID FROM lnk_l_p_m_votes WHERE PlayerID = '" & Session("PlayerID") & "' AND MatchID ='" & CheckString(intMatchID) & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		Response.Redirect "viewmatch.asp?ladder=" & Server.URLEncode(strLadderName & "") & "&MatchID=" & intMatchID
	Else 
		strSQL = "EXECUTE LadderMatchVoteFor @MatchID = '" & CheckString(intMatchID) & "', @PlayerID = '" & Session("PlayerID") & "', @VoteFor = '" & intVoteID & "'"
		oConn.Execute (strSQL)
	End If

	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/viewmatch.asp?ladder=" & Server.URLEncode(Request.Querystring("ladder") & "") & "&MatchID=" & intMatchID
end if
'-----------------------------------------------
' Add Rant
'-----------------------------------------------
if Request.form("SaveType") = "Add_Rant" then
	If Not(HasForumAccess()) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=34"	
	End If
	if Session("LoggedIn") AND Len(Trim(Request.Form("Rant"))) > 0 Then
		strSQL="INSERT INTO tbl_league_match_rants ( LeagueMatchID, PlayerID, RantTime, Rant) VALUES (" 
		strSQL = strSQL & "'" & Request.form("LeagueMatchID") & "','" & Session("PlayerID") & "', GetDate(),'" & Replace(RantEncode(Request.Form("Rant")), "'", "''") & "')"
		oconn.Execute(strSQL)
		strSQL = "UPDATE tbl_league_matches SET LastRantTime = GetDate(), LastRanterName = '" & CHeckString(Session("uName")) & "', Rants = Rants + 1 WHERE LeagueMatchID = '" & Request.Form("LeagueMatchID") & "'"
		
		oconn.Execute(strSQL)
	End If
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "ViewLeagueMatch.asp?League=" & server.urlencode(request("League")) & "&LeagueMatchID=" & server.urlencode(request("LeagueMatchID"))
end if
'-----------------------------------------------
' Edit Rant
'-----------------------------------------------
if Request.form("SaveType") = "Edit_Rant" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL= "update tbl_league_match_rants set Rant='" & replace(RantEncode(Request.Form("Rant")), "'", "''") &  "' where LMRID=" & Request.Form("LMRID")
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "ViewLeagueMatch.asp?League=" & server.urlencode(request("League")) & "&LeagueMatchID=" & server.urlencode(request("LeagueMatchID"))
end if
'-----------------------------------------------
' Delete Rant
'-----------------------------------------------
if Request.QueryString("SaveType") = "Delete_Rant" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	if Request.QueryString("LMRID") <> "" then
		strSQL= "delete from tbl_league_match_rants where LMRID=" & Request.QueryString("LMRID")
		'Response.Write strSQl
		oconn.execute(strSQL)
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "ViewLeagueMatch.asp?League=" & server.urlencode(request("League")) & "&LeagueMatchID=" & server.urlencode(request("LeagueMatchID"))
end if
'-----------------------------------------------
' Add Match Rant
'-----------------------------------------------
if Request.form("SaveType") = "Add_Match_Rant" then
	blnTime = false
	If IsNull(Session("lastRant")) OR Len(session("LastRant")) = 0 THen
		blnTime = true
	ElseIf Abs(DateDiff("s", Now(), Session("lastRant"))) > 15 Then
		blnTime = true
	End If
	if Session("LoggedIn") AND Len(Trim(Request.Form("Rant"))) > 0 Then
		If blnTime Then 
			strSQL="INSERT INTO tbl_match_rants ( MatchID, PlayerID, RantTime, Rant) VALUES (" 
			strSQL = strSQL & "'" & Request.form("MatchID") & "','" & Session("PlayerID") & "', GetDate(),'" & Replace(RantEncode(Request.Form("Rant")), "'", "''") & "')"
			oconn.Execute(strSQL)
			strSQL = "UPDATE tbl_matches SET LastRantTime = GetDate(), LastRanterName = '" & CHeckString(Session("uName")) & "', Rants = Rants + 1 WHERE MatchID = '" & Request.Form("MatchID") & "'"
			
			oconn.Execute(strSQL)
			Session("LastRant") = Now()
		Else
			Response.Clear
			Response.Redirect "ViewMatch.asp?Ladder=" & server.urlencode(request("Ladder")) & "&MatchID=" & server.urlencode(request("MatchID")) & "&e=1"
		End If
	End If
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "ViewMatch.asp?Ladder=" & server.urlencode(request("Ladder")) & "&MatchID=" & server.urlencode(request("MatchID"))
end if
'-----------------------------------------------
' Edit Match Rant
'-----------------------------------------------
if Request.form("SaveType") = "Edit_Match_Rant" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	strSQL= "update tbl_match_rants set Rant='" & replace(RantEncode(Request.Form("Rant")), "'", "''") &  "' where MRID=" & Request.Form("MRID")
	ors.open strSQL, oconn
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "ViewMatch.asp?Ladder=" & server.urlencode(request("Ladder")) & "&MatchID=" & server.urlencode(request("MatchID"))
end if
'-----------------------------------------------
' Delete Match Rant
'-----------------------------------------------
if Request.QueryString("SaveType") = "Delete_Match_Rant" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	if Request.QueryString("MRID") <> "" then
		strSQL= "EXECUTE RantsDeleteRant @MRID=" & Request.QueryString("MRID")
		'Response.Write strSQl
		oconn.execute(strSQL)
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "ViewMatch.asp?Ladder=" & server.urlencode(request("Ladder")) & "&MatchID=" & server.urlencode(request("MatchID"))
end if
'-----------------------------------------------
' Purge Match Rant
'-----------------------------------------------
if Request.QueryString("SaveType") = "Purge_Match_Rant" then
	if not(IsSysAdmin()) then
		oConn.Close 
		Set oConn = Nothing
		Set oRs = Nothing
		Response.Clear
		response.redirect "errorpage.asp?error=3"
	end if
	if Request.QueryString("P") <> "" then
		strSQL= "EXECUTE RantsPurgeRants @PlayerID='" & Request.QueryString("p") & "', @MatchID='" & Request.QueryString("MatchID") & "'"
		'Response.Write strSQl
		oconn.execute(strSQL)
	end if
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "ViewMatch.asp?Ladder=" & server.urlencode(request("Ladder")) & "&MatchID=" & server.urlencode(request("MatchID"))
end if

'-----------------------------------------------
' COPPA
'-----------------------------------------------
if Request.QueryString("SaveType") = "Coppa" then
	strSQL = "SELECT PlayerCoppa FROM tbl_players WHERE PlayerID = '" & Session("PlayerID") & "'"
	oRs.Open strSQl, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intPlayerCoppa = oRs.Fields("PlayerCoppa").Value
	End If
	oRs.NextRecordSet
	
	If IsNull(intPlayerCoppa) Then
		strSQL = "UPDATE tbl_players SET PlayerCoppa = '" & CheckString(Request.QuerySTring("Age")) & "' WHERE PlayerID = '" & Session("PlayerID") & "'"
		oConn.Execute(strSQL)
	End If

	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "CoppaThanks.asp"
end if

'-----------------------------------------------
' COPPA Change
'-----------------------------------------------
if Request.Form("SaveType") = "CoppaChange" then
	strSQL = "UPDATE tbl_players SET PlayerCoppa = '1', PlayerCoppaInfo = '" & CheckString(Request.Form("txtname") & " - " & Request.Form("txtDOB")) & " ' WHERE PlayerID = '" & Session("PlayerID") & "'"
	oConn.Execute(strSQL)

	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "CoppaThanks.asp"
end if

'-----------------------------------------------
' Match Dispute
'-----------------------------------------------
if Request.Form("SaveType") = "MatchDispute" then
	
	strPlayerHandle = Request.Form("hdnSubmittor")
	strDisputeTeam = Request.Form("hdnDisputingTeam")
	strDisputedTeam = Request.Form("hdnDisputedTeam")
	strLadderAbbr = Request.Form("hdnLadderAbbr")
	strLadderName = Request.Form("hdnLadderName")
	strLeagueName = Request.Form("hdnLeagueName")
	intLeagueID = Request.Form ("hdnLeagueID")
	intLadderID = Request.Form("hdnLadderID")
	strDisputeReason = Request.Form("DisputeReason")
	strDetails = Request.Form("Details")
	strCompetitionType = Request.Form("hdnCompetitionType")

	If strCompetitionType = "League" then
		strSubject = strLeagueName & ": " & strDisputeTeam & " vs " & strDisputedTeam & " - " & strDisputeReason
		strBody = strPlayerHandle & " has submitted a dispute on the " & strLeaguename & " League. " & vbCrLf & vbCrLf	
	Else
		strSubject = strLadderAbbr & ": " & strDisputeTeam & " vs " & strDisputedTeam & " - " & strDisputeReason
		strBody = strPlayerHandle & " has submitted a dispute on the " & strLaddername & " Ladder. " & vbCrLf & vbCrLf
	End If
	strBody = strBody & "Disputing Team: " & strDisputeTeam & vbCrLf
	strBody = strBody & "Disputed Team: " & strDisputedTeam & vbCrLf
	strBody = strBody & "Reason: " & strDisputeReason  & vbCrLf & vbCrLf
	strBody = strBody & "Details: " & vbCrLf & strDetails & vbCrLf & vbCrLf
	strBody = strBody & "Links: " & vBCrLf
	strBody = strBody & "<a href=""http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode(strDisputeTeam) & """>Disputor</a>"  & vbCrLf
	If strCompetitionType = "Scrim" then
		strBody = strBody & "<a href=""http://www.teamwarfare.com/TeamScrimladderAdmin.asp?team=" & Server.URLEncode(strDisputeTeam) & "&ladder=" & Server.URLEncode(strLadderName) & """>Disputor Admin</a>" & vbCrLf
	Else
		If strCompetitionType = "League" then
			strBody = strBody & "<a href=""http://www.teamwarfare.com/TeamLeagueAdmin.asp?team=" & Server.URLEncode(strDisputeTeam) & "&league=" & Server.URLEncode(strLeagueName) & """>Disputor Admin</a>" & vbCrLf
		Else
			strBody = strBody & "<a href=""http://www.teamwarfare.com/TeamladderAdmin.asp?team=" & Server.URLEncode(strDisputeTeam) & "&ladder=" & Server.URLEncode(strLadderName) & """>Disputor Admin</a>" & vbCrLf
		End If
	End If
	strBody = strBody & "<a href=""http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode(strDisputedTeam) & """>Disputed</a>" & vbCrLf
	If strCompetitionType = "Scrim" Then
		strBody = strBody & "<a href=""http://www.teamwarfare.com/TeamScrimladderAdmin.asp?team=" & Server.URLEncode(strDisputedTeam) & "&ladder=" & Server.URLEncode(strLadderName) & """>Disputed Admin</a>" & vbCrLf
		strBody = strBody & "<a href=""http://www.teamwarfare.com/viewladder.asp?ladder=" & Server.URLEncode(strLadderName) & """>Ladder</a>" & vbCrLf
		strBody = strBody & "<a href=""http://www.teamwarfare.com/adminops.asp?rAdmin=Match&ladderid=" & intLadderID & "&ladder=" & Server.URLEncode(strLadderName) & """>Admin Matches</a>" & vbCrLf
	Else
		If strCompetitionType = "League" Then
			strBody = strBody & "<a href=""http://www.teamwarfare.com/TeamLeagueAdmin.asp?team=" & Server.URLEncode(strDisputedTeam) & "&league=" & Server.URLEncode(strLeagueName) & """>Disputed Admin</a>" & vbCrLf
			strBody = strBody & "<a href=""http://www.teamwarfare.com/viewleague.asp?league=" & Server.URLEncode(strLeagueName) & """>League</a>" & vbCrLf
			strBody = strBody & "<a href=""http://www.teamwarfare.com/leagueadmin.asp"">Admin Matches</a>" & vbCrLf
		Else
			strBody = strBody & "<a href=""http://www.teamwarfare.com/TeamladderAdmin.asp?team=" & Server.URLEncode(strDisputedTeam) & "&ladder=" & Server.URLEncode(strLadderName) & """>Disputed Admin</a>" & vbCrLf
			strBody = strBody & "<a href=""http://www.teamwarfare.com/viewladder.asp?ladder=" & Server.URLEncode(strLadderName) & """>Ladder</a>" & vbCrLf
			strBody = strBody & "<a href=""http://www.teamwarfare.com/adminops.asp?rAdmin=Match&ladderid=" & intLadderID & "&ladder=" & Server.URLEncode(strLadderName) & """>Admin Matches</a>" & vbCrLf
		End If
	End If
	strBody = strBody & "<a href=""http://www.teamwarfare.com/forums/forumdisplay.asp?forumid=" & Request.Form("hdnDisputeForumID") & """>Match Dispute Forum</a>" & vbCrLf
	strThreadSubject = strSubject
	strThreadBody = strBody
	strThreadBody = ForumEncode(strThreadBody)
	strSQL = "EXECUTE ForumsNewThread "
	strSQL = strSQL & "'" & Request.Form("hdnDisputeForumID") & "', "
	strSQL = strSQL & "'" & CheckString(strThreadSubject) & "', "
	strSQL = strSQL & "'" & Session("PlayerID") & "', "
	strSQL = strSQL & " 0, "
	strSQL = strSQL & "'" & CheckString(strThreadBody) & "' "
	oConn.Execute(strSQL)
	
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.RemoteHost  = "127.0.0.1"
	Mailer.FromName    = "TWL: Match Dispute"
	mailer.FromAddress = "automailer@web.teamwarfare.com"
	Mailer.Subject     = strSubject
	Mailer.BodyText    = strBody
	strSQL = "SELECT PlayerHandle, PlayerEmail from lnk_l_a Inner join tbl_players on lnk_l_a.playerid = tbl_players.playerid where ladderid = " & intLadderID
	ors.Open strSQL, oConn
	If Not(ors.EOF AND ors.bof) Then
		Do While Not(ors.eof)
			Mailer.AddRecipient ors.fields("PlayerHandle").value, oRS.Fields("PlayerEmail").Value
			ors.movenext
		loop
	end if
	ors.nextrecordset
	On Error Resume Next
	Mailer.sendMail
	On Error Goto 0
	
	Set Mailer = Nothing
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "DisputeMatchConfirm.asp"
end if'-----------------------------------------------
' Remove team from tournament
'-----------------------------------------------
if Request.QuerySTring("SaveType") = "TournamentRemove" then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End if
	intTMLinkID = Request.QueryString("TMLinkID")
	strTournament = Request.QueryString("Tournament")
	
	strSQL = "EXECUTE TournamentRemoveTeam @TMLinkID = '" & intTMLinkID & "'"
	oConn.Execute(strSQL)

	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "tournament/viewteams.asp?Tournament=" & Server.URLEncode(strTournament & "")
end if

'-----------------------------------------------
' Clear borked match
'-----------------------------------------------
if Request.QuerySTring("SaveType") = "ClearStatus" then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End if
	intTLLinkID = Request.QueryString("TLLinkID")
	strTeam = Request.QueryString("Team")
	strLadder = Request.QueryString("Ladder")
	
	strSQL = "UPDATE lnk_t_l SET Status='Available' WHERE TLLinkID = '" & intTLLinkID & "'"
	oConn.Execute(strSQL)

	strSQL = "DELETE FROM tbl_matches WHERE MatchDefenderID = '" & intTLLinkID & "' OR MatchAttackerID='" & intTLLinkID & "'"
	oConn.Execute(strSQL)

	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "teamladderadmin.asp?team=" & Server.URLEncode(strTeam & "") & "&ladder=" & Server.URLEncode(strLadder & "")
end if

'-----------------------------------------------
' Reset Rest Team
'-----------------------------------------------
if Request.QuerySTring("SaveType") = "ResetRest" then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End if
	intTLLinkID = Request.QueryString("TLLinkID")
	strTeam = Request.QueryString("Team")
	strLadder = Request.QueryString("Ladder")
	
	strSQL = "UPDATE lnk_t_l SET RestDays = 0 WHERE TLLinkID = '" & intTLLinkID & "'"
	oConn.Execute(strSQL)

	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "teamladderadmin.asp?team=" & Server.URLEncode(strTeam & "") & "&ladder=" & Server.URLEncode(strLadder & "")
end if

'-----------------------------------------------
' Reset Rest Ladder
'-----------------------------------------------
if Request.QuerySTring("SaveType") = "LadderResetRest" then
	If Not(bSysAdmin) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End if
	intLadderID = Request.QueryString("LadderID")
	strTeam = Request.QueryString("Team")
	strLadder = Request.QueryString("Ladder")
	
	strSQL = "UPDATE lnk_t_l SET RestDays = 0 WHERE LadderID = '" & intLadderID & "'"
	oConn.Execute(strSQL)

	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "adminops.asp?aType=Ladder"
end if

'-----------------------------------------------
' Clear borked match
'-----------------------------------------------
if Request.QuerySTring("SaveType") = "ResetRest" then
	If Not(bSysAdmin OR IsLadderAdmin(request.querystring("ladder"))) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End if
	
	strSQL = "UPDATE lnk_t_l SET RestDays = 0 WHERE LadderID = '" & REquest.QueryString("LadderID")& "'"
	oConn.Execute(strSQL)

	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "adminops.asp?aType=ladder"
end if

if Request.Form("SaveType") = "AntiSmurfAdd" then
	strPlayerName = Trim(Request.Form("hdnPlayer"))
	If Not(bSysAdmin OR Session("uName") = strPlayerName) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End if
	intIdentifierID = Trim(Request.Form("selIdentifierID"))
	strIdentiferValue = Trim(Request.Form("txtIdentifierValue"))
	
	' Validate no dupes exist
	strSQL = "SELECT PlayerID FROM lnk_player_identifier WHERE IdentifierID = '" & CheckString(intIdentifierID) & "' AND IdentifierValue = '" & CheckString(strIdentiferValue) & "' AND IdentifierActive = 1"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intInUseID = oRs.Fields("PlayerID").Value
		oRs.Close
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Response.Clear
		Response.Redirect "IdentifierAdd.asp?Player=" & Server.URLEncode(strPlayerName) & "&Identifier=" & intIdentifierID & "&Relevant=" & Server.URLEncode(strIdentiferValue) & "&InUse=" & intInUseID
	End If
	oRs.NextRecordSet
	
	strSQL = "SELECT PlayerID FROM tbl_players WHERE PlayerHandle= '" & CheckString	(strPlayerName) & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intPlayerID = oRs.FieldS("PlayerID").Value
		strSQL = "INSERT INTO lnk_player_identifier (PlayerID, IdentifierID, IdentifierValue, DateAdded, IdentifierActive) VALUES ("
		strSQL = strSQL & "'" & intPlayerID & "', "
		strSQL = strSQL & "'" & CheckString(intIdentifierID) & "', "
		strSQL = strSQL & "'" & CheckString(strIdentiferValue) & "', GetDate(), 1) "
		oConn.Execute(strSQL)

		oRs.Close
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Response.Clear
		Response.Redirect "ViewPlayer.asp?Player=" & Server.URLEncode(strPlayerName)
	End If
	oRs.NextRecordset
	
End if
		
if Request.QueryString("SaveType") = "AntiSmurfDel" then
	strPlayerName = Trim(Request.QueryString("Player"))
	intIdentifier = Request.QueryString("Identifier")
	strSQL = "SELECT PlayerID FROM lnk_player_identifier WHERE lnkPlayerIdentifierID = '" & CheckString(intIdentifier) & "'"
	oRs.Open strSQL, oConn
	If NoT(oRs.EOF AND oRs.BOF) Then 
		intPlayerID = oRs.FIelds("PlayerID").Value
	End IF
	oRs.NextRecordSet
	
	If Not(bSysAdmin OR Session("PlayerID") = cStr(intPlayerID)) Then
		Response.Clear
		Response.Redirect "/errorpage.asp?error=3"
	End If
	
	strSQL = "UPDATE lnk_player_identifier SET IdentifierActive = 0 WHERE lnkPlayerIdentifierID = '" & CheckString(intIdentifier) & "'"
	oConn.Execute(strSQL)
	
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "ViewPlayer.asp?Player=" & Server.URLEncode(strPlayerName)
ENd If

'-----------------------------------------------
' Change Reported 1v1 Match
'-----------------------------------------------
if request.form("SaveType") = "Change1v1Match" then
        if not (IsSysAdmin() or IsLadderAdminByID(request.form("LadderID"))) then
                oConn.Close
                Set oConn = Nothing
                Set oRs = Nothing
                Response.Clear
                response.redirect "errorpage.asp?error=3"
        end if

        LadderID = Request.Form("LadderID")
        HistoryID = Request.Form("HistoryID")
        DefID = Request.Form("DefenderID")
        AttID = Request.Form("AttackerID")
        DefOldRank = Request.Form("DefOldRank")
        DefNewRank = Request.Form("DefNewRank")
        AttOldRank = Request.Form("AttOldRank")
        AttNewRank = Request.Form("AttNewRank")

        Map1 = Request.Form("Map1Name")
        DefMap1Score = Request.Form("Map1DefScore")
        AttMap1Score = Request.Form("Map1AttScore")
        Map1FT = Request.Form("Map1FT")

        OverAllForfeit = Request.Form("MatchFT")
        DefenderWin = Request.Form("DefWin")
        OldDefenderWin = Request.form("DefOldWin")

        matchdate = Request.Form("MatchDate")

        If DefenderWin then
                winid = defid
                losid = attid
        else
                winid = attid
                losid = defid
        end if

        if defenderwin <> olddefenderwin then
                strSQL = "update lnk_T_L set Losses=(Losses-1), wins=(wins+1) where TLLinkID=" & winid
                'response.write strsql
                ors.open strsql, oconn
                strSQL = "update lnk_T_L set Losses=(Losses+1), wins=(wins-1) where TLLinkID=" & losid
                'response.write strsql
                ors.open strsql, oconn

        end if

        strSQL = "update tbl_PlayerHistory set MatchWinnerID='" & WinID
        strsql = strsql & "', MatchLoserID='" & losID
        strsql = strsql & "', MatchMap1='" & replace(map1, "'", "''")
        strsql = strsql & "', MatchMap1DefenderScore='" & DefMap1Score
        strsql = strsql & "', MatchMap1AttackerScore='" & AttMap1Score
        strsql = strsql & "', MatchMap1Forfeit='" & Map1FT
        strsql = strsql & "', MatchForfeit='" & MatchFT
        strsql = strsql & "', MatchDate='" & replace(MatchDate, "'", "''")
        strsql = strsql & "', MatchWinnerDefending='" & DefenderWin
        strsql = strsql & "' where HistoryID='" & HistoryID & "'"

        ors.open strsql, oConn

        'Response.Write "<font color=white>" & strsql
        if defnewrank <> defoldrank then
                'response.write "<br>Defender getting a new rank."
                if defnewrank < defoldrank then
                        'response.write "Defender moving up in the world"
                        strsql="update lnk_p_L set rank=(rank+1) where (rank > " & defnewrank - 1 & "  and rank < " & defoldrank & ") and isactive=1 and ladderid = " & ladderid
                else
                        'response.write "Defender gets a demotion"
                        strsql="update lnk_p_L set rank=(rank-1) where (rank < " & defnewrank + 1 & "  and rank > " & defoldrank & ") and isactive=1 and ladderid = " & ladderid
                end if
                ors.open strsql, oconn
                'response.write "<br>" & strsql
                strsql = "update lnk_p_l set rank=" & defnewrank & " where PPLLinkID=" & defid
                ors.open strsql, oconn
                'response.write "<br>" & strsql
        end if

        if attnewrank <> attoldrank then
                'response.write "<br>Attacker getting a new rank."
                if attnewrank < attoldrank then
                        'response.write "Attacker moving up in the world"
                        strsql="update lnk_p_L set rank=(rank+1) where (rank > " & attnewrank - 1 & "  and rank < " & attoldrank & ") and isactive=1 and ladderid = " & ladderid
                else
                        'response.write "Attacker gets a demotion"
                        strsql="update lnk_p_L set rank=(rank-1) where (rank < " & attnewrank + 1 & "  and rank > " & attoldrank & ") and isactive=1 and ladderid = " & ladderid
                end if
                ors.open strsql, oconn
                'response.write "<br>" & strsql
                strsql = "update lnk_p_l set rank=" & attnewrank & " where PPLLinkID=" & attid
                ors.open strsql, oconn
                'response.write "<br>" & strsql

        end if

        oConn.Close
        Set oConn = Nothing
        Set oRs = Nothing

        Response.Clear
        Response.Redirect "edit1v1history.asp"

end if



'-----------------------------------------------
' Delete a match from history
'-----------------------------------------------
if Request.Form("SaveType") = "Delete1v1History" Then
        HistoryID = Request.Form("HistoryID")
        If Not(bSysAdmin) Then
                oConn.Close
                Set oConn = Nothing
                Set oRs = Nothing
                Response.Clear
                response.redirect "errorpage.asp?error=3"
        End If
        strSQL = "SELECT MatchWinnerID, MatchLoserID, MatchForfeit FROM tbl_PlayerHistory WHERE HistoryID='" & HistoryID & "'"
        oRS.Open strSQL, oConn
        If Not(oRS.EOF AND oRS.BOF) Then
                WinnerLinkID = oRS.Fields("MatchWinnerID").Value
                LoserLinkID = oRS.Fields("MatchLoserID").Value
                Forfeit = oRS.Fields("MatchForfeit").Value
                oRs.Close
                If cBool(forfeit) Then
                        strSQL = "UPDATE lnk_p_pl SET Forfeits = ForFeits - 1 WHERE PPLLinkID='" & LoserLinkID & "';DELETE FROM tbl_Playerhistory WHERE HistoryID='" & HistoryID & "'"
                        oConn.Execute strSQL
                Else
                        strSQL = "UPDATE lnk_p_pl SET Losses = Losses - 1 WHERE PPLLinkID='" & LoserLinkID & "'"
                        strSQL = strSQL & "UPDATE lnk_p_pl SET Wins = Wins - 1 WHERE PPLLinkID='" & WinnerLinkID & "';"
                        strSQL = strSQL & "DELETE FROM tbl_PlayerHistory WHERE HistoryID='" & HistoryID & "'"
                        oConn.Execute strSQL
                End If
        End If
        Response.Redirect "/edit1v1history.asp?ladder=" & Server.URLEncode(Request.Form("Ladder"))
End If
	
	
	'-----------------------------------------------
' Housekeeping
'-----------------------------------------------
On Error Resume Next
oConn.Close
On Error Goto 0
set ors = nothing
set oConn = nothing	
set ors2 = nothing	
Response.Clear
Response.Redirect "/default.asp"
%>