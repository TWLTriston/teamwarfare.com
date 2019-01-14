<% 'Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Security Request"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin, strGUID
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

Security Request: 
<%

SecType = Request.Form("SecType")
if sectype="" then
	if Request.QueryString("SecType") <> "" then
		SecType=Request.QueryString("SecType")
	else
		%>
		<script>
		window.close();
		window.login.opener.location='/';

		</script>
		<%
	end if
end if
if SecType="login" then
	session("uName")=""
	Dim lastCheck, thisCheck
	thisCheck = Second(Now)
	lastCheck = Session("LastCheck")
	If ( thisCheck = lastCheck ) Then
		response.redirect "login.asp?error=true&t=" & thisCheck
	End If
	Session("LastCheck") = thisCheck
	
	strSQL="select PlayerHandle, PlayerPassword, PlayerID, PlayerActive, StyleID, PlayerGUID from tbl_Players where PlayerHandle='" & CheckString(Request.Form("uName")) & "'"
	oRs.Open strSQL, oConn
	if ors.eof and ors.bof then
		response.redirect "login.asp?error=true&url=/"
	else
		If oRS.Fields("PlayerActive").Value <> "Y" Then
			oRS.Close
			Set oRS = Nothing
			oConn.Close
			Set oCOnn = Nothing
			Response.Clear 
			%>
			<HTML><HEAD></HEAD><BODY>
			<script language="javascript">
			window.opener.location='/activate.asp?error=1';
			window.close();
			</script>
			</BODY></HTML>
			<%
			Response.End 
		End If
		playerid = oRS("PlayerID")
		ppass = Request.Form("uPassword")
		' Redacted
	
			oConn.Close
			set ors=nothing
			set oconn=nothing
			Response.Clear 
			response.redirect "login.asp?error=true&url=/"
		end if
	end if
end if		
if SecType="SpecLogin" then
	If Not(IsSysAdmin()) Then
		Response.Clear 
		%>
		<HTML><HEAD></HEAD><BODY>
		<script language="javascript">
		window.close();
		</script>
		</BODY></HTML>
		<%
		Response.End
	End If
	strSQL="select PlayerHandle, PlayerPassword, PlayerID, PlayerActive from tbl_Players where PlayerHandle='" & CheckString(Request.Form("uName")) & "'"
	oRs.Open strSQL, oConn
	if ors.eof and ors.bof then
		response.redirect "login.asp?error=true&url=/"
	else
		Response.Cookies("User")("UserInfo")=ename
		Response.Cookies("User")("UserInfo2")=strGUID
		Session("LoggedIn") = True
		Session("PlayerID") = oRS.Fields ("PlayerID").Value 
		Session("uName")	= oRS.Fields ("playerHandle").Value 
		Session("SysAdminLevel") = ""
		Session("SysAdmin") = ""
		Session("AnyLadderAdmin") = ""
		ors.Close 
		oConn.Close
		set ors=nothing
		set oconn=nothing
		Response.Clear 
		'-------------------------------------------
		Response.cookies("User")("uName")=ename
		'-------------------------------------------
		Response.Cookies("User").expires = "1/1/2005"
		%>
		<script language="javascript">
		window.opener.location='<%=request.form("fromurl")%>';
		window.close();
		</script>
		<%
	end if
end if		
'--------------------------------------------------
if SecType="teamquit" then
	PlayerID=Request.Form("PlayerID")
	TeamID=Request.Form("TeamID")
	LadderName=Request.Form("LadderToQuit")
	LadderID=Request.Form("Ladderid")
	FromURL=Request.Form("fromurl")
'	strSQL="select ladderid from tbl_ladders where laddername='" & LadderName & "'"
'	
	'oRs.open strSQL, oConn
	'^^^ st00pid fucking error
'	ors.Open strSQL, oConn
'	if not (ors.eof and ors.bof) then
'		LadderID=ors.fields(0).value
'	end if
'	ors.close
	strSQL = "select TLLinkID from lnk_T_L where ((LadderID=" & ladderid & ") and (TeamID=" & teamid & "))"
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		linkID = ors.fields(0).value
	end if
	ors.close
	strSQL="delete from lnk_T_P_L where TLLinkID=" & linkID & " and playerID=" & playerid
	oConn.Execute (strSQL)
	oConn.Close
	set ors=nothing
	set oconn=nothing
	Response.Clear 
%>
	<script>
	window.opener.location='<%=fromurl%>';
	this.window.close();
	</script>
<%
end if
'--------------------------------------------------
if SecType="TeamScrimQuit" then
	PlayerID=Request.Form("PlayerID")
	TeamID=Request.Form("TeamID")
	LadderName=Request.Form("LadderToQuit")
	LadderID=Request.Form("Ladderid")
	FromURL=Request.Form("fromurl")

	strSQL = "select lnkELoTeamID from lnk_elo_team where EloLadderID=" & ladderid & " and TeamID=" & teamid & ""
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		linkID = ors.fields(0).value
	end if
	ors.close
	strSQL="delete from lnk_elo_team_player where lnkELoTeamID=" & linkID & " and playerID=" & playerid
	oConn.Execute (strSQL)
	oConn.Close
	set ors=nothing
	set oconn=nothing
	Response.Clear 
%>
	<script>
	window.opener.location='<%=fromurl%>';
	this.window.close();
	</script>
<%
end if
'--------------------------------------------------
if SecType="TeamLeagueQuit" then
	PlayerID=Request.Form("PlayerID")
	TeamID=Request.Form("TeamID")
	LeagueName=Request.Form("LadderToQuit")
	LeagueID=Request.Form("LeagueID")
	FromURL=Request.Form("fromurl")
	strSQL = "select lnkLeagueTeamID from lnk_league_team where LeagueID='" & LeagueID& "' and TeamID='" & teamid & "'"
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		linkID = ors.fields(0).value
	end if
	ors.close
	strSQL="delete from lnk_league_team_player where lnkLeagueTeamID=" & linkID & " and playerID=" & playerid
	oConn.Execute (strSQL)
	oConn.Close
	set ors=nothing
	set oconn=nothing
	Response.Clear 
%>
	<script>
	window.opener.location='<%=fromurl%>';
	this.window.close();
	</script>
<%
end if
'--------------------------------------------------
if SecType="teamtournamentquit" then
	PlayerID=Request.Form("PlayerID")
	TeamID=Request.Form("TeamID")
	TournamentName=Request.Form("TournamentToQuit")
	TournamentID=Request.Form("TournamentID")
	FromURL=Request.Form("fromurl")
	strSQL = "select TMLinkID from lnk_T_M where ((TournamentID=" & TournamentID & ") and (TeamID=" & teamid & "))"
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		linkID = ors.fields(0).value
	end if
	ors.close
	strSQL="delete from lnk_T_M_P where TMLinkID=" & linkID & " and playerID=" & playerid
	oConn.Execute(strSQL)
	oConn.Close
	set ors=nothing
	set oconn=nothing
	Response.Clear 
%>
	<script>
	window.opener.location='<%=fromurl%>';
	this.window.close();
	</script>
<%
end if

'--------------------------------------------------
if SecType="teamjoin" then
	if not (Session("LoggedIn")) then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		response.clear
		response.redirect "errorpage.asp?error=2"
	end if
	
	Response.Write Request.Form & "<br>"
	if Request.Form("LadderToJoin") = "" then
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Response.Clear
		Response.Redirect "join.asp?keydata=" & server.urlencode(Request.Form("keydata"))  & "&jointype=team&errortype=noladder"
	end if
	jdate=Request.Form("jDate")
	if request.form("teamid") = "" then 
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Response.Clear
		response.redirect "/default.asp"
	end if
	tid = request.form("teamid")
	strSQL="select TeamJoinPassword, TeamName, TeamAdmin from tbl_Teams where TeamID=" & tid
	oRs.Open strSQL, oConn
	if not (ors.eof and ors.bof) then
		pw = Request.Form("password")
		IsAdmin = 0
		if session("uName") = ors.fields(2).value then
			pw = ors.fields(0).value
			IsAdmin = 1
		end if
		ladder = Request.Form("LadderToJoin")
		if ladder="" then
			ladder=session("CurrentLadder")
		end if
		if ors.fields(0).value = pw then
			lID = "1"
			ors.close
			strsql= "select LadderID, IdentifierID from tbl_ladders where laddername='" & replace(ladder, "'", "''") & "'"
			ors.open strsql, oconn
			if (ors.eof and ors.bof) then
				oRs.Close
				oConn.Close
				Set oConn = Nothing
				Set oRS = Nothing
				Response.Clear
				Response.Redirect "/errorpage.asp?error=7"
			end if
			lid=ors.fields(0).value
			if ors.fields(1).value > 0 then
				needGUID = 1
			end if
			ors.close
			strSQL = "select PlayerID from tbl_players where PlayerHandle = '" & replace(session("uName"), "'", "''") & "'" 
			ors.open strsql, oconn
			if (ors.eof and ors.bof) then
				oRs.Close
				oConn.Close
				Set oConn = Nothing
				Set oRS = Nothing
				Response.Clear
				Response.Redirect "/errorpage.asp?error=7"
			end if
			pID = ors.fields(0).value
			ors.close
			strsql="select TLLinkID from lnk_T_L where TeamID = " & tID & " and LadderID = " & lID
			ors.open strsql, oconn
				linkID = ors.fields("TLLinkID").value
			ors.close
		  
		  'GUID check
		  if needGUID = 1 then
		  	strSQL="SELECT IdentifierID FROM lnk_player_identifier WHERE playerid = '" & pid & "' AND IdentifierID = '" & Request.Form("IdentifierID") & "' AND IdentifierActive = 1"
		  	ors.open strSQL, oconn
		  	if (ors.eof and oRs.bof) then
			  	fromurl=Request.Form ("fromurl")
			 		tid=Request.Form ("teamid")
		 	 		lid=Request.Form ("ladderid")
			 		response.redirect "joinTeamOnLadder.asp?teamID=" & tID & "&ladderid=" & lid & "&type=join&error=9&url=" & fromurl
				end if
				ors.close
			end if
			
			strSQL="select count(PlayerID) from lnk_T_P_L inner join lnk_T_L on lnk_T_P_L.TLLinkID=lnk_T_L.TLLinkID where lnk_T_P_L.PlayerID=" & pid & " and lnk_T_L.LadderID=" & lid
			ors.open strSQL, oconn
				lcnt=ors.fields(0).value
			ors.close 
			'Response.Write "<br>PID=" & pid
			'Response.Write "<br>lcnt=" & lcnt
			
			if lcnt = 0 then
				strsql="select count(*) from lnk_T_P_L where PlayerID = " & pID & " and TLLinkID = " & linkID
				ors.open strsql, oconn
				rcnt = ors.fields(0).value
				'Response.Write pID
				'Response.Write tID
				'Response.Write lID
				'response.Write linkID
				'Response.Write rcnt		
				if rcnt = 0 then
					strsql="insert into lnk_T_P_L (PlayerID, TLLinkID, isAdmin, DateJoined) values (" & pID & ", " & linkId & "," & IsAdmin & " ,'" & now & "')"
					'Response.Write strsql
					oConn.Execute(strSQL)
					Response.Write "<br>User add successful"
					fromurl=Request.Form ("fromurl")
					%>
					<script>
					window.opener.location='<%=fromurl%>';
					this.window.close();
					</script>
					<%
				end if
			else
		 	 fromurl=Request.Form ("fromurl")
			 tid=Request.Form ("teamid")
		 	 lid=Request.Form ("ladderid")
			 response.redirect "joinTeamOnLadder.asp?teamID=" & tID & "&ladderid=" & lid & "&type=join&error=2&url=" & fromurl
			end if
		else
		 fromurl=Request.Form ("fromurl")
		 tid=Request.Form ("teamid")
		 lid=Request.Form ("ladderid")
		 response.redirect "joinTeamOnLadder.asp?teamID=" & tID & "&ladderid=" & lid & "&type=join&error=1&url=" & fromurl
		end if
	end if
	fromurl=Request.Form("fromurl")
	oConn.Close
	set ors=nothing
	set oconn=nothing
	Response.Clear 
	if request("join") = 2 then 
		response.redirect "/default.asp"
	else%>
	<script>
	window.opener.location='<%=fromurl%>';
	this.window.close();
	</script>
	<%
	end if
end if

'-------------------
' Join Scrim Ladder
'-------------------
if SecType="TeamScrimJoin" then
	if not (Session("LoggedIn")) then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		response.clear
		response.redirect "errorpage.asp?error=2"
	end if
	
	Response.Write Request.Form & "<br>"
	if Request.Form("LadderToJoin") = "" then
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Response.Clear
		Response.Redirect "join.asp?keydata=" & server.urlencode(Request.Form("keydata"))  & "&jointype=team&errortype=noladder"
	end if

	if request.form("teamid") = "" then 
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Response.Clear
		response.redirect "/errorpage.asp?error=7"
	end if
	tid = request.form("teamid")
	strSQL="select TeamJoinPassword, TeamName, TeamAdmin from tbl_Teams where TeamID=" & tid
	oRs.Open strSQL, oConn
	if not (ors.eof and ors.bof) then
		pw = Request.Form("password")
		IsAdmin = 0
		if session("uName") = ors.fields(2).value then
			pw = ors.fields(0).value
			IsAdmin = 1
		end if
		ladder = Request.Form("LadderToJoin")
		if ladder="" then
			ladder=session("CurrentLadder")
		end if
		if ors.fields(0).value = pw then
			lID = "1"
			ors.close
			strsql= "select EloLadderID from tbl_elo_ladders where Eloladdername='" & replace(ladder, "'", "''") & "'"
			ors.open strsql, oconn
			if (ors.eof and ors.bof) then
				oRs.Close
				oConn.Close
				Set oConn = Nothing
				Set oRS = Nothing
				Response.Clear
				Response.Redirect "/errorpage.asp?error=7"
			end if
			lid=ors.fields(0).value
			ors.close
			strSQL = "select PlayerID from tbl_players where PlayerHandle = '" & replace(session("uName"), "'", "''") & "'" 
			ors.open strsql, oconn
			if (ors.eof and ors.bof) then
				oRs.Close
				oConn.Close
				Set oConn = Nothing
				Set oRS = Nothing
				Response.Clear
				Response.Redirect "/errorpage.asp?error=7"
			end if
			pID = ors.fields(0).value
			ors.close
			strsql="select lnkEloTeamID from lnk_elo_team where TeamID = " & tID & " and EloLadderID = " & lID
			ors.open strsql, oconn
				linkID = ors.fields("lnkEloTeamID").value
			ors.close
		
			strSQL="select count(PlayerID) from lnk_elo_team_player letp inner join lnk_elo_team et ON et.lnkEloTeamID = letp.lnkEloTeamID WHERE PlayerID=" & pid & " and EloLadderID=" & lid
			ors.open strSQL, oconn
				lcnt=ors.fields(0).value
			ors.close 
			'Response.Write "<br>PID=" & pid
			'Response.Write "<br>lcnt=" & lcnt
			
			if lcnt = 0 then
				strsql="select count(*) from lnk_elo_team_player where PlayerID = " & pID & " and lnkEloTeamID = " & linkID
				ors.open strsql, oconn
				rcnt = ors.fields(0).value
				'Response.Write pID
				'Response.Write tID
				'Response.Write lID
				'response.Write linkID
				'Response.Write rcnt		
				if rcnt = 0 then
					strsql="insert into lnk_elo_team_player (PlayerID, lnkEloTeamID, isAdmin, JoinDate) values (" & pID & ", " & linkId & "," & IsAdmin & " ,GetDate())"
					'Response.Write strsql
					oConn.Execute(strSQL)
					'Response.Write "<br>User add successful"
					fromurl=Request.Form ("fromurl")
					%>
					<script>
					window.opener.location=window.opener.location.href;
					window.close();
					</script>
					<%
				end if
			else
		 	 fromurl=Request.Form ("fromurl")
			 tid=Request.Form ("teamid")
		 	 lid=Request.Form ("ladderid")
			 response.redirect "joinTeamOnScrimLadder.asp?teamID=" & tID & "&ladderid=" & lid & "&type=join&error=2&url=" & fromurl
			end if
		else
		 fromurl=Request.Form ("fromurl")
		 tid=Request.Form ("teamid")
		 lid=Request.Form ("ladderid")
		 response.redirect "joinTeamOnScrimLadder.asp?teamID=" & tID & "&ladderid=" & lid & "&type=join&error=1&url=" & fromurl
		end if
	end if
	fromurl=Request.Form("fromurl")
	oConn.Close
	set ors=nothing
	set oconn=nothing
	Response.Clear 
	if request("join") = 2 then 
		response.redirect "/default.asp"
	else%>
	<script>
	window.opener.location=window.opener.location.href;
	window.close();
	</script>
	<%
	end if
end if
'''''''''''''''''''''
' Join League
'''''''''''''''''''''
if SecType="TeamLeagueJoin" then
	if not (Session("LoggedIn")) then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		response.clear
		response.redirect "errorpage.asp?error=2"
	end if
	
	Response.Write Request.Form & "<br>"
	if Request.Form("LeagueToJoin") = "" then
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Response.Clear
		Response.Redirect "join.asp?keydata=" & server.urlencode(Request.Form("keydata"))  & "&jointype=team&errortype=noladder"
	end if
	jdate=Request.Form("jDate")
	if request.form("teamid") = "" then 
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Response.Clear
		response.redirect "/default.asp"
	end if
	tid = request.form("teamid")

	lid=Request.Form("LeagueID")
	
	strSQL="select TeamJoinPassword, TeamName, TeamAdmin from tbl_Teams where TeamID=" & tid
	oRs.Open strSQL, oConn
	if not (ors.eof and ors.bof) then
		pw = Request.Form("password")
		IsAdmin = 0
		if session("uName") = ors.fields(2).value then
			pw = ors.fields(0).value
			IsAdmin = 1
		end if
		if ors.fields(0).value = pw then
			ors.close
			
			strSQL = "select PlayerID from tbl_players where PlayerHandle = '" & replace(session("uName"), "'", "''") & "'" 
			ors.open strsql, oconn
			if (ors.eof and ors.bof) then
				oRs.Close
				oConn.Close
				Set oConn = Nothing
				Set oRS = Nothing
				Response.Clear
				Response.Redirect "/errorpage.asp?error=7"
			end if
			pID = ors.fields(0).value
			ors.close
			
			'GUID check
			if Request.Form ("needGUID") = 1 then
		  	strSQL="SELECT IdentifierID FROM lnk_player_identifier WHERE playerid = '" & pid & "' AND IdentifierID = '" & Request.Form("IdentifierID") & "' AND IdentifierActive = 1"
		  	ors.open strSQL, oconn
		  	if (ors.eof and oRs.bof) then
			  	fromurl=Request.Form ("fromurl")
			 		tid=Request.Form ("teamid")
		 	 		lid=Request.Form ("leagueid")
			 		response.redirect "joinTeamOnLeague.asp?teamID=" & tID & "&leagueid=" & lid & "&type=join&error=9&url=" & fromurl
				end if
				ors.close
			end if
			
			strsql="select lnkLeagueTeamID from lnk_league_team WHERE TeamID = " & tID & " and LeagueID = " & lID 
			ors.open strsql, oconn
			linkID = ors.fields(0).value
			ors.close
			strSQL="select count(PlayerID) from lnk_league_team_player lltp inner join lnk_league_team llt on lltp.lnkLeagueTeamID=llt.lnkLeagueTeamID where llt.Active = 1 AND lltp.PlayerID=" & pid & " and llt.LeagueID=" & lid
			ors.open strSQL, oconn
			lcnt=ors.fields(0).value
			'Response.Write "<br>PID=" & pid
			'Response.Write "<br>lcnt=" & lcnt
			ors.close 
			if lcnt = 0 then
				strsql="select count(PlayerID) from lnk_league_team_player where PlayerID = " & pID & " and lnkLeagueTeamID = " & linkID
				
				ors.open strsql, oconn
				rcnt = ors.fields(0).value
				'Response.Write pID
				'Response.Write tID
				'Response.Write lID
				'response.Write linkID
				'Response.Write rcnt		
				if rcnt = 0 then
					strsql="insert into lnk_league_team_player (PlayerID, lnkLeagueTeamID, isAdmin, JoinDate) values (" & pID & ", " & linkId & "," & IsAdmin & " ,GetDate())"
					'Response.Write strsql
					oConn.Execute(strSQL)
'					Response.Write "<br>User add successful"
					fromurl=Request.Form ("fromurl")
					%>
					<script>
					window.opener.location='<%=fromurl%>';
					this.window.close();
					</script>
					<%
				end if
			else
			 	 fromurl=Request.Form ("fromurl")
				 tid=Request.Form ("teamid")
			 	 lid=Request.Form ("leagueid")
			 response.redirect "joinTeamOnLeague.asp?teamID=" & tID & "&leagueid=" & lid & "&type=join&error=2&url=" & fromurl
			end if
		else
		 fromurl=Request.Form ("fromurl")
		 tid=Request.Form ("teamid")
		 lid=Request.Form ("leagueid")
		 response.redirect "joinTeamOnLeague.asp?teamID=" & tID & "&leagueid=" & lid & "&type=join&error=1&url=" & fromurl
		end if
	end if
	fromurl=Request.Form("fromurl")
	oConn.Close
	set ors=nothing
	set oconn=nothing
	Response.Clear 
	if request("join") = 2 then 
		response.redirect "/default.asp"
	else%>
	<script>
	window.opener.location='<%=fromurl%>';
	this.window.close();
	</script>
	<%
	end if
end if
'--------------------------------------------------
if SecType="teamtournamentjoin" then
	if not (Session("LoggedIn")) then
		oConn.Close
		set ors=nothing
		set oconn=nothing
		Response.Clear 
		response.redirect "errorpage.asp?error=2"
	end if
	
	Response.Write Request.Form & "<br>"
	jdate=Request.Form("jDate")
	'if Request.Form("keydata") = "" then
	'	teamtouse = session("CurrentTeam")
	'else
	'	teamtouse = Request.Form("keydata")
	'end if
	if request.form("teamid") = "" then 
		response.redirect "teamlist.asp"
	end if
	tid = request.form("teamid")
	strSQL="select TeamJoinPassword, TeamName, TeamAdmin from tbl_Teams where TeamID=" & tid
	oRs.Open strSQL, oConn
	if not (ors.eof and ors.bof) then
		Response.Write Request.Form("keydata") 
		pw = Request.Form("password")
		IsAdmin = 0
		if session("uName") = ors.fields(2).value then
			pw = ors.fields(0).value
			IsAdmin = 1
		end if
		TournamentToJoin = Request.Form("TournamentToJoin")
		if ors.fields(0).value = pw then
			%>
			<br>Password matched
			<br>Ladder Selected: 
			<%
			Response.write Request.Form("TournamentToJoin")
			%><br>Player Selected:<%
			Response.Write struname
			ors.close
			strSQL = "select PlayerID from tbl_players where PlayerHandle = '" & replace(session("uName"), "'", "''") & "'" 
			ors.open strsql, oconn
				pID = ors.fields(0).value
			ors.close
			strsql="select * from lnk_T_M where TeamID = " & tID & " and TournamentID = " & Request.Form("TournamentID")
			ors.open strsql, oconn
			linkID = ors.fields(0).value
			ors.close
			strSQL="select count(PlayerID) from lnk_T_M_P inner join lnk_T_M on lnk_T_M_P.TMLinkID=lnk_T_M.TMLinkID where lnk_T_M_P.PlayerID=" & pid & " and lnk_T_M.TournamentID=" & Request.Form("TournamentID")
			ors.open strSQL, oconn
			lcnt=ors.fields(0).value
			Response.Write "<br>PID=" & pid
			Response.Write "<br>lcnt=" & lcnt
			
			if lcnt = 0 then
				strsql="select count(*) from lnk_T_M_P where PlayerID = " & pID & " and TMLinkID = " & linkID
				ors.close 
				ors.open strsql, oconn
				rcnt = ors.fields(0).value
				Response.Write pID
				Response.Write tID
				Response.Write lID
				response.Write linkID
				Response.Write rcnt		
				if rcnt = 0 then
					strsql="insert into lnk_T_M_P (PlayerID, TMLinkID, isAdmin, DateJoined) values (" & pID & ", " & linkId & "," & IsAdmin & " ,'" & now & "')"
					Response.Write strsql
					ors.close
					ors.open strsql, oconn
					Response.Write "<br>User add successful"
					fromurl=Request.Form ("fromurl")
					%>
					<script>
					window.opener.location='<%=fromurl%>';
					this.window.close();
					</script>
					<%
				end if
			else
		 	 fromurl=Request.Form ("fromurl")
			 tid=Request.Form ("teamid")
		 	 lid=Request.Form ("tournamentID")
			oConn.Close
			set ors=nothing
			set oconn=nothing
			Response.Clear 
			 response.redirect "joinTeamOnTournament.asp?teamID=" & tID & "&tournamentID=" & lid & "&type=join&error=2&url=" & fromurl
			end if
		else
		 fromurl=Request.Form ("fromurl")
		 tid=Request.Form ("teamid")
		 lid=Request.Form ("tournamentID")
			oConn.Close
			set ors=nothing
			set oconn=nothing
			Response.Clear 
		 response.redirect "joinTeamOnTournament.asp?teamID=" & tID & "&tournamentID=" & lid & "&type=join&error=1&url=" & fromurl
		end if
	end if
	fromurl=Request.Form("fromurl")
	if request("join") = 2 then 
		response.redirect "teamlist.asp"
	else
		%>
		<script>
		window.opener.location='<%=fromurl%>';
		this.window.close();
		</script>
		<%
	end if
end if

if SecType="founderjoin" then
	laddertojoin=request("ladder")
	teamtojoin=request("team")
	
	'==============
	' Get Team ID
	'==============
	strsql="select TeamID, TeamFounderID from tbl_Teams where TeamName='" & replace(teamtojoin, "'", "''") & "'"
	ors.open strsql,oconn
	if not (ors.bof and ors.eof) then
		tid=ors.fields(0).value
		pid=ors.fields(1).value
	end if
	ors.close
	'==============
	' Get Ladder ID
	'==============
	strsql="select LadderID from tbl_ladders where Laddername='" & replace(laddertojoin, "'", "''") & "'"
	ors.open strsql,oconn
	if not (ors.bof and ors.eof) then
		lid=ors.fields(0).value
	end if
	ors.close
		
	
	strsql="select TLLinkID from lnk_T_L where teamid=" & tid & " and ladderid=" & lid
	ors.open strsql, oconn
	linkID = ors.fields(0).value
	ors.close
		
	strSQL="select count(PlayerID) from lnk_T_P_L inner join lnk_T_L on lnk_T_P_L.TLLinkID=lnk_T_L.TLLinkID where lnk_T_P_L.PlayerID=" & pid & " and lnk_T_L.LadderID=" & lid
	ors.open strSQL, oconn
	lcnt=ors.fields(0).value
	Response.Write "<br>PID=" & pid
	Response.Write "<br>lcnt=" & lcnt
	
	if lcnt = 0 then
		strsql="select count(*) from lnk_T_P_L where PlayerID = " & pID & " and TLLinkID = " & linkID
		ors.close 
		ors.open strsql, oconn
		rcnt = ors.fields(0).value
		Response.Write pID
		Response.Write tID
		Response.Write lID
		response.Write linkID
		Response.Write rcnt		
		isadmin=1
		if rcnt = 0 then
			strsql="insert into lnk_T_P_L (PlayerID, TLLinkID, isAdmin, DateJoined) values (" & pID & ", " & linkId & "," & IsAdmin & " ,'" & now & "')"
			Response.Write strsql
			ors.close
			ors.open strsql, oconn
			Response.Write "<br>User add successful"
		end if
	
	end if
	oConn.Close
	set ors=nothing
	set oconn=nothing
	Response.Clear 
	response.redirect "viewteam.asp?team=" & server.urlencode(teamtojoin)
end if
on error resume next
oconn.close
on error goto 0
set oconn = nothing
set ors = nothing
%>