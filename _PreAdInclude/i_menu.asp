<script language="JavaScript">
function fMenu2() {
	<% If Session("LoggedIn") Then %>
	
	oCMenu.makeMenu('my_twl','','<%=Server.HTMLEncode(Replace(Replace(Session("uName"), "\", "\\"), "'", "\'"))%>','default.asp')
			oCMenu.makeMenu('home','my_twl','Home', 'default.asp')
			oCMenu.makeMenu('account','my_twl','Account Maintenance')
			oCMenu.makeMenu('profile','account','Profile','viewplayer.asp?player=<%=Server.URLEncode(Session("uName"))%>')
			oCMenu.makeMenu('edit_profile','account','Edit Profile','addplayer.asp?IsEdit=true')
			oCMenu.makeMenu('register_team','account','Register Team','addteam.asp')
			oCMenu.makeMenu('logout','account','Logout', '', '', '','', '', '', '', '', '', '', '', 'javascript:popup("/login.asp?url=' + this.location.href + '", "login", 175, 300, "no");')
			oCMenu.makeMenu('preferences','account','Preferences','preferences.asp')
			oCMenu.makeMenu('requestnamechange','account','Request Name Change','request/ReqNameChange.asp?player=<%Server.URLEncode(Session("uName"))%>')
	
		oCMenu.makeMenu('news_archive','my_twl','News Archive','newsarchive.asp')
		oCMenu.makeMenu('search','my_twl','Search','')
			oCMenu.makeMenu('search1','search','Find Player By Name','searchPlayerByName.asp')
			oCMenu.makeMenu('search4','search','Find Player By In Game Identifier','searchPlayerByIdentifier.asp')
			oCMenu.makeMenu('search2','search','Find Team By Name','searchTeamByName.asp')
			oCMenu.makeMenu('search3','search','Find Team By Tag','searchTeamByTag.asp')
		
		oCMenu.makeMenu('myteams','my_twl','My Teams')
		<%	
		Dim intNumber, strCurrentTeam
		Dim strLName, strLAbbr, strTName, strTTag, strMenuTeamName
		intNumber = 0
		strMenuTeamName = -1
	
		strSQL = "EXECUTE PlayerGetTeams '" & Session("PlayerID") & "'"
		oRS.Open strSQL, oConn
		If Not (oRS.EOF AND oRS.BOF) Then
			Do While Not(oRS.EOF)
				If strMenuTeamName <> oRS.Fields("TeamName").Value Then 
						intNumber = intNumber + 1
						Response.Write "oCMenu.makeMenu('myteams" & intNumber & "','myteams','" & Server.HTMLEncode(Replace(Replace(ors.fields("TeamTag").value, "\", "\\"), "'", "\'")) & "','viewteam.asp?team=" & Server.URLEncode(ors.fields("TeamName").value) & "')" & vbcrlf
						strMenuTeamName = oRS.Fields("TeamName").Value
				End If
				oRS.MoveNext
			Loop
		Else
			Response.Write "oCMenu.makeMenu('myteams" & intNumber & "','myteams','No teams found.')" & vbcrlf
		End If
		oRS.NextRecordSet
	
		%>
		oCMenu.makeMenu('contactus','my_twl','Contact Us','staff.asp')
	
	<% Else %>
	oCMenu.makeMenu('my_twl','','my twl', 'default.asp')
		oCMenu.makeMenu('home','my_twl','Home', 'default.asp')
		oCMenu.makeMenu('login','my_twl','Login', '', '', '','', '', '', '', '', '', '', '', 'javascript:popup("/login.asp?url=' + this.location.href + '", "login", 175, 300, "no");')
		oCMenu.makeMenu('forgot_password','my_twl','Forgot Password','forgotpassword.asp')
		oCMenu.makeMenu('activate','my_twl','Deactivated Account?','activate.asp')
		oCMenu.makeMenu('register','my_twl','Register','addplayer.asp')
		oCMenu.makeMenu('contactus','my_twl','Contact Us','staff.asp')
	<% End If %>
}

function fMenu4() {
<% If bSysAdmin Or bAnyLadderAdmin Then %>
	oCMenu.makeMenu('staffForums','forums','Staff Forums','')
		oCMenu.makeMenu('forum3','staffForums','TWL Staff','forums/forumdisplay.asp?forumid=3')
		oCMenu.makeMenu('forum30','staffForums','League Admin','forums/forumdisplay.asp?forumid=30')
		oCMenu.makeMenu('forum35','staffForums','Match Observers','forums/forumdisplay.asp?forumid=35')
		oCMenu.makeMenu('forum45','staffForums','Cheating','forums/forumdisplay.asp?forumid=45')
		oCMenu.makeMenu('forum5','staffForums','Development','forums/forumdisplay.asp?forumid=90')
		oCMenu.makeMenu('forum22','staffForums','Event Management','forums/forumdisplay.asp?forumid=22')
		oCMenu.makeMenu('forum85','staffForums','CyberXGaming','forums/forumdisplay.asp?forumid=85')
		oCMenu.makeMenu('forum92','staffForums','Punkbuster','forums/forumdisplay.asp?forumid=92')
		oCMenu.makeMenu('forum93','staffForums','AADS / OTTO Administration','forums/forumdisplay.asp?forumid=93')
		oCMenu.makeMenu('forum97','staffForums','AA Administration','forums/forumdisplay.asp?forumid=97')
		<% If bSysAdmin Then %>
		oCMenu.makeMenu('sysAdminForums','staffForums','SysAdmin Forums','forums/forumdisplay.asp?forumid=8')
		oCMenu.makeMenu('forum8','sysAdminForums','SysAdmin Forum','forums/forumdisplay.asp?forumid=8')
		oCMenu.makeMenu('forum44','sysAdminForums','Quality Control','forums/forumdisplay.asp?forumid=44')
		oCMenu.makeMenu('forum96','sysAdminForums','TWLHosting.com','forums/forumdisplay.asp?forumid=96')
		<% End If %>
	oCMenu.makeMenu('staffMDForums','staffForums','Match Disputes','forums/default.asp#Category7')
		oCMenu.makeMenu('forum42','staffMDForums','General','forums/forumdisplay.asp?forumid=42')
		oCMenu.makeMenu('forum43','staffMDForums','Closed','forums/forumdisplay.asp?forumid=43')
		oCMenu.makeMenu('forum60','staffMDForums','Americas Army','forums/forumdisplay.asp?forumid=60')
		oCMenu.makeMenu('forum61','staffMDForums','Battlefield 1942','forums/forumdisplay.asp?forumid=61')
		oCMenu.makeMenu('forum62','staffMDForums','Command &amp; Conquer: Generals','forums/forumdisplay.asp?forumid=62')
		oCMenu.makeMenu('forum72','staffMDForums','Counter Strike','forums/forumdisplay.asp?forumid=72')
		oCMenu.makeMenu('forum75','staffMDForums','Day of Defeat','forums/forumdisplay.asp?forumid=75')
		oCMenu.makeMenu('forum80','staffMDForums','Delta Force: Black Hawk Down','forums/forumdisplay.asp?forumid=80')
		oCMenu.makeMenu('forum76','staffMDForums','Ghost Recon','forums/forumdisplay.asp?forumid=76')
		oCMenu.makeMenu('forum79','staffMDForums','Global Operations','forums/forumdisplay.asp?forumid=79')
		oCMenu.makeMenu('forum78','staffMDForums','Jedi Outcast','forums/forumdisplay.asp?forumid=78')
		oCMenu.makeMenu('forum77','staffMDForums','Mechwarrior 4','forums/forumdisplay.asp?forumid=77')
		oCMenu.makeMenu('forum64','staffMDForums','Medal of Honor','forums/forumdisplay.asp?forumid=64')
		oCMenu.makeMenu('forum66','staffMDForums','Rainbow 6: Raven Shield','forums/forumdisplay.asp?forumid=66')
		oCMenu.makeMenu('forum70','staffMDForums','Return To Castle Wolfenstein','forums/forumdisplay.asp?forumid=70')
		oCMenu.makeMenu('forum73','staffMDForums','Soldier of Fortune 2','forums/forumdisplay.asp?forumid=73')
		oCMenu.makeMenu('forum71','staffMDForums','Tribes','forums/forumdisplay.asp?forumid=71')
		oCMenu.makeMenu('forum59','staffMDForums','Tribes 2','forums/forumdisplay.asp?forumid=59')
		oCMenu.makeMenu('forum68','staffMDForums','Urban Terror','forums/forumdisplay.asp?forumid=68')
		oCMenu.makeMenu('forum65','staffMDForums','UT2003','forums/forumdisplay.asp?forumid=65')
		oCMenu.makeMenu('forum69','staffMDForums','Vietcong','forums/forumdisplay.asp?forumid=69')
		oCMenu.makeMenu('forum67','staffMDForums','Warcraft 3','forums/forumdisplay.asp?forumid=67')
		oCMenu.makeMenu('forum81','staffMDForums','XBox: Ghost Recon','forums/forumdisplay.asp?forumid=81')
<% End If %>	
}

function fMenu7() {
	<% If bSysAdmin Or bAnyLadderAdmin Or IsAnyLeagueAdmin() Then %>
		oCMenu.makeMenu('admin','','admin')
			oCMenu.makeMenu('admain','admin','Menu','adminmenu.asp')
			oCMenu.makeMenu('adopsnew','admin','News','newsdesk.asp')
			oCMenu.makeMenu('teamLadMenu','admin','Team Ladders ','')
						oCMenu.makeMenu('adopsmat','teamLadMenu','Match','adminops.asp?aType=Match')
				oCMenu.makeMenu('adopsfor','teamLadMenu','Forfeit','adminops.asp?aType=Forfeit')
				oCMenu.makeMenu('adopshis','teamLadMenu','History','adminops.asp?aType=History')
				oCMenu.makeMenu('adopslad','teamLadMenu','Ladder','adminops.asp?aType=Ladder')
				oCMenu.makeMenu('adopsrank','teamLadMenu','Rank','adminops.asp?aType=Rank')
				<% If bSysadmin Then %>
				oCMenu.makeMenu('laAdmins','teamLadMenu','Assign Admins','assignadmin.asp')
				oCMenu.makeMenu('ladmatchoptions','teamLadMenu','Match Options','ladder/ladderoptions.asp')
				oCMenu.makeMenu('addladder','teamLadMenu','Add Ladder','addladder.asp')
				<% End If %>
			oCMenu.makeMenu('playerLadMenu','admin','Player Ladders ','')
				oCMenu.makeMenu('adopspmat','playerLadMenu','Match','adminops.asp?aType=PMatch')
				oCMenu.makeMenu('adopspfor','playerLadMenu','Forfeit','adminops.asp?aType=PForfeit')
				oCMenu.makeMenu('adopsplad','playerLadMenu','Ladder Admin','adminops.asp?aType=PLadder')
				oCMenu.makeMenu('adopsplrank','playerLadMenu','Player Rank','editplayerrank.asp')
				oCMenu.makeMenu('adopsplhis','playerLadMenu','History','edit1v1history.asp')
				<% if bSysAdmin Then %>
				oCMenu.makeMenu('addpladmin','playerLadMenu','Add Player Ladder','addplayerladder.asp')
				oCMenu.makeMenu('listpladmin','playerLadMenu','List Player Ladders','playerladderlist.asp')
				<% end if %>
			oCMenu.makeMenu('leagueadmin','admin','Leagues','leagueadmin.asp')
				oCMenu.makeMenu('leaguegen','leagueadmin','General Admin','leagueadmin.asp')
			oCMenu.makeMenu('helprules','admin','Help/Rules','help/admin')
			oCMenu.makeMenu('reports','admin','Reports ','')
				oCMenu.makeMenu('plfor','reports','Player Forfeit Report','reports/playerforfietreport.asp')
				oCMenu.makeMenu('plrr','reports','Player Roster Report','reports/playerrosterreport.asp')
				oCMenu.makeMenu('rr','reports','Roster Report','reports/rosterreport.asp')
				oCMenu.makeMenu('act','reports','Ladder Activity','reports/activity.asp')
		<% If bSysAdmin Then %>
				oCMenu.makeMenu('leagueadd','leagueadmin','Add League','leagueadd.asp')
				oCMenu.makeMenu('leagueaa','leagueadmin','League Assign Admin','leagueassignadmin.asp')
			oCMenu.makeMenu('votingadmin','admin','Voting Booth ','')
				oCMenu.makeMenu('addballot','votingadmin','Add Ballot','ballot/addballot.asp')
				oCMenu.makeMenu('actballot','votingadmin','Activate Ballot','ballot/activateballot.asp')
				oCMenu.makeMenu('ballotresults','votingadmin','Old Ballot Results','ballot/results.asp')
			oCMenu.makeMenu('sysadminstuff','admin','Sysadmin Tools ','')
				oCMenu.makeMenu('menus','sysadminstuff','Update Menus','menu/updatemenus.asp')
				oCMenu.makeMenu('player','sysadminstuff','Delete / Sysadmin Player ','adminops.asp?aType=Player')
				oCMenu.makeMenu('team','sysadminstuff','Delete Team','adminops.asp?aType=Team')
				oCMenu.makeMenu('mm','sysadminstuff','Mass Mail','massmail.asp')
				oCMenu.makeMenu('emailsearch','sysadminstuff','Email Search','emailsearch.asp')
				oCMenu.makeMenu('status','sysadminstuff','Server Status','reports/server_status.asp')
				oCMenu.makeMenu('tracker','sysadminstuff','IP Tracker','tracker.asp')
				oCMenu.makeMenu('ipban','sysadminstuff','IP Banner','ipban.asp')
				oCMenu.makeMenu('gamelist','sysadminstuff','Game List','gamelist.asp')
				oCMenu.makeMenu('newgame','sysadminstuff','New Game','addgame.asp')
				oCMenu.makeMenu('forum','sysadminstuff','Forum','forums/admin')
				oCMenu.makeMenu('addtourny','sysadminstuff','Add Tournament','tournament/createtourny.asp')
		<% End If %>
	<% Else %>
		oCMenu.makeMenu('help', '', 'help', '', '', '','', '', '', '', '', '', '', '', 'javascript:popup("/help", "help", 300, 400, "yes");') 
	<% End If %>
}
</script>
