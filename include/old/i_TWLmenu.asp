<script language="JavaScript" type="text/javascript">
function fMenu2() {
	<% If Session("LoggedIn") Then %>
	
	fMakeMenu('my_twl','','<%=Server.HTMLEncode(Replace(Replace(Session("uName"), "\", "\\"), "'", "\'"))%>','default.asp')
			fMakeMenu('home','my_twl','Home', 'default.asp')
			fMakeMenu('account','my_twl','Account Maintenance')
			fMakeMenu('profile','account','Profile','viewplayer.asp?player=<%=Server.URLEncode(Session("uName"))%>')
			fMakeMenu('edit_profile','account','Edit Profile','addplayer.asp?IsEdit=true')
			fMakeMenu('register_team','account','Register Team','addteam.asp')
			fMakeMenu('logout','account','Logout', '', 'fPopLogin();')
			fMakeMenu('preferences','account','Preferences','preferences.asp')
	
		fMakeMenu('news_archive','my_twl','News Archive','newsarchive.asp')
		fMakeMenu('search','my_twl','Search','')
			fMakeMenu('search1','search','Find Player By Name','searchPlayerByName.asp')
			fMakeMenu('search4','search','Find Player By In Game Identifier','searchPlayerByIdentifier.asp')
			fMakeMenu('search2','search','Find Team By Name','searchTeamByName.asp')
			fMakeMenu('search3','search','Find Team By Tag','searchTeamByTag.asp')
		
		fMakeMenu('myteams','my_twl','My Teams')
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
						Response.Write "fMakeMenu('myteams" & intNumber & "','myteams','" & Server.HTMLEncode(Replace(Replace(ors.fields("TeamTag").value, "\", "\\"), "'", "\'")) & "','viewteam.asp?team=" & Server.URLEncode(ors.fields("TeamName").value) & "')" & vbcrlf
						strMenuTeamName = oRS.Fields("TeamName").Value
				End If
				oRS.MoveNext
			Loop
		Else
			Response.Write "fMakeMenu('myteams" & intNumber & "','myteams','No teams found.')" & vbcrlf
		End If
		oRS.NextRecordSet
	
		%>
		fMakeMenu('contactus','my_twl','Contact Us','staff.asp')
	
	<% Else %>
	fMakeMenu('my_twl','','my twl', 'default.asp')
		fMakeMenu('home','my_twl','Home', 'default.asp')
		fMakeMenu('login','my_twl','Login', '', 'fPopLogin();')
		fMakeMenu('forgot_password','my_twl','Forgot Password','forgotpassword.asp')
		fMakeMenu('activate','my_twl','Deactivated Account?','activate.asp')
		fMakeMenu('register','my_twl','Register','addplayer.asp')
		fMakeMenu('contactus','my_twl','Contact Us','staff.asp')
	<% End If %>
}

function fMenu4() {
<% If bSysAdmin Or bAnyLadderAdmin Then %>
	fMakeMenu('staffForums','forums','Staff Forums','')
		fMakeMenu('forum3','staffForums','TWL Staff','forums/forumdisplay.asp?forumid=3')
		fMakeMenu('forum30','staffForums','League Admin','forums/forumdisplay.asp?forumid=30')
		fMakeMenu('forum35','staffForums','Match Observers','forums/forumdisplay.asp?forumid=35')
		fMakeMenu('forum45','staffForums','Cheating','forums/forumdisplay.asp?forumid=45')
		fMakeMenu('forum5','staffForums','Development','forums/forumdisplay.asp?forumid=90')
		fMakeMenu('forum22','staffForums','Event Management','forums/forumdisplay.asp?forumid=22')
		fMakeMenu('forum85','staffForums','CyberXGaming','forums/forumdisplay.asp?forumid=85')
		fMakeMenu('forum92','staffForums','Punkbuster','forums/forumdisplay.asp?forumid=92')
		fMakeMenu('forum93','staffForums','AADS / OTTO Administration','forums/forumdisplay.asp?forumid=93')
		fMakeMenu('forum97','staffForums','AA Administration','forums/forumdisplay.asp?forumid=97')
	fMakeMenu('staffMDForums','staffForums','Match Disputes','forums/default.asp#Category7')
		fMakeMenu('forum42','staffMDForums','General','forums/forumdisplay.asp?forumid=42')
		fMakeMenu('forum43','staffMDForums','Closed','forums/forumdisplay.asp?forumid=43')
		fMakeMenu('forum60','staffMDForums','Americas Army','forums/forumdisplay.asp?forumid=60')
		fMakeMenu('forum61','staffMDForums','Battlefield 1942','forums/forumdisplay.asp?forumid=61')
		fMakeMenu('forum62','staffMDForums','Command &amp; Conquer: Generals','forums/forumdisplay.asp?forumid=62')
		fMakeMenu('forum72','staffMDForums','Counter Strike','forums/forumdisplay.asp?forumid=72')
		fMakeMenu('forum75','staffMDForums','Day of Defeat','forums/forumdisplay.asp?forumid=75')
		fMakeMenu('forum80','staffMDForums','Delta Force: Black Hawk Down','forums/forumdisplay.asp?forumid=80')
		fMakeMenu('forum76','staffMDForums','Ghost Recon','forums/forumdisplay.asp?forumid=76')
		fMakeMenu('forum79','staffMDForums','Global Operations','forums/forumdisplay.asp?forumid=79')
		fMakeMenu('forum78','staffMDForums','Jedi Outcast','forums/forumdisplay.asp?forumid=78')
		fMakeMenu('forum77','staffMDForums','Mechwarrior 4','forums/forumdisplay.asp?forumid=77')
		fMakeMenu('forum64','staffMDForums','Medal of Honor','forums/forumdisplay.asp?forumid=64')
		fMakeMenu('forum66','staffMDForums','Rainbow 6: Raven Shield','forums/forumdisplay.asp?forumid=66')
		fMakeMenu('forum70','staffMDForums','Return To Castle Wolfenstein','forums/forumdisplay.asp?forumid=70')
		fMakeMenu('forum73','staffMDForums','Soldier of Fortune 2','forums/forumdisplay.asp?forumid=73')
		fMakeMenu('forum71','staffMDForums','Tribes','forums/forumdisplay.asp?forumid=71')
		fMakeMenu('forum59','staffMDForums','Tribes 2','forums/forumdisplay.asp?forumid=59')
		fMakeMenu('forum68','staffMDForums','Urban Terror','forums/forumdisplay.asp?forumid=68')
		fMakeMenu('forum65','staffMDForums','UT2003','forums/forumdisplay.asp?forumid=65')
		fMakeMenu('forum69','staffMDForums','Vietcong','forums/forumdisplay.asp?forumid=69')
		fMakeMenu('forum67','staffMDForums','Warcraft 3','forums/forumdisplay.asp?forumid=67')
		fMakeMenu('forum81','staffMDForums','XBox: Ghost Recon','forums/forumdisplay.asp?forumid=81')
<% End If %>	

		<% If bSysAdmin Then %>
		fMakeMenu('sysAdminForums','staffForums','SysAdmin Forums','forums/forumdisplay.asp?forumid=8')
		fMakeMenu('forum8','sysAdminForums','SysAdmin Forum','forums/forumdisplay.asp?forumid=8')
		fMakeMenu('forum44','sysAdminForums','Quality Control','forums/forumdisplay.asp?forumid=44')
		fMakeMenu('forum96','sysAdminForums','TWLHosting.com','forums/forumdisplay.asp?forumid=96')
		<% End If %>
}

function fMenu7() {
	<% If bSysAdmin Or bAnyLadderAdmin Or IsAnyLeagueAdmin() Then %>
		fMakeMenu('admin','','admin')
			fMakeMenu('admain','admin','Menu','adminmenu.asp')
			fMakeMenu('adopsnew','admin','News','newsdesk.asp')
			fMakeMenu('teamLadMenu','admin','Team Ladders ','')
						fMakeMenu('adopsmat','teamLadMenu','Match','adminops.asp?aType=Match')
				fMakeMenu('adopsfor','teamLadMenu','Forfeit','adminops.asp?aType=Forfeit')
				fMakeMenu('adopshis','teamLadMenu','History','adminops.asp?aType=History')
				fMakeMenu('adopslad','teamLadMenu','Ladder','adminops.asp?aType=Ladder')
				fMakeMenu('adopsrank','teamLadMenu','Rank','adminops.asp?aType=Rank')
				<% If bSysadmin Then %>
				fMakeMenu('laAdmins','teamLadMenu','Assign Admins','assignadmin.asp')
				fMakeMenu('ladmatchoptions','teamLadMenu','Match Options','ladder/ladderoptions.asp')
				fMakeMenu('addladder','teamLadMenu','Add Ladder','addladder.asp')
				<% End If %>
			fMakeMenu('playerLadMenu','admin','Player Ladders ','')
				fMakeMenu('adopspmat','playerLadMenu','Match','adminops.asp?aType=PMatch')
				fMakeMenu('adopspfor','playerLadMenu','Forfeit','adminops.asp?aType=PForfeit')
				fMakeMenu('adopsplad','playerLadMenu','Ladder Admin','adminops.asp?aType=PLadder')
				fMakeMenu('adopsplrank','playerLadMenu','Player Rank','editplayerrank.asp')
				<% if bSysAdmin Then %>
				fMakeMenu('addpladmin','playerLadMenu','Add Player Ladder','addplayerladder.asp')
				fMakeMenu('listpladmin','playerLadMenu','List Player Ladders','playerladderlist.asp')
				<% end if %>
			fMakeMenu('leagueadmin','admin','Leagues','leagueadmin.asp')
				fMakeMenu('leaguegen','leagueadmin','General Admin','leagueadmin.asp')
			fMakeMenu('helprules','admin','Help/Rules','help/admin')
			fMakeMenu('reports','admin','Reports ','')
				fMakeMenu('plfor','reports','Player Forfeit Report','reports/playerforfietreport.asp')
				fMakeMenu('plrr','reports','Player Roster Report','reports/playerrosterreport.asp')
				fMakeMenu('rr','reports','Roster Report','reports/rosterreport.asp')
				fMakeMenu('act','reports','Ladder Activity','reports/activity.asp')
		<% If bSysAdmin Then %>
				fMakeMenu('leagueadd','leagueadmin','Add League','leagueadd.asp')
				fMakeMenu('leagueaa','leagueadmin','League Assign Admin','leagueassignadmin.asp')
			fMakeMenu('votingadmin','admin','Voting Booth ','')
				fMakeMenu('addballot','votingadmin','Add Ballot','ballot/addballot.asp')
				fMakeMenu('actballot','votingadmin','Activate Ballot','ballot/activateballot.asp')
				fMakeMenu('ballotresults','votingadmin','Old Ballot Results','ballot/results.asp')
			fMakeMenu('sysadminstuff','admin','Sysadmin Tools ','')
				fMakeMenu('menus','sysadminstuff','Update Menus','menu/updatemenus.asp')
				fMakeMenu('player','sysadminstuff','Delete / Sysadmin Player ','adminops.asp?aType=Player')
				fMakeMenu('team','sysadminstuff','Delete Team','adminops.asp?aType=Team')
				fMakeMenu('mm','sysadminstuff','Mass Mail','massmail.asp')
				fMakeMenu('emailsearch','sysadminstuff','Email Search','emailsearch.asp')
				fMakeMenu('status','sysadminstuff','Server Status','reports/server_status.asp')
				fMakeMenu('tracker','sysadminstuff','IP Tracker','tracker.asp')
				fMakeMenu('ipban','sysadminstuff','IP Banner','ipban.asp')
				fMakeMenu('gamelist','sysadminstuff','Game List','gamelist.asp')
				fMakeMenu('newgame','sysadminstuff','New Game','addgame.asp')
				fMakeMenu('forum','sysadminstuff','Forum','forums/admin')
				fMakeMenu('addtourny','sysadminstuff','Add Tournament','tournament/createtourny.asp')
		<% End If %>
	<% Else %>
		fMakeMenu('help', '', 'help', '', 'fPopHelp();') 
	<% End If %>
}
</script>
