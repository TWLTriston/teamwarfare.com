<script language="JavaScript">
oCMenu=new makeCM("oCMenu") //Making the menu object. Argument: menuname

//Menu properties   
oCMenu.pxBetween=0
oCMenu.fromLeft=0 
oCMenu.fromTop=0   
oCMenu.rows=1 
oCMenu.menuPlacement="center"
                                                             
oCMenu.offlineRoot="/"
oCMenu.onlineRoot="/" 
oCMenu.resizeCheck=0
oCMenu.wait=500
oCMenu.fillImg=""
oCMenu.zIndex=0

//Background bar properties
oCMenu.useBar=0
oCMenu.barWidth="100%"
oCMenu.barHeight="menu" 
oCMenu.barClass="clBar"
oCMenu.barX=0 
oCMenu.barY=0
oCMenu.barBorderX=0
oCMenu.barBorderY=0
oCMenu.barBorderClass=""

oCMenu.level[0]=new cm_makeLevel() //Add this for each new level
oCMenu.level[0].width=130
oCMenu.level[0].height=19
if (bw.ie) {
	oCMenu.level[0].regClass="clLevel0"
	oCMenu.level[0].overClass="clLevel0over"
} else {
	oCMenu.level[0].regClass="clLevel0nonIE"
	oCMenu.level[0].overClass="clLevel0nonIEover"
}
oCMenu.level[0].borderX=0
oCMenu.level[0].borderY=0
oCMenu.level[0].borderClass="clLevel0border"
oCMenu.level[0].offsetX=0
oCMenu.level[0].offsetY=0
oCMenu.level[0].rows=0
oCMenu.level[0].arrow=0
oCMenu.level[0].arrowWidth=0
oCMenu.level[0].arrowHeight=0
oCMenu.level[0].align="bottom"

oCMenu.level[1]=new cm_makeLevel() //Add this for each new level (adding one to the number)
oCMenu.level[1].width=200
oCMenu.level[1].height=22
oCMenu.level[1].regClass="clLevel1"
oCMenu.level[1].overClass="clLevel1over"
oCMenu.level[1].borderX=1
oCMenu.level[1].borderY=1
oCMenu.level[1].align="left" 
oCMenu.level[1].arrow="images/tri.gif"
//oCMenu.level[1].arrow=0
oCMenu.level[1].arrowWidth=10
oCMenu.level[1].arrowHeight=9
oCMenu.level[1].offsetX=350
oCMenu.level[1].offsetY=0
oCMenu.level[1].borderClass="clLevel1border"

oCMenu.level[2]=new cm_makeLevel() //Add this for each new level (adding one to the number)
oCMenu.level[2].width=200
oCMenu.level[2].height=20
oCMenu.level[2].offsetX=350
oCMenu.level[2].offsetY=0
oCMenu.level[2].arrow="images/tri.gif"
oCMenu.level[2].regClass="clLevel2"
oCMenu.level[2].overClass="clLevel2over"
oCMenu.level[2].borderClass="clLevel2border"

oCMenu.level[3]=new cm_makeLevel() //Add this for each new level (adding one to the number)
oCMenu.level[3].width=200
oCMenu.level[3].height=20
oCMenu.level[3].offsetX=150
oCMenu.level[3].offsetY=0
oCMenu.level[3].arrow=0
oCMenu.level[3].regClass="clLevel3"
oCMenu.level[3].overClass="clLevel3over"
oCMenu.level[3].borderClass="clLevel3border"

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
	oCMenu.makeMenu('search','my_twl','Search','search.asp')
	
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

oCMenu.makeMenu('forums','','forums','forums/')
	oCMenu.makeMenu('forumindex','forums','Forum Index','forums/')
	oCMenu.makeMenu('forum9','forums','General Forums','forums/forumdisplay.asp?forumid=9')
	oCMenu.makeMenu('forum19','forums','Site Support','forums/forumdisplay.asp?forumid=19')
	oCMenu.makeMenu('forum11','forums','Recruiting','forums/forumdisplay.asp?forumid=11')
	oCMenu.makeMenu('gameForums','forums','Game Specific','forums/default.asp#Category1')
	oCMenu.makeMenu('forum54','gameForums','AA CDS Support','forums/forumdisplay.asp?forumid=54')
	oCMenu.makeMenu('forum25','gameForums','America\'s Army','forums/forumdisplay.asp?forumid=25')
	oCMenu.makeMenu('forum27','gameForums','Battlefield 1942','forums/forumdisplay.asp?forumid=27')
	oCMenu.makeMenu('cwforms','gameForums','Castle Wolfenstein','')
	oCMenu.makeMenu('forum12','cwforms','Castle Wolfenstein','forums/forumdisplay.asp?forumid=12')
	oCMenu.makeMenu('forum40','cwforms','Castle Wolfenstein: Enemy Territory','forums/forumdisplay.asp?forumid=40')
	oCMenu.makeMenu('forum52','cwforms','Castle Wolfenstein: Wolftactics','forums/forumdisplay.asp?forumid=52')
	oCMenu.makeMenu('forum38','gameForums','C&amp;C: Generals','forums/forumdisplay.asp?forumid=38')
	oCMenu.makeMenu('forum14','gameForums','Counter Strike','forums/forumdisplay.asp?forumid=14')
	oCMenu.makeMenu('forum51','gameForums','Delta Force: Black Hawk Down','forums/forumdisplay.asp?forumid=51')
	oCMenu.makeMenu('forum48','gameForums','Devastation','forums/forumdisplay.asp?forumid=48')
	oCMenu.makeMenu('forum58','gameForums','Day of Defeat','forums/forumdisplay.asp?forumid=58')
	oCMenu.makeMenu('forum24','gameForums','Ghost Recon','forums/forumdisplay.asp?forumid=24')
	oCMenu.makeMenu('forum21','gameForums','Global Operations','forums/forumdisplay.asp?forumid=21')
	oCMenu.makeMenu('forum18','gameForums','Jedi Outcast','forums/forumdisplay.asp?forumid=18')
	oCMenu.makeMenu('forum2','gameForums','Mechwarrior 4','forums/forumdisplay.asp?forumid=2')
	oCMenu.makeMenu('mohforum','gameForums','Medal Of Honor','')
	oCMenu.makeMenu('forum41','mohforum','Medal of Honor League','forums/forumdisplay.asp?forumid=41')
	oCMenu.makeMenu('forum13','mohforum','Medal of Honor: Allied Assault','forums/forumdisplay.asp?forumid=13')
	oCMenu.makeMenu('forum37','mohforum','Medal of Honor: Spearhead','forums/forumdisplay.asp?forumid=37')
	oCMenu.makeMenu('forum23','gameForums','Soldier Of Fortune 2','forums/forumdisplay.asp?forumid=23')
	oCMenu.makeMenu('forum34','gameForums','Rainbow Six: Raven Shield','forums/forumdisplay.asp?forumid=34')
	oCMenu.makeMenu('forum1','gameForums','Tribes','forums/forumdisplay.asp?forumid=1')
	oCMenu.makeMenu('t2forum','gameForums','Tribes 2','')
	oCMenu.makeMenu('forum4','t2forum','Tribes 2 CTF','forums/forumdisplay.asp?forumid=4')
	oCMenu.makeMenu('forum53','t2forum','Tribes 2 Base','forums/forumdisplay.asp?forumid=53')
	oCMenu.makeMenu('forum17','t2forum','Tribes Renegades','forums/forumdisplay.asp?forumid=17')
	oCMenu.makeMenu('forum10','gameForums','UT2003','forums/forumdisplay.asp?forumid=10')
	oCMenu.makeMenu('forum47','gameForums','Vietcong','forums/forumdisplay.asp?forumid=47')
	oCMenu.makeMenu('forum26','gameForums','Warcraft 3','forums/forumdisplay.asp?forumid=26')
	oCMenu.makeMenu('forum32','gameForums','Urban Terror','forums/forumdisplay.asp?forumid=32')
	oCMenu.makeMenu('forum35','forums','Match Observers','forums/forumdisplay.asp?forumid=35')
	oCMenu.makeMenu('xboxforum','forums','Xbox Live','forums/forumdisplay.asp?forumid=57')
		oCMenu.makeMenu('forum57','xboxforum','General Xbox Live','forums/forumdisplay.asp?forumid=57')
		oCMenu.makeMenu('forum84','xboxforum','Castle Wolfenstein','forums/forumdisplay.asp?forumid=84')
		oCMenu.makeMenu('forum56','xboxforum','Ghost Recon','forums/forumdisplay.asp?forumid=56')
	oCMenu.makeMenu('ps2fourm','forums','Play Station 2 Online','forums/forumdisplay.asp?forumid=87')
		oCMenu.makeMenu('forum87','ps2fourm','General PS2 Online','forums/forumdisplay.asp?forumid=87')
		oCMenu.makeMenu('forum86','ps2fourm','Tribes: Aerial Assault','forums/forumdisplay.asp?forumid=86')
<% If bSysAdmin Or bAnyLadderAdmin Then %>
	oCMenu.makeMenu('staffForums','forums','Staff Forums','')
		oCMenu.makeMenu('forum3','staffForums','TWL Staff','forums/forumdisplay.asp?forumid=3')
		oCMenu.makeMenu('forum30','staffForums','League Admin','forums/forumdisplay.asp?forumid=30')
		oCMenu.makeMenu('forum5','staffForums','Development','forums/forumdisplay.asp?forumid=5')
		oCMenu.makeMenu('forum22','staffForums','Event Management','forums/forumdisplay.asp?forumid=22')
		<% If bSysAdmin Then %>
		oCMenu.makeMenu('forum8','staffForums','SysAdmin Forum','forums/forumdisplay.asp?forumid=8')
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
		oCMenu.makeMenu('forum74','staffMDForums','Devastation','forums/forumdisplay.asp?forumid=74')
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

// Start DB Menus
	oCMenu.makeMenu('dbm1','','competition','ladderlist.asp')
		oCMenu.makeMenu('dbm211','dbm1','Full Competition List','ladderlist.asp')
		oCMenu.makeMenu('dbm4','dbm1','America\'s Army','')
			oCMenu.makeMenu('dbm235','dbm4','Europe','')
				oCMenu.makeMenu('dbm308','dbm235','Combat 2v2','viewladder.asp?ladder=America%27s+Army++2v2%2D+Europe')
				oCMenu.makeMenu('dbm288','dbm235','CQB 7v7','viewladder.asp?ladder=America%27s+Army+Euro+CQB+7v7')
				oCMenu.makeMenu('dbm62','dbm235','Objective 4v4','viewladder.asp?Ladder=America%27s+Army+4v4+%2D+Europe')
				oCMenu.makeMenu('dbm61','dbm235','Objective 6v6','viewladder.asp?ladder=America%27s+Army+Objective+6v6+%2D+Europe')
			oCMenu.makeMenu('dbm236','dbm4','Leagues','')
				oCMenu.makeMenu('dbm334','dbm236','6v6 Invitational','viewleague.asp?league=AA+6v6+Invitational')
				oCMenu.makeMenu('dbm57','dbm236','6v6 Open','viewleague.asp?league=AA+6+v+6+Open')
				oCMenu.makeMenu('dbm271','dbm236','Euro Objective 6v6','viewleague.asp?league=AA+Euro+6v6+Objective+')
			oCMenu.makeMenu('dbm433','dbm4','North America','')
				oCMenu.makeMenu('dbm233','dbm433','Combat 2v2','viewladder.asp?ladder=America%27s+Army+2v2')
				oCMenu.makeMenu('dbm437','dbm433','Combat 6v6','viewladder.asp?ladder=America%27s+Army+Team+Combat')
				oCMenu.makeMenu('dbm435','dbm433','Objective 4v4','viewladder.asp?ladder=America%27s+Army+4v4')
				oCMenu.makeMenu('dbm438','dbm433','Objective 6v6','viewladder.asp?ladder=America%27s+Army+Objective+6v6')
				oCMenu.makeMenu('dbm440','dbm433','Objective 8v8','viewladder.asp?ladder=America%27s+Army+Objective+8v8')
				oCMenu.makeMenu('dbm439','dbm433','Objective CQB 6v6','viewladder.asp?ladder=America%27s+Army+CQB+6v6+')
			oCMenu.makeMenu('dbm424','dbm4','Oceania','')
				oCMenu.makeMenu('dbm425','dbm424','Combat 2v2','viewladder.asp?ladder=America%27s+Army+Oceanic+Combat+2v2')
				oCMenu.makeMenu('dbm426','dbm424','Objective 4v4','viewladder.asp?ladder=America%27s+Army+Oceanic+Objective+4v4')
		oCMenu.makeMenu('dbm5','dbm1','Battlefield 1942','')
			oCMenu.makeMenu('dbm68','dbm5','Capture The Flag 8v8','viewladder.asp?Ladder=Battlefield+1942+CTF')
			oCMenu.makeMenu('dbm66','dbm5','Conquest 12v12','viewladder.asp?Ladder=Battlefield+1942+Conquest')
			oCMenu.makeMenu('dbm67','dbm5','Conquest 8v8','viewladder.asp?Ladder=Battlefield+1942+Conquest+%2D+8+man')
			oCMenu.makeMenu('dbm459','dbm5','Conquest League 8v8','viewleague.asp?league=Battlefield+1942+8v8')
			oCMenu.makeMenu('dbm337','dbm5','Conquest League Season #2 12v12','viewleague.asp?league=Battlefield+1942+%2D+Season+2')
			oCMenu.makeMenu('dbm460','dbm5','Conquest League Season #2 Playoffs','tournament/viewBracket.asp?tournament=BF1942+Season+%232+Playoffs&amp;div=2')
			oCMenu.makeMenu('dbm412','dbm5','Desert Combat Conquest 10v10','viewladder.asp?ladder=BF1942%3A+Desert+Combat+Conquest+10v10')
		oCMenu.makeMenu('dbm290','dbm1','Command &amp; Conquer Generals','')
			oCMenu.makeMenu('dbm291','dbm290','1v1','viewplayerladder.asp?ladder=Command+and+Conquer+Generals+1v1')
			oCMenu.makeMenu('dbm292','dbm290','2v2','viewladder.asp?ladder=Command+and+Conquer+Generals+2v2')
			oCMenu.makeMenu('dbm453','dbm290','3v3','viewladder.asp?ladder=Command+%26+Conquer+Generals+3v3')
			oCMenu.makeMenu('dbm312','dbm290','Europe','')
				oCMenu.makeMenu('dbm313','dbm312','2v2','viewladder.asp?ladder=Command+and+Conquer+Generals+Euro+2v2')
		oCMenu.makeMenu('dbm6','dbm1','Counter-Strike','')
			oCMenu.makeMenu('dbm55','dbm6','2v2','viewladder.asp?Ladder=CounterStrike+2v2')
			oCMenu.makeMenu('dbm56','dbm6','5v5','viewladder.asp?Ladder=CounterStrike+Open')
		oCMenu.makeMenu('dbm420','dbm1','Day of Defeat','')
			oCMenu.makeMenu('dbm421','dbm420','6v6','viewladder.asp?ladder=Day+of+Defeat+6v6')
		oCMenu.makeMenu('dbm351','dbm1','Delta Force: Black Hawk Down','')
			oCMenu.makeMenu('dbm352','dbm351','CTF 6v6','viewladder.asp?ladder=Delta+Force%3ABlack+Hawk+Down+CTF+6v6')
			oCMenu.makeMenu('dbm466','dbm351','TDM 8v8','viewladder.asp?ladder=Delta+Force%3ABlack+Hawk+Down+TDM+8v8')
			oCMenu.makeMenu('dbm401','dbm351','TKOTH 8v8','viewladder.asp?ladder=Delta+Force%3ABlack+Hawk+Down+TKOTH+8v8')
		oCMenu.makeMenu('dbm349','dbm1','Devastation','')
			oCMenu.makeMenu('dbm350','dbm349','Composite 6v6','viewladder.asp?ladder=Devastation+Composite+6v6')
		oCMenu.makeMenu('dbm8','dbm1','Ghost Recon','')
			oCMenu.makeMenu('dbm387','dbm8','1v1','viewPlayerladder.asp?ladder=Ghost+Recon+1v1')
			oCMenu.makeMenu('dbm52','dbm8','Last Man Standing 4v4','viewladder.asp?ladder=Ghost+Recon+Last+Man+Standing')
		oCMenu.makeMenu('dbm7','dbm1','Global Operations','')
			oCMenu.makeMenu('dbm296','dbm7','4v4','viewladder.asp?ladder=Global+Operations+4v4')
			oCMenu.makeMenu('dbm53','dbm7','6v6','viewladder.asp?Ladder=Global+Operations')
			oCMenu.makeMenu('dbm54','dbm7','Europe 6v6','viewladder.asp?Ladder=Global+Operations+%2D+Europe')
		oCMenu.makeMenu('dbm9','dbm1','Jedi Knight 2: Jedi Outcast','')
			oCMenu.makeMenu('dbm98','dbm9','Full Force Saber Duel 1v1','viewPlayerladder.asp?ladder=Jedi+Outcast+1v1+Full+Force+Sabers')
			oCMenu.makeMenu('dbm99','dbm9','No Force Saber Duel 1v1','viewPlayerladder.asp?ladder=Jedi+Outcast+1v1+No+Force+Sabers')
		oCMenu.makeMenu('dbm10','dbm1','MechWarrior 4','')
			oCMenu.makeMenu('dbm243','dbm10','Mercenaries','viewladder.asp?ladder=MW4+Mercenaries')
			oCMenu.makeMenu('dbm281','dbm10','Mercenaries 2v2','viewladder.asp?ladder=MW4+Mercs+2v2')
			oCMenu.makeMenu('dbm416','dbm10','Mercenaries Euro 2v2','viewladder.asp?ladder=MW4+Mercs++Euro+2v2')
			oCMenu.makeMenu('dbm347','dbm10','Mercs Duel','viewplayerladder.asp?ladder=MW4+Mercs+Duel')
		oCMenu.makeMenu('dbm11','dbm1','Medal Of Honor','')
			oCMenu.makeMenu('dbm318','dbm11','Allied Assault','')
				oCMenu.makeMenu('dbm457','dbm318','Composite Playoffs &amp; Finals Season 1','tournament/viewBracket.asp?tournament=MoH%3AAA+Invitational+Composite+Playoffs&amp;div=2')
				oCMenu.makeMenu('dbm319','dbm318','Invitational Composite League 7v7','viewleague.asp?league=MoH%3AAA+Invitational+Composite+7v7')
				oCMenu.makeMenu('dbm464','dbm318','Invitational TDM 6v6 League','viewleague.asp?league=MoH%3AAA+Invitational+TDM+6v6')
				oCMenu.makeMenu('dbm386','dbm318','Objective 7v7','viewladder.asp?ladder=MoH%3AAA+OBJ+7v7')
				oCMenu.makeMenu('dbm378','dbm318','Team Death Match 2v2','viewladder.asp?ladder=MoH%3AAA+2+Man+TDM')
				oCMenu.makeMenu('dbm371','dbm318','Team Death Match 6v6','viewladder.asp?ladder=MoH%3AAA+Team+Deathmatch')
			oCMenu.makeMenu('dbm70','dbm11','Allied Assault Realism','')
				oCMenu.makeMenu('dbm104','dbm70','Objective 7v7','viewladder.asp?ladder=MoH%3AAA+Objective+Realism')
				oCMenu.makeMenu('dbm105','dbm70','Round Based Team Deathmatch 7v7','viewladder.asp?ladder=MoH%3AAA+Round+Based+TDM+Realism')
			oCMenu.makeMenu('dbm110','dbm11','Spearhead Realism','')
				oCMenu.makeMenu('dbm113','dbm110','Objective 7v7','viewladder.asp?ladder=MoH%3ASpearhead+OBJ+Realism+7v7')
				oCMenu.makeMenu('dbm114','dbm110','Round Based Team Deathmatch 7v7','viewladder.asp?ladder=MoH%3ASpearhead+RB+TDM+Realism+7v7')
				oCMenu.makeMenu('dbm111','dbm110','Team Deathmatch 2v2','viewladder.asp?ladder=MoH%3ASpearhead+TDM+Realism+2v2')
				oCMenu.makeMenu('dbm112','dbm110','Team Deathmatch 6v6','viewladder.asp?ladder=MoH%3ASpearhead+TDM+Realism+6v6+')
			oCMenu.makeMenu('dbm71','dbm11','Spearhead Standard','')
				oCMenu.makeMenu('dbm115','dbm71','Objective 5v5','viewladder.asp?ladder=MoH%3ASpearhead+OBJ+5v5')
				oCMenu.makeMenu('dbm259','dbm71','Objective 7v7','viewladder.asp?ladder=MoH%3ASpearhead+OBJ+7v7')
				oCMenu.makeMenu('dbm107','dbm71','Round Based Team Deathmatch 7v7','viewladder.asp?ladder=MoH%3ASpearhead+RB+TDM+7v7')
				oCMenu.makeMenu('dbm108','dbm71','Team Deathmatch 2v2','viewladder.asp?ladder=MoH%3ASpearhead+TDM+2v2')
				oCMenu.makeMenu('dbm109','dbm71','Team Deathmatch 6v6','viewladder.asp?ladder=MoH%3ASpearhead+TDM+6v6')
				oCMenu.makeMenu('dbm116','dbm71','Tug of War 7v7','viewladder.asp?ladder=MoH%3ASpearhead+TOW+7v7')
		oCMenu.makeMenu('dbm451','dbm1','PS2: Tribes Aerial Assault','')
			oCMenu.makeMenu('dbm452','dbm451','Capture the Flag 8v8','viewladder.asp?ladder=Tribes+Aerial+Assault+CTF+8v8')
			oCMenu.makeMenu('dbm463','dbm451','Duel 1v1','viewplayerladder.asp?ladder=Tribes+Aerial+Assault+Duel')
		oCMenu.makeMenu('dbm63','dbm1','Rainbow 6: Raven Shield','')
			oCMenu.makeMenu('dbm355','dbm63','Europe','')
				oCMenu.makeMenu('dbm353','dbm355','Team Survival Europe 4v4','viewladder.asp?ladder=Rainbow+6%3A+Raven+Shield+Team+Survival+Euro+4v4')
				oCMenu.makeMenu('dbm356','dbm355','Team Survival Europe 7v7','viewladder.asp?ladder=Rainbow+6%3A+Raven+Shield+Team+Survival+Euro+7v7')
			oCMenu.makeMenu('dbm357','dbm63','Leagues','')
				oCMenu.makeMenu('dbm370','dbm357','Composite Objective 5v5','viewleague.asp?league=Rainbow+6%3A+Raven+Sheild+Composite+5v5')
				oCMenu.makeMenu('dbm320','dbm357','Team Survival League 5v5','viewleague.asp?league=Rainbow+6%3A+Raven+Shield+Team+Survival+5v5')
			oCMenu.makeMenu('dbm359','dbm63','Objective','')
				oCMenu.makeMenu('dbm360','dbm359','Composite Objective 6v6','viewladder.asp?ladder=Rainbow+6%3A+Raven+Shield+Composite+Objective+6v6')
			oCMenu.makeMenu('dbm354','dbm63','Team Survival','')
				oCMenu.makeMenu('dbm399','dbm354','2v2','viewladder.asp?ladder=Rainbow+6%3A+Raven+Shield+Team+Survival+2v2')
				oCMenu.makeMenu('dbm257','dbm354','4v4','viewladder.asp?ladder=Rainbow+6%3A+Raven+Shield+Team+Survival+4v4')
				oCMenu.makeMenu('dbm64','dbm354','7v7','viewladder.asp?ladder=Rainbow+6%3A+Raven+Shield+Team+Survival+7v7')
			oCMenu.makeMenu('dbm358','dbm63','Tournaments','')
				oCMenu.makeMenu('dbm316','dbm358','Operation Triple Threat (Closed)','tournament/viewBracket.asp?tournament=Operation+Triple+Threat')
		oCMenu.makeMenu('dbm14','dbm1','Return To Castle Wolfenstein','')
			oCMenu.makeMenu('dbm383','dbm14','Custom','')
				oCMenu.makeMenu('dbm384','dbm383','Custom 6v6','viewladder.asp?ladder=RTCW+Custom+')
			oCMenu.makeMenu('dbm391','dbm14','Europe','')
				oCMenu.makeMenu('dbm392','dbm391','StopWatch 6v6','viewladder.asp?ladder=RTCW+Euro+6v6')
			oCMenu.makeMenu('dbm247','dbm14','Leagues','')
				oCMenu.makeMenu('dbm414','dbm247','Alpha League Season 3','viewleague.asp?league=RtCW+Alpha+Season+%233')
				oCMenu.makeMenu('dbm415','dbm247','Beta League Season 3','viewleague.asp?league=RtCW+Beta+Season+%233')
				oCMenu.makeMenu('dbm382','dbm247','Invitational League Season 3','viewleague.asp?league=RtCW+Invite++Season+%233+')
			oCMenu.makeMenu('dbm248','dbm14','One Life To Live','')
				oCMenu.makeMenu('dbm346','dbm248','One Life to Live 5v5','viewladder.asp?ladder=RTCW+%2D+One+Life+to+Live+5v5')
				oCMenu.makeMenu('dbm88','dbm248','One Life To Live 7v7','viewladder.asp?ladder=RTCW+%2D+One+Life+to+Live')
			oCMenu.makeMenu('dbm448','dbm14','Shrub','')
				oCMenu.makeMenu('dbm449','dbm448','Shrub 4Lives 8v8','viewladder.asp?ladder=RTCW+Shrub+4Lives+8v8')
				oCMenu.makeMenu('dbm462','dbm448','Shrub 8v8','viewladder.asp?ladder=RTCW+Shrub++8v8')
			oCMenu.makeMenu('dbm249','dbm14','StopWatch','')
				oCMenu.makeMenu('dbm87','dbm249','StopWatch 5v5','viewladder.asp?ladder=RTCW+5v5')
				oCMenu.makeMenu('dbm310','dbm249','StopWatch 6v6','viewladder.asp?ladder=RTCW+6v6')
				oCMenu.makeMenu('dbm86','dbm249','StopWatch 7v7','viewladder.asp?ladder=RTCW+7v7')
		oCMenu.makeMenu('dbm12','dbm1','Soldier of Fortune 2','')
			oCMenu.makeMenu('dbm72','dbm12','Capture The Flag 8v8','viewladder.asp?ladder=Soldier+of+Fortune+2+CTF')
			oCMenu.makeMenu('dbm73','dbm12','Elimination 2v2','viewladder.asp?ladder=Soldier+of+Fortune+2+Elimination+2v2')
			oCMenu.makeMenu('dbm74','dbm12','Infiltration East 5v5','viewladder.asp?ladder=Soldier+of+Fortune+2+East')
			oCMenu.makeMenu('dbm297','dbm12','Infiltration League','viewleague.asp?league=Soldier+of+Fortune+2+Infiltration')
			oCMenu.makeMenu('dbm75','dbm12','Infiltration West 5v5','viewladder.asp?ladder=Soldier+of+Fortune+2+West')
		oCMenu.makeMenu('dbm13','dbm1','Tribes','')
			oCMenu.makeMenu('dbm230','dbm13','Arena','')
				oCMenu.makeMenu('dbm90','dbm230','2v2','viewladder.asp?ladder=Tribes+Arena+2v2')
				oCMenu.makeMenu('dbm305','dbm230','4v4','viewladder.asp?ladder=Tribes+Arena+4v4')
			oCMenu.makeMenu('dbm231','dbm13','Capture The Flag','')
				oCMenu.makeMenu('dbm91','dbm231','10v10','viewladder.asp?ladder=Tribes+CTF+10v10')
				oCMenu.makeMenu('dbm227','dbm231','5v5','viewladder.asp?ladder=Tribes+CTF+5v5')
			oCMenu.makeMenu('dbm229','dbm13','Duel','')
				oCMenu.makeMenu('dbm92','dbm229','Duel Mod 1v1','viewplayerladder.asp?Ladder=Tribes+Duel')
		oCMenu.makeMenu('dbm3','dbm1','Tribes 2','')
			oCMenu.makeMenu('dbm241','dbm3','Arena','')
				oCMenu.makeMenu('dbm38','dbm241','2v2','viewladder.asp?ladder=Tribes+2+Arena+2v2')
				oCMenu.makeMenu('dbm37','dbm241','6v6','viewladder.asp?ladder=Tribes+2+Arena+6v6')
				oCMenu.makeMenu('dbm324','dbm241','Classic 2v2','viewladder.asp?ladder=Tribes+2+Classic+Arena+2v2')
				oCMenu.makeMenu('dbm280','dbm241','Classic 4v4','viewladder.asp?ladder=Tribes+2+Classic+Arena+4v4')
			oCMenu.makeMenu('dbm20','dbm3','Capture the Flag','')
				oCMenu.makeMenu('dbm33','dbm20','12v12','viewladder.asp?ladder=Tribes+2+CTF+12v12')
				oCMenu.makeMenu('dbm24','dbm20','7v7','viewladder.asp?ladder=Tribes+2+CTF+7v7')
				oCMenu.makeMenu('dbm35','dbm20','Classic 14v14 ','viewladder.asp?ladder=Tribes+2+Classic+CTF+14v14')
				oCMenu.makeMenu('dbm34','dbm20','Classic 7v7','viewladder.asp?ladder=Tribes+2+Classic+CTF+7v7')
			oCMenu.makeMenu('dbm23','dbm3','Duel','')
				oCMenu.makeMenu('dbm40','dbm23','Classic Duel Mod 1v1','viewplayerladder.asp?ladder=Tribes+2+Classic+Duel')
				oCMenu.makeMenu('dbm39','dbm23','Duel Mod 1v1','viewplayerladder.asp?ladder=Tribes+2+Duel')
			oCMenu.makeMenu('dbm393','dbm3','Leagues','')
				oCMenu.makeMenu('dbm394','dbm393','Classic CTF 10v10','viewleague.asp?league=Tribes+2+Classic+CTF+10v10')
			oCMenu.makeMenu('dbm25','dbm3','Oceania','')
				oCMenu.makeMenu('dbm28','dbm25','Classic Capture The Flag 12v12 ','viewladder.asp?ladder=Tribes+2+Oceania+Classic+CTF')
				oCMenu.makeMenu('dbm27','dbm25','Classic Duel Mod 1v1','viewplayerladder.asp?ladder=Tribes+2+Oceania+Classic+Duel')
			oCMenu.makeMenu('dbm22','dbm3','Renegades','')
				oCMenu.makeMenu('dbm30','dbm22','CTF 12v12','viewladder.asp?ladder=Tribes+2+Renegades+CTF+12v12')
				oCMenu.makeMenu('dbm31','dbm22','Ren Spawn Duel 1v1','viewPlayerladder.asp?ladder=Tribes+2+Renegades+Duel')
			oCMenu.makeMenu('dbm402','dbm3','Siege','')
				oCMenu.makeMenu('dbm403','dbm402','Classic 8v8','viewladder.asp?ladder=Tribes+2+Classic+Siege')
			oCMenu.makeMenu('dbm372','dbm3','Team Gauntlet','')
				oCMenu.makeMenu('dbm373','dbm372','5v5','viewladder.asp?ladder=Tribes+2+Team+Gauntlet+5v5')
				oCMenu.makeMenu('dbm374','dbm372','Classic 5v5','viewladder.asp?ladder=Tribes+2+Classic+Team+Gauntlet+5v5')
		oCMenu.makeMenu('dbm16','dbm1','Unreal Tournament 2003','')
			oCMenu.makeMenu('dbm76','dbm16','Deathmatch 1v1','viewPlayerladder.asp?ladder=Unreal+Tournament+2003+%2D+1v1+DM')
		oCMenu.makeMenu('dbm15','dbm1','Urban Terror','')
			oCMenu.makeMenu('dbm48','dbm15','Capture The Flag 6v6','viewladder.asp?ladder=Urban+Terror+CTF')
			oCMenu.makeMenu('dbm293','dbm15','CTF League','viewleague.asp?league=Urban+Terror+CTF')
			oCMenu.makeMenu('dbm343','dbm15','Team Survivor 2v2','viewladder.asp?ladder=Urban+Terror+2v2+TS')
			oCMenu.makeMenu('dbm47','dbm15','Team Survivor 6v6','viewladder.asp?ladder=Urban+Terror+TS')
		oCMenu.makeMenu('dbm330','dbm1','Vietcong','')
			oCMenu.makeMenu('dbm342','dbm330','8v8','viewladder.asp?ladder=Vietcong+Composite+8v8')
		oCMenu.makeMenu('dbm404','dbm1','Xbox Live','')
			oCMenu.makeMenu('dbm405','dbm404','Ghost Recon','')
				oCMenu.makeMenu('dbm406','dbm405','4v4 Team Deathmatch','viewladder.asp?ladder=Xbox+%2D+Ghost+Recon+4v4+Team+Deathmatch')
	oCMenu.makeMenu('dbm2','','rules','')
		oCMenu.makeMenu('dbm117','dbm2','America\'s Army','')
			oCMenu.makeMenu('dbm239','dbm117','Europe','')
				oCMenu.makeMenu('dbm309','dbm239','Combat 2v2','rules.asp?set=America%27s+Army+2+v+2++%2D+Europe')
				oCMenu.makeMenu('dbm289','dbm239','CQB 7v7','rules.asp?set=America%27s+Army+Euro+CQB+7v7')
				oCMenu.makeMenu('dbm146','dbm239','Objective 4v4','rules.asp?set=America%27s+Army+4v4+%2D+Europe')
				oCMenu.makeMenu('dbm145','dbm239','Objective 6v6','rules.asp?set=America%27s+Army+Team+Objective+%2D+Europe')
			oCMenu.makeMenu('dbm240','dbm117','Leagues','')
				oCMenu.makeMenu('dbm335','dbm240','6v6 Invitational','rules.asp?set=America%27s+Army+6+v+6++Invitational+League')
				oCMenu.makeMenu('dbm140','dbm240','6v6 Open','rules.asp?set=America%27s+Army+6+v+6++Open+League')
				oCMenu.makeMenu('dbm272','dbm240','Euro Objective 6v6','rules.asp?set=America%27s+Army+European+6+v+6+League')
			oCMenu.makeMenu('dbm441','dbm117','North America','')
				oCMenu.makeMenu('dbm442','dbm441','Combat 2v2','rules.asp?set=America%27s+Army+2+v+2')
				oCMenu.makeMenu('dbm443','dbm441','Combat 6V6','rules.asp?set=America%27s+Army+Team+Combat')
				oCMenu.makeMenu('dbm444','dbm441','Objective 4V4','rules.asp?set=America%27s+Army+4+v+4')
				oCMenu.makeMenu('dbm445','dbm441','Objective 6V6','rules.asp?set=America%27s+Army+Objective+')
				oCMenu.makeMenu('dbm447','dbm441','Objective 8v8','rules.asp?set=America%27s+Army+8+v+8')
				oCMenu.makeMenu('dbm446','dbm441','Objective CQB 6v6','rules.asp?set=America%27s+Army+CQB+6v6')
			oCMenu.makeMenu('dbm428','dbm117','Oceania','')
				oCMenu.makeMenu('dbm431','dbm428','Combat 2v2','rules.asp?set=America%27s+Army+2v2%2D+Oceania')
				oCMenu.makeMenu('dbm432','dbm428','Objective 4v4','rules.asp?set=America%27s+Army+4v4+%2D+Oceania')
		oCMenu.makeMenu('dbm118','dbm2','Battlefield 1942','')
			oCMenu.makeMenu('dbm147','dbm118','Conquest 12v12','rules.asp?set=Battlefield+1942+Conquest')
			oCMenu.makeMenu('dbm148','dbm118','Conquest 8v8','rules.asp?set=Battlefield+1942+Conquest+%2D+8+man')
			oCMenu.makeMenu('dbm268','dbm118','CTF 8v8','rules.asp?set=Battlefield+1942+Capture+the+Flag')
			oCMenu.makeMenu('dbm413','dbm118','Desert Combat Conquest 10v10','rules.asp?set=Desert+Combat+Conquest+')
			oCMenu.makeMenu('dbm264','dbm118','League 12v12','rules.asp?set=Battlefield+1942+Short+League')
			oCMenu.makeMenu('dbm461','dbm118','League 8v8','rules.asp?set=Battlefield+1942+Short+League')
		oCMenu.makeMenu('dbm298','dbm2','Command &amp; Conquer Generals','')
			oCMenu.makeMenu('dbm300','dbm298','1v1 Ladder','rules.asp?set=1v1+Ladder+Rules')
			oCMenu.makeMenu('dbm301','dbm298','2v2 Ladder','rules.asp?set=2v2+Ladder+Rules')
			oCMenu.makeMenu('dbm314','dbm298','Europe','')
				oCMenu.makeMenu('dbm315','dbm314','2v2 Ladder','rules.asp?set=2v2+Ladder+Rules')
		oCMenu.makeMenu('dbm119','dbm2','Counter Strike','')
			oCMenu.makeMenu('dbm149','dbm119','2v2','rules.asp?set=Counter+Strike')
			oCMenu.makeMenu('dbm150','dbm119','5v5','rules.asp?set=Counter+Strike')
		oCMenu.makeMenu('dbm418','dbm2','Day of Defeat','')
			oCMenu.makeMenu('dbm419','dbm418','6v6','rules.asp?set=Day+of+Defeat+Rules')
		oCMenu.makeMenu('dbm395','dbm2','Delta Force: Black Hawk Down','')
			oCMenu.makeMenu('dbm396','dbm395','CTF 6v6','rules.asp?set=Black+Hawk+Down+%2D+CTF+6v6')
			oCMenu.makeMenu('dbm467','dbm395','TDM 8v8','rules.asp?set=Black+Hawk+Down+%2D+TDM+8v8')
			oCMenu.makeMenu('dbm423','dbm395','TKOTH 8v8','rules.asp?set=Black+Hawk+Down+%2D+TKOTH+8v8')
		oCMenu.makeMenu('dbm121','dbm2','Ghost Recon','')
			oCMenu.makeMenu('dbm397','dbm121','1v1','rules.asp?set=Ghost+Recon+1v1')
			oCMenu.makeMenu('dbm151','dbm121','Last Man Standing 4v4','rules.asp?set=Ghost+Recon+Last+Man+Standing')
		oCMenu.makeMenu('dbm120','dbm2','Global Operations','')
			oCMenu.makeMenu('dbm385','dbm120','4v4','rules.asp?set=Global+Operations+4v4')
			oCMenu.makeMenu('dbm152','dbm120','6v6','rules.asp?set=Global+Operations')
		oCMenu.makeMenu('dbm122','dbm2','Jedi Knight 2: Jedi Outcast','')
			oCMenu.makeMenu('dbm153','dbm122','Saber Duel 1v1','rules.asp?set=Jedi+Outcast+1v1+Rules')
		oCMenu.makeMenu('dbm123','dbm2','MechWarrior 4','')
			oCMenu.makeMenu('dbm246','dbm123','Mercenaries','rules.asp?set=MW4+Mercenaries')
			oCMenu.makeMenu('dbm282','dbm123','Mercenaries 2v2','rules.asp?set=MW4+Mercenaries+Double+Duel')
			oCMenu.makeMenu('dbm417','dbm123','Mercenaries Euro 2v2','rules.asp?set=MW4+Mercenaries+Euro+Double+Duel+')
			oCMenu.makeMenu('dbm348','dbm123','Mercs Duel','rules.asp?set=MW4+Mercs+Duel')
		oCMenu.makeMenu('dbm124','dbm2','Medal Of Honor','')
			oCMenu.makeMenu('dbm132','dbm124','Allied Assault Realism','')
				oCMenu.makeMenu('dbm159','dbm132','Objective 7v7','rules.asp?set=MoH%3AAA+Objective+Realism')
				oCMenu.makeMenu('dbm222','dbm132','Round Based Team Deathmatch 7v7','rules.asp?set=Medal+of+Honor%3A+Realism')
			oCMenu.makeMenu('dbm133','dbm124','Allied Assault Standard','')
				oCMenu.makeMenu('dbm295','dbm133','Invitational Composite 7v7 League','rules.asp?set=MoH%3AAA+Invitational+Composite+7v7')
				oCMenu.makeMenu('dbm465','dbm133','Invitational TDM 6v6 League','rules.asp?set=MoH%3AAA+Invitational+TDM+6v6+League')
				oCMenu.makeMenu('dbm160','dbm133','Objective 7v7','rules.asp?set=Medal+of+Honor%3A%3AAllied+Assault')
				oCMenu.makeMenu('dbm163','dbm133','Round Based Team Deathmatch 7v7','rules.asp?set=Medal+of+Honor%3A+Round+Based+TDM')
				oCMenu.makeMenu('dbm162','dbm133','Team Deathmatch 2v2','rules.asp?set=MoH%3AAA+2+Man+TDM')
				oCMenu.makeMenu('dbm161','dbm133','Team Deathmatch 6v6','rules.asp?set=Medal+of+Honor+TDM+Rules')
			oCMenu.makeMenu('dbm135','dbm124','Spearhead Realism','')
				oCMenu.makeMenu('dbm212','dbm135','Objective 7v7','viewladder.asp?ladder=MoH%3ASpearhead+OBJ+Realism+7v7')
				oCMenu.makeMenu('dbm213','dbm135','Round Based Team Deathmatch 7v7','rules.asp?set=Medal+of+Honor%3A+Spearhead+Round+Based+TDM+7v7+Realism+')
				oCMenu.makeMenu('dbm214','dbm135','Team Deathmatch 2v2','rules.asp?set=MoH%3ASpearhead+TDM+Realism+2v2')
				oCMenu.makeMenu('dbm215','dbm135','Team Deathmatch 6v6','rules.asp?set=Medal+of+Honor%3A+Spearhead+TDM+6v6+Realism+')
				oCMenu.makeMenu('dbm216','dbm135','Tug of War 7v7','rules.asp?set=Medal+of+Honor%3A+Spearhead+Tug+of+War+7v7+Realism+')
			oCMenu.makeMenu('dbm134','dbm124','Spearhead Standard','')
				oCMenu.makeMenu('dbm217','dbm134','Objective 7v7','rules.asp?set=Medal+of+Honor%3A+Spearhead+Objective')
				oCMenu.makeMenu('dbm218','dbm134','Round Based Team Deathmatch 7v7','rules.asp?set=Medal+of+Honor%3A+Spearhead+Round+Based+TDM+7v7')
				oCMenu.makeMenu('dbm219','dbm134','Team Deathmatch 2v2','rules.asp?set=Medal+of+Honor%3A+Spearhead+TDM+2v2')
				oCMenu.makeMenu('dbm220','dbm134','Team Deathmatch 6v6','rules.asp?set=Medal+of+Honor%3A+Spearhead+TDM+6v6')
				oCMenu.makeMenu('dbm221','dbm134','Tug of War 7v7','rules.asp?set=Medal+of+Honor%3A+Spearhead+Tug+of+War+7v7')
		oCMenu.makeMenu('dbm455','dbm2','PS2: Tribes Aerial Assault','')
			oCMenu.makeMenu('dbm456','dbm455','CTF 8v8','rules.asp?set=Tribes+AA+Capture+the+Flag+%28CTF%29')
		oCMenu.makeMenu('dbm223','dbm2','Rainbow 6: Raven Shield','')
			oCMenu.makeMenu('dbm361','dbm223','Europe','')
				oCMenu.makeMenu('dbm267','dbm361','Team Survival Europe 4v4','rules.asp?set=Rainbow+6%3A+Raven+Shield+4+Man+Team+Survival+Europe+Rules')
				oCMenu.makeMenu('dbm226','dbm361','Team Survival Europe 7v7','rules.asp?set=Rainbow+6%3A+Raven+Shield+Team+Survival+Europe')
			oCMenu.makeMenu('dbm363','dbm223','Leagues','')
				oCMenu.makeMenu('dbm380','dbm363','Composite Objective 5v5 (Coming Soon)','')
				oCMenu.makeMenu('dbm367','dbm363','Team Survival 5v5 League','rules.asp?set=Rainbow+Six%3A+Raven+Shield+Team+Survival+5v5+League')
			oCMenu.makeMenu('dbm365','dbm223','Objective','')
				oCMenu.makeMenu('dbm369','dbm365','Composite Objective 6v6','rules.asp?set=Rainbow+6%3A+Raven+Shield+6+Man+Composite+Objective')
			oCMenu.makeMenu('dbm364','dbm223','Team Survival','')
				oCMenu.makeMenu('dbm398','dbm364','2v2','rules.asp?set=Rainbow+6%3A+Raven+Shield+2+Man+Team+Survival+Rules')
				oCMenu.makeMenu('dbm258','dbm364','4v4','rules.asp?set=Rainbow+6%3A+Raven+Shield+4+Man+Team+Survival+Rules')
				oCMenu.makeMenu('dbm224','dbm364','7v7','rules.asp?set=Rainbow+6%3A+Raven+Shield+Team+Survival')
			oCMenu.makeMenu('dbm366','dbm223','Tournaments','')
				oCMenu.makeMenu('dbm317','dbm366','Operation Triple Threat (Closed)','rules.asp?set=Rainbow+6:+Raven+Shield+3v3+Tournament')
		oCMenu.makeMenu('dbm128','dbm2','Return To Castle Wolfenstein','')
			oCMenu.makeMenu('dbm388','dbm128','Custom','rules.asp?set=Return+to+Castle+Wolfenstein+%2D+Customs')
			oCMenu.makeMenu('dbm166','dbm128','League 6v6','rules.asp?set=Return+to+Castle+Wolfenstein+League')
			oCMenu.makeMenu('dbm165','dbm128','One Life To Live 7v7','rules.asp?set=One+Life+to+Live')
			oCMenu.makeMenu('dbm164','dbm128','StopWatch 5v5 + 6v6 + 7v7','rules.asp?set=Return+To+Castle+Wolfenstein')
		oCMenu.makeMenu('dbm125','dbm2','Soldier Of Fortune 2','')
			oCMenu.makeMenu('dbm167','dbm125','Capture The Flag 8v8','rules.asp?set=Soldier+of+Fortune+2+Capture+the+Flag')
			oCMenu.makeMenu('dbm168','dbm125','Elimination 2v2','rules.asp?set=Soldier+of+Fortune+2+Elimination+2v2')
			oCMenu.makeMenu('dbm169','dbm125','Infiltration 5v5','rules.asp?set=Soldier+of+Fortune+2+Infiltration')
			oCMenu.makeMenu('dbm379','dbm125','Infiltration League','rules.asp?set=Soldier+of+Fortune+2+League+Rules')
		oCMenu.makeMenu('dbm126','dbm2','Tribes','')
			oCMenu.makeMenu('dbm176','dbm126','Arena','')
				oCMenu.makeMenu('dbm179','dbm176','2v2','rules.asp?set=Tribes+Arena+2v2')
				oCMenu.makeMenu('dbm306','dbm176','4v4','rules.asp?set=Tribes+Arena+4v4')
			oCMenu.makeMenu('dbm173','dbm126','Capture The Flag','')
				oCMenu.makeMenu('dbm175','dbm173','10v10','rules.asp?set=Tribes+CTF+10v10')
				oCMenu.makeMenu('dbm228','dbm173','5v5','rules.asp?set=Tribes+CTF+5v5')
			oCMenu.makeMenu('dbm183','dbm126','Duel','')
				oCMenu.makeMenu('dbm184','dbm183','Duel Mod 1v1','rules.asp?set=Tribes+1+Duel')
		oCMenu.makeMenu('dbm127','dbm2','Tribes 2','')
			oCMenu.makeMenu('dbm242','dbm127','Arena','')
				oCMenu.makeMenu('dbm190','dbm242','2v2','rules.asp?set=Tribes+2+Arena+2v2')
				oCMenu.makeMenu('dbm189','dbm242','6v6','rules.asp?set=Tribes+2+Arena+6v6')
				oCMenu.makeMenu('dbm325','dbm242','Classic 2v2','rules.asp?set=Tribes+2+Classic+Arena+2v2')
				oCMenu.makeMenu('dbm283','dbm242','Classic 4v4','rules.asp?set=Tribes+2+Classic+Arena+4v4')
			oCMenu.makeMenu('dbm136','dbm127','Capture The Flag','')
				oCMenu.makeMenu('dbm185','dbm136','Capture The Flag','rules.asp?set=Capture+The+Flag+%28CTF%29')
			oCMenu.makeMenu('dbm137','dbm127','Duel','')
				oCMenu.makeMenu('dbm192','dbm137','Classic Duel Mod 1v1','rules.asp?set=Tribes+2+Classic+Duel')
				oCMenu.makeMenu('dbm191','dbm137','Duel Mod 1v1','rules.asp?set=Tribes+2+Duel')
			oCMenu.makeMenu('dbm138','dbm127','Oceania','')
				oCMenu.makeMenu('dbm197','dbm138','Classic Capture The Flag 12v12','rules.asp?set=Tribes+2+Oceania+Classic')
				oCMenu.makeMenu('dbm188','dbm138','Classic Duel Mod 1v1','rules.asp?set=Tribes+2+Oceania+Classic+Duel')
			oCMenu.makeMenu('dbm139','dbm127','Renegades','')
				oCMenu.makeMenu('dbm194','dbm139','CTF 12v12','rules.asp?set=Tribes+2+Renegades+CTF+12v12')
				oCMenu.makeMenu('dbm195','dbm139','Duel Mod 1v1','rules.asp?set=Tribes+2+Renegades+II+Duel')
				oCMenu.makeMenu('dbm345','dbm139','Elite Renegades 5v5','rules.asp?set=Tribes+2+Elite+Renegades2+5v5')
			oCMenu.makeMenu('dbm410','dbm127','Siege','')
				oCMenu.makeMenu('dbm411','dbm410','Siege','rules.asp?set=Tribes+2+Siege')
			oCMenu.makeMenu('dbm375','dbm127','Team Gauntlet','')
				oCMenu.makeMenu('dbm376','dbm375','5v5','rules.asp?set=Tribes+2+Team+Gauntlet+5v5')
				oCMenu.makeMenu('dbm377','dbm375','Classic 5v5','rules.asp?set=Tribes+2+Team+Gauntlet+5v5')
		oCMenu.makeMenu('dbm130','dbm2','Unreal Tournament 2003','')
			oCMenu.makeMenu('dbm203','dbm130','Deathmatch 1v1','rules.asp?set=UT+2003+DM')
		oCMenu.makeMenu('dbm129','dbm2','Urban Terror','')
			oCMenu.makeMenu('dbm199','dbm129','Capture The Flag 6v6','rules.asp?set=Urban+Terror+Capture+The+Flag')
			oCMenu.makeMenu('dbm344','dbm129','Team Survivor 2v2','rules.asp?set=2v2+Team+Survivor+Ladder')
			oCMenu.makeMenu('dbm201','dbm129','Team Survivor 6v6','rules.asp?set=Urban+Terror+Team+Survivor')
			oCMenu.makeMenu('dbm294','dbm129','Urban Terror CTF League','rules.asp?set=Urban+Terror+CTF+League')
		oCMenu.makeMenu('dbm332','dbm2','Vietcong','')
			oCMenu.makeMenu('dbm333','dbm332','8v8','rules.asp?set=Vietcong+Composite+8v8')
		oCMenu.makeMenu('dbm407','dbm2','Xbox Live','')
			oCMenu.makeMenu('dbm408','dbm407','Ghost Recon','')
				oCMenu.makeMenu('dbm409','dbm408','Coming Soon','')
// End DB Menus
 	
oCMenu.makeMenu('oper','','operations')
	oCMenu.makeMenu('files','oper','Downloads / Files','files')
	oCMenu.makeMenu('staff','oper','Staff','staff.asp')
	oCMenu.makeMenu('java_irc','oper','Java IRC','jirc/')
	oCMenu.makeMenu('winners','oper','Prize Winners','winners.asp')

	oCMenu.makeMenu('demos','oper','Demo Library', 'demos')
	oCMenu.makeMenu('Voting','oper','Voting Booth', 'ballot/')
	oCMenu.makeMenu('Contrib','oper','TWL Contributors', 'contributors.asp')<% If bSysAdmin Or bAnyLadderAdmin Or IsAnyLeagueAdmin() Then %>
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
<% End If %>oCMenu.construct()	
</SCRIPT>