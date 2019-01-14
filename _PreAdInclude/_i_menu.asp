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
oCMenu.level[1].width=150
oCMenu.level[1].height=22
oCMenu.level[1].regClass="clLevel1"
oCMenu.level[1].overClass="clLevel1over"
oCMenu.level[1].borderX=1
oCMenu.level[1].borderY=1
oCMenu.level[1].align="left" 
oCMenu.level[1].arrow="images/tri.gif"
//oCMenu.level[1].arrow=0
oCMenu.level[1].arrowWidth=10
oCMenu.level[1].arrowHeight=10
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
					Response.Write "oCMenu.makeMenu('myteams" & intNumber & "','myteams','" & Server.HTMLEncode(Replace(Replace(Replace(ors.fields("TeamTag").value, "/", "\/"), "\", "\\"), "'", "\'")) & "','viewteam.asp?team=" & Server.URLEncode(ors.fields("TeamName").value) & "')" & vbcrlf
'					Response.Write "oCMenu.makeMenu('myteams" & intNumber & "','myteams','" & Server.HTMLEncode(Replace(ors.fields("TeamTag").value, "'", "\'")) & "','viewteam.asp?team=" & Server.URLEncode(ors.fields("TeamName").value) & "')" & vbcrlf
					strMenuTeamName = oRS.Fields("TeamName").Value
			End If
			oRS.MoveNext
		Loop
	Else
		Response.Write "oCMenu.makeMenu('myteams" & intNumber & "','myteams','No teams found.')" & vbcrlf
	End If
	oRS.NextRecordSet

	%>
<% Else %>
oCMenu.makeMenu('my_twl','','my twl', 'default.asp')
	oCMenu.makeMenu('home','my_twl','Home', 'default.asp')
	oCMenu.makeMenu('login','my_twl','Login', '', '', '','', '', '', '', '', '', '', '', 'javascript:popup("/login.asp?url=' + this.location.href + '", "login", 175, 300, "no");')
	oCMenu.makeMenu('forgot_password','my_twl','Forgot Password','forgotpassword.asp')
	oCMenu.makeMenu('activate','my_twl','Deactivated Account?','activate.asp')
	oCMenu.makeMenu('register','my_twl','Register','addplayer.asp')
<% End If %>

oCMenu.makeMenu('forums','','forums','forums/')
	oCMenu.makeMenu('forumindex','forums','Forum Index','forums/')
	oCMenu.makeMenu('forum9','forums','General Forums','forums/forumdisplay.asp?forumid=9')
	oCMenu.makeMenu('forum19','forums','Site Support','forums/forumdisplay.asp?forumid=19')
	oCMenu.makeMenu('forum11','forums','Recruiting','forums/forumdisplay.asp?forumid=11')
	oCMenu.makeMenu('gameForums','forums','Game Specific','forums/default.asp#Category1')
	oCMenu.makeMenu('forum25','gameForums','America\'s Army','forums/forumdisplay.asp?forumid=25')
	oCMenu.makeMenu('forum27','gameForums','Battlefield 1942','forums/forumdisplay.asp?forumid=27')
	oCMenu.makeMenu('forum12','gameForums','Castle Wolfenstein','forums/forumdisplay.asp?forumid=12')
	oCMenu.makeMenu('forum14','gameForums','Counter Strike','forums/forumdisplay.asp?forumid=14')
	oCMenu.makeMenu('forum24','gameForums','Ghost Recon','forums/forumdisplay.asp?forumid=24')
	oCMenu.makeMenu('forum21','gameForums','Global Operations','forums/forumdisplay.asp?forumid=21')
	oCMenu.makeMenu('forum18','gameForums','Jedi Outcast','forums/forumdisplay.asp?forumid=18')
	oCMenu.makeMenu('forum2','gameForums','Mechwarrior 4','forums/forumdisplay.asp?forumid=2')
	oCMenu.makeMenu('forum13','gameForums','Medal of Honor','forums/forumdisplay.asp?forumid=13')
	oCMenu.makeMenu('forum23','gameForums','Soldier Of Fortune 2','forums/forumdisplay.asp?forumid=23')
	oCMenu.makeMenu('forum34','gameForums','Rainbow Six: Raven Shield','forums/forumdisplay.asp?forumid=34')
	oCMenu.makeMenu('forum1','gameForums','Tribes','forums/forumdisplay.asp?forumid=1')
	oCMenu.makeMenu('forum17','gameForums','Tribes Renegades','forums/forumdisplay.asp?forumid=17')
	oCMenu.makeMenu('forum4','gameForums','Tribes 2 CTF','forums/forumdisplay.asp?forumid=4')
	oCMenu.makeMenu('forum29','gameForums','Tribes 2 TR2','forums/forumdisplay.asp?forumid=29')
	oCMenu.makeMenu('forum31','gameForums','Tribes 2 Map Council','forums/forumdisplay.asp?forumid=31')
	oCMenu.makeMenu('forum10','gameForums','UT2003','forums/forumdisplay.asp?forumid=10')
	oCMenu.makeMenu('forum26','gameForums','Warcraft 3','forums/forumdisplay.asp?forumid=26')
	oCMenu.makeMenu('forum32','gameForums','Urban Terror','forums/forumdisplay.asp?forumid=32')
<% If bSysAdmin Or bAnyLadderAdmin Then %>
	oCMenu.makeMenu('forum3','forums','TWL Staff','forums/forumdisplay.asp?forumid=3')
	oCMenu.makeMenu('forum30','forums','League Admin','forums/forumdisplay.asp?forumid=30')
	oCMenu.makeMenu('forum5','forums','Development','forums/forumdisplay.asp?forumid=5')
	oCMenu.makeMenu('forum22','forums','Event Management','forums/forumdisplay.asp?forumid=22')
	<% If bSysAdmin Then %>
	oCMenu.makeMenu('forum8','forums','SysAdmin Forum','forums/forumdisplay.asp?forumid=8')
	<% End If %>
<% End If %>

oCMenu.makeMenu('comp','','competition','ladderlist.asp')
	oCMenu.makeMenu('full','comp','Full Competition List','ladderlist.asp')
	oCMenu.makeMenu('aaladder','comp','America\'s Army')
		oCMenu.makeMenu('aa6v6League','aaladder','6v6 League','viewleague.asp?league=America%27s+Army+6+v+6')
		oCMenu.makeMenu('aa22Ladder','aaladder','2v2 Team Combat','viewladder.asp?ladder=America%27s+Army+2v2')
		oCMenu.makeMenu('aa44Ladder','aaladder','4v4 Team Objective','viewladder.asp?ladder=America%27s+Army+4v4')
		oCMenu.makeMenu('aatcLadder','aaladder','Team Combat','viewladder.asp?Ladder=America%27s+Army+Team+Combat')
		oCMenu.makeMenu('aaooLadder','aaladder','Team Objective','viewladder.asp?Ladder=America%27s+Army+Objective')
		oCMenu.makeMenu('aaTOEuroLadder','aaladder','Team Objective - Euro','viewladder.asp?Ladder=America%27s+Army+Team+Objective+%2D+Europe')
		oCMenu.makeMenu('aa4v4EuroLadder','aaladder','4v4 Team Objective - Euro','viewladder.asp?Ladder=America%27s+Army+4v4+%2D+Europe')	oCMenu.makeMenu('bfladder','comp','Battlefield 1942 ')
		oCMenu.makeMenu('bf1League','bfladder','Battlefied 1942 League','viewleague.asp?league=Battlefield+1942')
		oCMenu.makeMenu('bf1Ladder','bfladder','Conquest','viewladder.asp?Ladder=Battlefield+1942+Conquest')
		oCMenu.makeMenu('bf2Ladder','bfladder','Conquest - 8 Man','viewladder.asp?Ladder=Battlefield+1942+Conquest+%2D+8+man')
		oCMenu.makeMenu('bf3Ladder','bfladder','CTF','viewladder.asp?Ladder=Battlefield+1942+CTF')
	
	oCMenu.makeMenu('csladder','comp','Counter-Strike ')
		oCMenu.makeMenu('cs2v2','csladder','Counter Strike 2v2','viewladder.asp?Ladder=CounterStrike+2v2')
		oCMenu.makeMenu('csopen','csladder','Counter Strike Open','viewladder.asp?Ladder=CounterStrike+Open')

	oCMenu.makeMenu('game12','comp','Global Operations')
		oCMenu.makeMenu('gops','game12','Global Operations','viewladder.asp?Ladder=Global+Operations')
		oCMenu.makeMenu('gopseuro','game12','Global Operations - Europe','viewladder.asp?Ladder=Global+Operations+%2D+Europe')

	oCMenu.makeMenu('GameGhostRecon','comp','Ghost Recon')
		oCMenu.makeMenu('GGRLMS','GameGhostRecon','Last Man Standing','viewladder.asp?ladder=Ghost+Recon+Last+Man+Standing')

	oCMenu.makeMenu('game11','comp','JK2: Jedi Outcast')
		oCMenu.makeMenu('jk2ctfff','game11','CTF Full Force','viewladder.asp?Ladder=Jedi+Outcast+CTF+Full+Force')
		oCMenu.makeMenu('jk2tdmff','game11','TDM Full Force','viewladder.asp?Ladder=Jedi+Outcast+TDM+Full+Force')
		oCMenu.makeMenu('jk2tdmsaber','game11','TDM Sabers','viewladder.asp?Ladder=Jedi+Outcast+TDM+Sabers')
		oCMenu.makeMenu('jk21v1ffs','game11','1v1 Full Force Sabers','viewPlayerladder.asp?ladder=Jedi+Outcast+1v1+Full+Force+Sabers')
//		oCMenu.makeMenu('jk21v1pacrim','game11','Oceania Saber Duel','viewPlayerladder.asp?ladder=Jedi+Outcast+Oceania+Sabel+Duel')
		oCMenu.makeMenu('jk21v1nfs','game11','1v1 No Force Sabers','viewPlayerladder.asp?ladder=Jedi+Outcast+1v1+No+Force+Sabers')

	oCMenu.makeMenu('game1','comp','MW4 ')
		oCMenu.makeMenu('ta','game1','Team Attrition','viewladder.asp?Ladder=MW4+TA')
		oCMenu.makeMenu('td','game1','Team Destruction','viewladder.asp?Ladder=MW4+TD')
		oCMenu.makeMenu('mw4dd','game1','BK Double Duel','viewladder.asp?Ladder=MW4+Black+Knight+Double+Duel')

	oCMenu.makeMenu('mohladder','comp','Medal Of Honor')
		oCMenu.makeMenu('mohladderset2','mohladder','Standard')
			oCMenu.makeMenu('mohladder1','mohladderset2','MoH: Allied Assault','viewladder.asp?Ladder=MoH%3AAA')
			oCMenu.makeMenu('mohladder2','mohladderset2','MoH: AA Team Deathmatch','viewladder.asp?Ladder=MoH%3AAA+Team+Deathmatch')
			oCMenu.makeMenu('mohladder3','mohladderset2','MoH: AA 2 Man TDM','viewladder.asp?Ladder=MoH%3AAA+2+Man+TDM')
			oCMenu.makeMenu('mohladder5','mohladderset2','MoH: AA Round Based TDM ','viewladder.asp?Ladder=MoH%3AAA+Round+Based+TDM')
		oCMenu.makeMenu('mohladderset1','mohladder','Realism')
			oCMenu.makeMenu('mohladder4','mohladderset1','MoH: AA Objective','viewladder.asp?ladder=MoH%3AAA+Objective+Realism')
			oCMenu.makeMenu('mohladder6','mohladderset1','MoH: AA Round Based TDM','viewladder.asp?ladder=MoH%3AAA+Round+Based+TDM+Realism')
		oCMenu.makeMenu('mohladderset3','mohladder','Spearhead')
			oCMenu.makeMenu('mohladder7','mohladderset3','MoH:Spearhead OBJ 7-Man','viewladder.asp?ladder=MoH%3ASpearhead+OBJ+7%2Dman')
			oCMenu.makeMenu('mohladder8','mohladderset3','MoH:Spearhead RB TDM 7-Man','viewladder.asp?ladder=MoH%3ASpearhead+RB+TDM+7%2Dman')
			oCMenu.makeMenu('mohladder9','mohladderset3','MoH:Spearhead TDM 2-Man','viewladder.asp?ladder=MoH%3ASpearhead+TDM+2%2Dman')
			oCMenu.makeMenu('mohladder10','mohladderset3','MoH:Spearhead TDM 6-Man','viewladder.asp?ladder=MoH%3ASpearhead+TDM+6%2Dman')
		
	oCMenu.makeMenu('sofladder','comp','Soldier of Fortune 2')
		oCMenu.makeMenu('sofctf','sofladder','Capture the Flag','viewladder.asp?ladder=Soldier+of+Fortune+2+CTF')
		oCMenu.makeMenu('sofelim2v2','sofladder','Elimination 2v2','viewladder.asp?ladder=Soldier+of+Fortune+2+Elimination+2v2')
		oCMenu.makeMenu('sofeast','sofladder','Infiltration - East','viewladder.asp?ladder=Soldier+of+Fortune+2+East')
		oCMenu.makeMenu('sofwest','sofladder','Infiltration - West','viewladder.asp?ladder=Soldier+of+Fortune+2+West')

	oCMenu.makeMenu('game2','comp','Tribes ')
		oCMenu.makeMenu('arena','game2','Arena','viewladder.asp?Ladder=Tribes+Arena')
		oCMenu.makeMenu('arena2v2','game2','2v2 Arena','viewladder.asp?Ladder=Tribes+Arena+2v2')
		oCMenu.makeMenu('ctf','game2','CTF','viewladder.asp?Ladder=Tribes+CTF')
		oCMenu.makeMenu('t1duel','game2','Duel','viewplayerladder.asp?Ladder=Tribes+Duel')
		oCMenu.makeMenu('t1pac','game2','Oceania CTF','viewladder.asp?ladder=Tribes+Oceania+CTF')
		oCMenu.makeMenu('t1pacduel','game2','Oceania Duel','viewPlayerladder.asp?ladder=Tribes+Oceania+Duel')

	oCMenu.makeMenu('game3','comp','Tribes 2 ')
		oCMenu.makeMenu('t2ctfsub','game3','Capture the Flag','')
			oCMenu.makeMenu('7man','t2ctfsub','7-Man CTF','viewladder.asp?ladder=Tribes+2+7%2DMan')
			oCMenu.makeMenu('12man','t2ctfsub','12-Man CTF','viewladder.asp?ladder=Tribes+2+12%2DMan')
			oCMenu.makeMenu('7manbasecl','t2ctfsub','7-Man Classic','viewladder.asp?ladder=Tribes+2+7%2DMan+Classic')
			oCMenu.makeMenu('t2basecl','t2ctfsub','Classic','viewladder.asp?ladder=Tribes+2+Classic')
			oCMenu.makeMenu('t2basev2','t2ctfsub','Version 2','viewladder.asp?ladder=Tribes+2+Version+2')

		oCMenu.makeMenu('t2pacrimsub','game3','Oceania','')
			oCMenu.makeMenu('pacrimduel','t2pacrimsub','Oceania Duel','viewplayerladder.asp?ladder=Tribes+2+Oceania+Duel')
			oCMenu.makeMenu('pacrimclassicduel','t2pacrimsub','Oceania Classic Duel','viewplayerladder.asp?ladder=Tribes+2+Oceania+Classic+Duel')
			oCMenu.makeMenu('pacrimclassic','t2pacrimsub','Oceania Classic CTF','viewladder.asp?ladder=Tribes+2+Oceania+Classic+CTF')
			oCMenu.makeMenu('pacrimt2tr2','t2pacrimsub','Oceania TR2','viewladder.asp?ladder=Tribes+2+Oceania+Team+Rabbit+2')

		oCMenu.makeMenu('t2renegadesub','game3','Renegades','')
			oCMenu.makeMenu('t2rene','t2renegadesub','Renegades','viewladder.asp?ladder=Tribes+2+Renegades')
			oCMenu.makeMenu('t2reneduel','t2renegadesub','Renegades Duel','viewPlayerladder.asp?ladder=Tribes+2+Renegades+Duel')
			oCMenu.makeMenu('t2renearmorduel','t2renegadesub','Renegades Armor Duel','viewPlayerladder.asp?ladder=Tribes+2+Renegades+Armor+Duel')

		oCMenu.makeMenu('t2ordersub','game3','All Others','')
			oCMenu.makeMenu('t2arena','t2ordersub','Arena','viewladder.asp?ladder=Tribes+2+Arena')
			oCMenu.makeMenu('t2arena2v2','t2ordersub','2v2 Arena','viewladder.asp?ladder=Tribes+2+Arena+2v2')
			oCMenu.makeMenu('t2Duel','t2ordersub','Duel','viewplayerladder.asp?ladder=Tribes+2+Duel')
			oCMenu.makeMenu('t2Duelclassic','t2ordersub','Classic Duel','viewplayerladder.asp?ladder=Tribes+2+Classic+Duel')
			oCMenu.makeMenu('t2tr2leagueinvite','t2ordersub','TR2 Invitational League','viewleague.asp?league=Tribes+2+%2D+Team+Rabbit+2+Invitational')
			oCMenu.makeMenu('t2tr2leagueopen','t2ordersub','TR2 Open League','viewleague.asp?league=Tribes+2+%2D+Team+Rabbit+2+Open')

	oCMenu.makeMenu('game5','comp','Castle Wolfenstein')
		oCMenu.makeMenu('rtcwl1','game5','7v7 Alpha League','viewleague.asp?league=RTCW+7v7+Alpha')
		oCMenu.makeMenu('rtcwl1p','game5','7v7 Alpha League Playoffs','tournament/viewBracket.asp?tournament=Alpha+RtCW+Playoffs')
		oCMenu.makeMenu('rtcwl2','game5','7v7 Beta League','viewleague.asp?league=RTCW+7v7+Beta')
		oCMenu.makeMenu('rtcwl2p','game5','7v7 Beta League Playoffs','tournament/viewBracket.asp?tournament=Beta+RtCW+Playoffs')
		oCMenu.makeMenu('rtcw','game5','RTCW','viewladder.asp?ladder=RTCW')
		oCMenu.makeMenu('rtcw5','game5','RTCW 5v5','viewladder.asp?ladder=RTCW+5v5')
		oCMenu.makeMenu('rtcwoltl','game5','RTCW One Life To Live','viewladder.asp?ladder=RTCW+%2D+One+Life+to+Live')

	oCMenu.makeMenu('game6','comp','Urban Terror')
		oCMenu.makeMenu('urbanterror','game6','Urban Terror Team Survivor','viewladder.asp?ladder=Urban+Terror+TS')
		oCMenu.makeMenu('utrCTF','game6','Urban Terror CTF','viewladder.asp?ladder=Urban+Terror+CTF')

	oCMenu.makeMenu('game20','comp','Unreal Tournament 2003')
		oCMenu.makeMenu('ut20031v1','game20','1v1 Deathmatch Ladder','viewPlayerladder.asp?ladder=Unreal+Tournament+2003+%2D+1v1+DM')
		oCMenu.makeMenu('ut20034manTDM','game20','4 Man TDM Ladder','viewladder.asp?ladder=Unreal+Tournament+2003+%2D+4%2Dman+TDM')
		oCMenu.makeMenu('ut20035manbr','game20','5 Man Bombing Run Ladder','viewladder.asp?ladder=Unreal+Tournament+2003+%2D+5%2Dman+Bombing+Run')
		oCMenu.makeMenu('ut20035manctf','game20','5 Man CTF Ladder','viewladder.asp?ladder=Unreal+Tournament+2003+%2D+5%2Dman+CTF')
		oCMenu.makeMenu('ut20035manigctf','game20','5 Man Instagib CTF Ladder','viewladder.asp?ladder=Unreal+Tournament+2003+%2D+5%2Dman+Instagib+CTF')
		oCMenu.makeMenu('ut20035manigbr','game20','5 Man Instagib Bombing Run Ladder','viewladder.asp?ladder=Unreal+Tournament+2003+%2D+5%2Dman+Instagib+Bombing+Run')

	oCMenu.makeMenu('wcladder','comp','Warcraft 3')
		oCMenu.makeMenu('1v1wcladder','wcladder','1v1','viewplayerladder.asp?ladder=Warcraft+3+1v1')
		oCMenu.makeMenu('2v2wcladder','wcladder','2v2','viewladder.asp?ladder=Warcraft+3+2v2')
		oCMenu.makeMenu('2v2wcladderpacrim','wcladder','Oceania 2v2','viewladder.asp?ladder=Warcraft+3+Oceania+2v2')

	oCMenu.makeMenu('rules','','rules')
//		oCMenu.makeMenu('official','rules','Official Game Rules','rules.asp?set=Official+Game+Rules')

		oCMenu.makeMenu('AARules','rules','America\'s Army')
			oCMenu.makeMenu('AALeagueRules','AARules','6v6 League Rules', 'rules.asp?set=America%27s+Army+6+v+6+League')
			oCMenu.makeMenu('AA22Rules','AARules','2v2 Team Combat Rules', 'rules.asp?set=America%27s+Army+2+v+2')
			oCMenu.makeMenu('AA44Rules','AARules','4v4 Team Objective Rules', 'rules.asp?set=America%27s+Army+4+v+4')
			oCMenu.makeMenu('AATCRules','AARules','Team Combat Rules', 'rules.asp?set=America%27s+Army+Team+Combat')
			oCMenu.makeMenu('AAORules','AARules','Objective Rules', 'rules.asp?set=America%27s+Army+Objective')
			oCMenu.makeMenu('AAEuroORules','AARules','Euro Objective Rules', 'rules.asp?set=America%27s+Army+Team+Combat+%2D+Europe')
			oCMenu.makeMenu('AAEuro4v4Rules','AARules','Euro 4V4 Rules', 'rules.asp?set=America%27s+Army+4v4+%2D+Europe')

		oCMenu.makeMenu('BFRules','rules','Battlefield 1942')
			oCMenu.makeMenu('BFCQRules','BFRules','Conquest Rules', 'rules.asp?set=Battlefield+1942+Conquest')
			oCMenu.makeMenu('BFCQ8Rules','BFRules','Conquest 8 Man Rules', 'rules.asp?set=Battlefield+1942+Conquest+%2D+8+man')

		oCMenu.makeMenu('CSRules','rules','Counter Strike ')
			oCMenu.makeMenu('CS2Rules','CSRules','CS 2v2 Rules', 'rules.asp?set=Counter+Strike')
			oCMenu.makeMenu('CSOpenRules','CSRules','CS Open Rules', 'rules.asp?set=Counter+Strike')
		
		oCMenu.makeMenu('GORules','rules','Global Operations')
			oCMenu.makeMenu('GORules1','GORules','Global Operations', 'rules.asp?set=Global+Operations')

		oCMenu.makeMenu('GRRules','rules','Ghost Recon')
			oCMenu.makeMenu('GRLMSRules','GRRules','Last Man Standing Rules', 'rules.asp?set=Ghost+Recon+Last+Man+Standing')

		oCMenu.makeMenu('JORules','rules','JK2: Jedi Outcast')
			oCMenu.makeMenu('JO1v1Rules','JORules','1v1 Rules', 'rules.asp?set=Jedi+Outcast+1v1+Rules')
			oCMenu.makeMenu('JOCTFRules','JORules','CTF Rules', 'rules.asp?set=Jedi+Outcast+CTF+Rules')
			oCMenu.makeMenu('JOTDMRules','JORules','TDM Rules', 'rules.asp?set=Jedi+Outcast+TDM+Rules')
		
		oCMenu.makeMenu('MW4Rules','rules','MechWarrior 4 ')
			oCMenu.makeMenu('TARules','MW4Rules','MW4 TA Rules', 'rules.asp?set=MW4+Team+Attrition')
			oCMenu.makeMenu('TDRules','MW4Rules','MW4 TD Rules','rules.asp?set=MW4+Team+Destruction')
			oCMenu.makeMenu('mw4bkddrules','MW4Rules','MW4 BK Double Duel','rules.asp?set=MW4+Black+Knight+Double+Duel')

		oCMenu.makeMenu('mohladderrules','rules','Medal Of Honor')
			oCMenu.makeMenu('mohladder1rules','mohladderrules','MoH: Allied Assault','rules.asp?set=Medal+of+Honor%3A%3AAllied+Assault')
			oCMenu.makeMenu('mohladder2rules','mohladderrules','MoH: Team Deathmatch','rules.asp?set=Medal+of+Honor+TDM+Rules')
			oCMenu.makeMenu('mohladder3rules','mohladderrules','MoH: 2 Man TDM','rules.asp?set=MoH%3AAA+2+Man+TDM')
			oCMenu.makeMenu('mohladder4rules','mohladderrules','MoH: Realism','rules.asp?set=Medal+of+Honor%3A+Realism')
			oCMenu.makeMenu('mohladder5rules','mohladderrules','MoH: Round Based TDM','rules.asp?set=Medal+of+Honor%3A+Round+Based+TDM')

		oCMenu.makeMenu('SOF2Rules','rules','Soldier of Fortune 2')
			oCMenu.makeMenu('sof2ctf','SOF2Rules','Capture the Flag', 'rules.asp?set=Soldier+of+Fortune+2+Capture+the+Flag')
			oCMenu.makeMenu('sof2elim2v2','SOF2Rules','Elimination 2v2', 'rules.asp?set=Soldier+of+Fortune+2+Elimination+2v2')
			oCMenu.makeMenu('sof2inf','SOF2Rules','Infiltration', 'rules.asp?set=Soldier+of+Fortune+2+Infiltration')
		
		oCMenu.makeMenu('T1Rules','rules','Tribes ')
			oCMenu.makeMenu('TribesRules','T1Rules','CTF Rules','rules.asp?set=Tribes+Capture+the+Flag%28CTF%29+Rules')
			oCMenu.makeMenu('TribesArenaRules','T1Rules','Arena Rules','rules.asp?set=Tribes+Arena+Rules')
			oCMenu.makeMenu('TribesArena2v2Rules','T1Rules','2v2 Arena Rules','rules.asp?set=Tribes+Arena+2v2')
			oCMenu.makeMenu('TribesDuelRules','T1Rules','Duel Rules','rules.asp?set=Tribes+1+Duel')
			oCMenu.makeMenu('TribesPRDuelRules','T1Rules','Oceania Duel Rules','rules.asp?set=Tribes+Oceania+Duel')
			oCMenu.makeMenu('Tribes1PacRules','T1Rules','Oceania CTF Rules','rules.asp?set=Tribes+Oceania+CTF')

		oCMenu.makeMenu('T2Rules','rules','Tribes 2 ')
			oCMenu.makeMenu('t2ctfsubrules','T2Rules','Capture the Flag','')
				oCMenu.makeMenu('Tribes2Rules','t2ctfsubrules','CTF Rules','rules.asp?set=Capture+The+Flag+%28CTF%29')
				oCMenu.makeMenu('Tribes2V2Rules','t2ctfsubrules','Version 2 CTF Rules','rules.asp?set=Tribes+2+Version+2')
	
			oCMenu.makeMenu('t2pacrimsubrules','T2Rules','Oceania','')
				oCMenu.makeMenu('Tribes2PacDuelRules','t2pacrimsubrules','Oceania Duel','rules.asp?set=Tribes+2+Duel')
				oCMenu.makeMenu('Tribes2PacClassicDuelRules','t2pacrimsubrules','Oceania Classic Duel','rules.asp?set=Tribes+2+Oceania+Classic+Duel')
				oCMenu.makeMenu('Tribes2PacClassicRules','t2pacrimsubrules','Oceania Classic CTF','rules.asp?set=Tribes+2+Oceania+Classic')
				oCMenu.makeMenu('Tribes2PacTR2Rules','t2pacrimsubrules','Oceania Team Rabbit 2 Rules','rules.asp?set=Tribes+2+Oceania+Team+Rabbit+2')
			
			oCMenu.makeMenu('t2renegadesubrules','T2Rules','Renegades','')
				oCMenu.makeMenu('Tribes2Ren','t2renegadesubrules','Renegades II','rules.asp?set=Tribes+2+Renegades+II+10%2DMan')
				oCMenu.makeMenu('Tribes2RenDuel','t2renegadesubrules','Renegades II Duel','rules.asp?set=Tribes+2+Renegades+II+Duel')
				oCMenu.makeMenu('Tribes2RenDuelArmor','t2renegadesubrules','Ren II Armor Duel','rules.asp?set=Tribes+2+Renegades+II+Armor+Duel')
	
			oCMenu.makeMenu('t2ordersubrules','T2Rules','All Others','')
				oCMenu.makeMenu('Tribes2ArenaRules','t2ordersubrules','Arena Rules','rules.asp?set=Tribes+2+Arena')
				oCMenu.makeMenu('Tribes2Arena2manRules','t2ordersubrules','Arena 2v2 Rules','rules.asp?set=Tribes+2+Arena+2v2')
				oCMenu.makeMenu('Tribes2DuelRules','t2ordersubrules','Duel Ladder Rules','rules.asp?set=Tribes+2+Duel')
				oCMenu.makeMenu('Tribes2DuelClassicRules','t2ordersubrules','Classic Duel Rules','rules.asp?set=Tribes+2+Classic+Duel')
				oCMenu.makeMenu('Tribes2TR2LeagueRules','t2ordersubrules','Team Rabbit 2 League Rules','rules.asp?set=Tribes+2+Team+Rabbit+2+League')

		oCMenu.makeMenu('urbanterrorrules','rules','Urban Terror')
			oCMenu.makeMenu('utrCTFRules','urbanterrorrules','Urban Terror CTF','rules.asp?set=Urban+Terror+Capture+The+Flag')
			oCMenu.makeMenu('utrTSRules','urbanterrorrules','Urban Terror Team Survivor','rules.asp?set=Urban+Terror+Team+Survivor')
			oCMenu.makeMenu('utrgeninfoRules','urbanterrorrules','General Urban Terror FAQ','rules.asp?set=General+Urban+Terror+FAQ')

		oCMenu.makeMenu('wc3rules','rules','Warcraft 3')
			oCMenu.makeMenu('wc31v1Rules','wc3rules','1v1 Rules','rules.asp?set=Warcraft+3+1v1')
			oCMenu.makeMenu('wc3CTFRules','wc3rules','2v2 Rules','rules.asp?set=Warcraft+3+2v2')
			oCMenu.makeMenu('wc3CTFRulespacrim','wc3rules','Oceania 2v2 Rules','rules.asp?set=WarCraft+3+Oceania+2v2')

		oCMenu.makeMenu('ut2k3rules','rules','Unreal Tournament 2003')
			oCMenu.makeMenu('ut2k3brrules','ut2k3rules','Bombing Run','rules.asp?set=UT+2003+Bombing+Run')
			oCMenu.makeMenu('ut2k3dmrules','ut2k3rules','Death Match','rules.asp?set=UT+2003+DM')
			oCMenu.makeMenu('ut2k3tdmrules','ut2k3rules','Team Death Match','rules.asp?set=UT+2003+TDM')
			oCMenu.makeMenu('ut2k35manctf','ut2k3rules','5 Man CTF','rules.asp?set=UT+2003+5%2DMan+CTF')
			oCMenu.makeMenu('ut2k35manigctf','ut2k3rules','5 Man Instagib CTF','rules.asp?set=UT+2003+Instagib+CTF')
			oCMenu.makeMenu('ut2k35manigbr','ut2k3rules','5 Man Instagib Bombing Run','rules.asp?set=UT+2003+Instagib+Bombing+Run')
//			oCMenu.makeMenu('ut2k3brleagrules','ut2k3rules','Bombing Run League','rules.asp?set=UT+2003+BR+League')

		oCMenu.makeMenu('WolfRules','rules','RTCW ')
			oCMenu.makeMenu('RTCWRules','WolfRules','RTCW Rules','rules.asp?set=Return+To+Castle+Wolfenstein')
			oCMenu.makeMenu('OLTLRTCWRules','WolfRules','RTCW One Life To Live Rules','rules.asp?set=One+Life+to+Live')
			oCMenu.makeMenu('RTCWLeagueRules','WolfRules','RTCW League Rules','rules.asp?set=Return+to+Castle+Wolfenstein+League')
 	
oCMenu.makeMenu('oper','','operations')
	oCMenu.makeMenu('files','oper','Downloads / Files','files')
	oCMenu.makeMenu('staff','oper','Staff','staff.asp')
	oCMenu.makeMenu('maps','oper','Map Guide','mapguide.asp')
	oCMenu.makeMenu('java_irc','oper','Java IRC','jirc/')
	oCMenu.makeMenu('winners','oper','Prize Winners','winners.asp')

	oCMenu.makeMenu('demos','oper','Demo Library', 'demos')
	oCMenu.makeMenu('Voting','oper','Voting Booth', 'ballot/')
	oCMenu.makeMenu('Contrib','oper','TWL Contributors', 'contributors.asp')<% If bSysAdmin Or bAnyLadderAdmin Then %>
	oCMenu.makeMenu('admin','','admin')
		oCMenu.makeMenu('admain','admin','Menu','adminmenu.asp')
		oCMenu.makeMenu('adopsnew','admin','News','newsdesk.asp')
		oCMenu.makeMenu('teamLadMenu','admin','Team Ladders ','')
		
		oCMenu.makeMenu('adopsmat','teamLadMenu','Match','adminops.asp?aType=Match')
		oCMenu.makeMenu('adopsfor','teamLadMenu','Forfeit','adminops.asp?aType=Forfeit')
		oCMenu.makeMenu('adopshis','teamLadMenu','History','adminops.asp?aType=History')
		oCMenu.makeMenu('adopslad','teamLadMenu','Ladder','adminops.asp?aType=Ladder')
		oCMenu.makeMenu('adopsrank','teamLadMenu','Rank','adminops.asp?aType=Rank')

		oCMenu.makeMenu('playerLadMenu','admin','Player Ladders ','')
			oCMenu.makeMenu('adopspmat','playerLadMenu','Match','adminops.asp?aType=PMatch')
			oCMenu.makeMenu('adopspfor','playerLadMenu','Forfeit','adminops.asp?aType=PForfeit')
			oCMenu.makeMenu('adopsplad','playerLadMenu','Ladder Admin','adminops.asp?aType=PLadder')
			oCMenu.makeMenu('adopsplrank','playerLadMenu','Player Rank','editplayerrank.asp')

		oCMenu.makeMenu('helprules','admin','Help/Rules','help/admin')

		oCMenu.makeMenu('reports','admin','Reports ','')
			oCMenu.makeMenu('plfor','reports','Player Forfeit Report','reports/playerforfietreport.asp')
			oCMenu.makeMenu('plrr','reports','Player Roster Report','reports/playerrosterreport.asp')
			oCMenu.makeMenu('rr','reports','Roster Report','reports/rosterreport.asp')
			oCMenu.makeMenu('act','reports','Ladder Activity','reports/activity.asp')

	<% If bSysAdmin then %>
		oCMenu.makeMenu('mm','admin','Mass Mail','massmail.asp')
		oCMenu.makeMenu('laddermenu','admin','Ladders ','')
			oCMenu.makeMenu('laAdmins','laddermenu','Assign Admins','assignadmin.asp')
			oCMenu.makeMenu('ladmatchoptions','laddermenu','Match Options','ladder/ladderoptions.asp')
			oCMenu.makeMenu('addladder','laddermenu','Add Ladder','addladder.asp')
		oCMenu.makeMenu('league','admin','League Admin','leagueadmin.asp')
		oCMenu.makeMenu('leagueaa','admin','League Assign Admin','leagueassignadmin.asp')
		oCMenu.makeMenu('forum','admin','Forum','forums/admin')
		oCMenu.makeMenu('votingadmin','admin','Voting Admin ','')
			oCMenu.makeMenu('addballot','votingadmin','Add Ballot','ballot/addballot.asp')
			oCMenu.makeMenu('actballot','votingadmin','Activate Ballot','ballot/activateballot.asp')
			oCMenu.makeMenu('ballotresults','votingadmin','Old Ballot Results','ballot/results.asp')
		oCMenu.makeMenu('player','admin','Player','adminops.asp?aType=Player')
		oCMenu.makeMenu('team','admin','Team','adminops.asp?aType=Team')
		oCMenu.makeMenu('addtourny','admin','Add Tournament','tournament/createtourny.asp')

		oCMenu.makeMenu('sysadminstuff','admin','Sysadmin Tools ','')
			oCMenu.makeMenu('emailsearch','sysadminstuff','Email Search','emailsearch.asp')
			oCMenu.makeMenu('status','sysadminstuff','Server Status','reports/server_status.asp')
			oCMenu.makeMenu('tracker','sysadminstuff','IP Tracker','tracker.asp')
			oCMenu.makeMenu('ipban','sysadminstuff','IP Banner','ipban.asp')
			oCMenu.makeMenu('gamelist','sysadminstuff','Game List','gamelist.asp')
			oCMenu.makeMenu('newgame','sysadminstuff','New Game','addgame.asp')

		oCMenu.makeMenu('pladmin','admin','Player Ladders ')
			oCMenu.makeMenu('addpladmin','pladmin','Add Player Ladder','addplayerladder.asp')
			oCMenu.makeMenu('listpladmin','pladmin','List Player Ladders','playerladderlist.asp')

	<% End If %>
<% Else %>

	oCMenu.makeMenu('help', '', 'help', '', '', '','', '', '', '', '', '', '', '', 'javascript:popup("/help", "help", 300, 400, "yes");') 
<% End If %>oCMenu.construct()	
</SCRIPT>