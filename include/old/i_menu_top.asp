function fMenu1() {
	oCMenu=new makeCM("oCMenu");
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
	oCMenu.useBar=0
	oCMenu.barWidth="100%"
	oCMenu.barHeight="menu" 
	oCMenu.barClass="clBar"
	oCMenu.barX=0 
	oCMenu.barY=0
	oCMenu.barBorderX=0
	oCMenu.barBorderY=0
	oCMenu.barBorderClass=""
	oCMenu.level[0]=new cm_makeLevel()
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
	oCMenu.level[1]=new cm_makeLevel()
	oCMenu.level[1].width=200
	oCMenu.level[1].height=22
	oCMenu.level[1].regClass="clLevel1"
	oCMenu.level[1].overClass="clLevel1over"
	oCMenu.level[1].borderX=1
	oCMenu.level[1].borderY=1
	oCMenu.level[1].align="left" 
	oCMenu.level[1].arrow="images/tri.gif"
	oCMenu.level[1].arrowWidth=10
	oCMenu.level[1].arrowHeight=9
	oCMenu.level[1].offsetX=350
	oCMenu.level[1].offsetY=0
	oCMenu.level[1].borderClass="clLevel1border"
	oCMenu.level[2]=new cm_makeLevel()
	oCMenu.level[2].width=200
	oCMenu.level[2].height=20
	oCMenu.level[2].offsetX=350
	oCMenu.level[2].offsetY=0
	oCMenu.level[2].arrow="images/tri.gif"
	oCMenu.level[2].regClass="clLevel2"
	oCMenu.level[2].overClass="clLevel2over"
	oCMenu.level[2].borderClass="clLevel2border"
	oCMenu.level[3]=new cm_makeLevel()
	oCMenu.level[3].width=200
	oCMenu.level[3].height=20
	oCMenu.level[3].offsetX=150
	oCMenu.level[3].offsetY=0
	oCMenu.level[3].arrow=0
	oCMenu.level[3].regClass="clLevel3"
	oCMenu.level[3].overClass="clLevel3over"
	oCMenu.level[3].borderClass="clLevel3border"
}function fMenu3() {
	oCMenu.makeMenu('forums','','forums','forums/')
		oCMenu.makeMenu('forumindex','forums','Forum Index','forums/')
		oCMenu.makeMenu('forum9','forums','General Forums','forums/forumdisplay.asp?forumid=9')
		oCMenu.makeMenu('forum19','forums','Site Support','forums/forumdisplay.asp?forumid=19')
		oCMenu.makeMenu('forum11','forums','Recruiting','forums/forumdisplay.asp?forumid=11')
		oCMenu.makeMenu('gameForums','forums','Game Specific','forums/default.asp#Category1')
		oCMenu.makeMenu('forum54','gameForums','AA CDS Support','forums/forumdisplay.asp?forumid=54')
		oCMenu.makeMenu('forum25','gameForums','America\'s Army','forums/forumdisplay.asp?forumid=25')
		oCMenu.makeMenu('forum27','gameForums','Battlefield 1942','forums/forumdisplay.asp?forumid=27')
		oCMenu.makeMenu('forum101','gameForums','Battlefield 1942: Desert Combat','forums/forumdisplay.asp?forumid=101')
		oCMenu.makeMenu('cwforms','gameForums','Castle Wolfenstein','')
		oCMenu.makeMenu('forum12','cwforms','Castle Wolfenstein','forums/forumdisplay.asp?forumid=12')
		oCMenu.makeMenu('forum40','cwforms','Castle Wolfenstein: Enemy Territory','forums/forumdisplay.asp?forumid=40')
		oCMenu.makeMenu('forum52','cwforms','Castle Wolfenstein: Shrub','forums/forumdisplay.asp?forumid=52')
		oCMenu.makeMenu('forum38','gameForums','C&amp;C: Generals','forums/forumdisplay.asp?forumid=38')
		oCMenu.makeMenu('forum14','gameForums','Counter Strike','forums/forumdisplay.asp?forumid=14')
		oCMenu.makeMenu('forum51','gameForums','Delta Force: Black Hawk Down','forums/forumdisplay.asp?forumid=51')
		oCMenu.makeMenu('forum58','gameForums','Day of Defeat','forums/forumdisplay.asp?forumid=58')
		oCMenu.makeMenu('forum24','gameForums','Ghost Recon','forums/forumdisplay.asp?forumid=24')
		oCMenu.makeMenu('forum99','gameForums','Global Operations','forums/forumdisplay.asp?forumid=99')
		oCMenu.makeMenu('forum18','gameForums','Jedi Outcast','forums/forumdisplay.asp?forumid=18')
		oCMenu.makeMenu('forum2','gameForums','Mechwarrior 4','forums/forumdisplay.asp?forumid=2')
		oCMenu.makeMenu('mohforum','gameForums','Medal Of Honor','')
		oCMenu.makeMenu('forum41','mohforum','Medal of Honor League','forums/forumdisplay.asp?forumid=41')
		oCMenu.makeMenu('forum13','mohforum','Medal of Honor: Allied Assault','forums/forumdisplay.asp?forumid=13')
		oCMenu.makeMenu('forum37','mohforum','Medal of Honor: Spearhead','forums/forumdisplay.asp?forumid=37')
		oCMenu.makeMenu('forum94','gameForums','Savage','forums/forumdisplay.asp?forumid=94')
		oCMenu.makeMenu('forum23','gameForums','Soldier Of Fortune 2','forums/forumdisplay.asp?forumid=23')
		oCMenu.makeMenu('forum100','gameForums','Söldner','forums/forumdisplay.asp?forumid=100')
		oCMenu.makeMenu('forum34','gameForums','Rainbow Six: Raven Shield','forums/forumdisplay.asp?forumid=34')
		oCMenu.makeMenu('forum98','gameForums','Rise of Nations','forums/forumdisplay.asp?forumid=98')
		oCMenu.makeMenu('forum1','gameForums','Tribes','forums/forumdisplay.asp?forumid=1')
		oCMenu.makeMenu('t2forum','gameForums','Tribes 2','')
		oCMenu.makeMenu('forum4','t2forum','Tribes 2 CTF','forums/forumdisplay.asp?forumid=4')
		oCMenu.makeMenu('forum53','t2forum','Tribes 2 Base','forums/forumdisplay.asp?forumid=53')
		oCMenu.makeMenu('forum17','t2forum','Tribes Renegades','forums/forumdisplay.asp?forumid=17')
		oCMenu.makeMenu('forum10','gameForums','UT2003','forums/forumdisplay.asp?forumid=10')
		oCMenu.makeMenu('forum26','gameForums','Warcraft 3','forums/forumdisplay.asp?forumid=26')
		oCMenu.makeMenu('forum32','gameForums','Urban Terror','forums/forumdisplay.asp?forumid=32')
		oCMenu.makeMenu('forum35','forums','Match Observers','forums/forumdisplay.asp?forumid=35')
		oCMenu.makeMenu('xboxforum','forums','Xbox Live','forums/forumdisplay.asp?forumid=57')
			oCMenu.makeMenu('forum57','xboxforum','General Xbox Live','forums/forumdisplay.asp?forumid=57')
			oCMenu.makeMenu('forum84','xboxforum','Castle Wolfenstein','forums/forumdisplay.asp?forumid=84')
			oCMenu.makeMenu('forum56','xboxforum','Ghost Recon','forums/forumdisplay.asp?forumid=56')
			oCMenu.makeMenu('forum91','xboxforum','Unreal Championship','forums/forumdisplay.asp?forumid=91')
		oCMenu.makeMenu('ps2forum','forums','Play Station 2 Online','forums/forumdisplay.asp?forumid=87')
			oCMenu.makeMenu('forum87','ps2forum','General PS2 Online','forums/forumdisplay.asp?forumid=87')
			oCMenu.makeMenu('forum86','ps2forum','Tribes: Aerial Assault','forums/forumdisplay.asp?forumid=86')
			oCMenu.makeMenu('forum89','ps2forum','Tribes: Aerial Assault Scheduling','forums/forumdisplay.asp?forumid=89')
}

function fMenu5() {
