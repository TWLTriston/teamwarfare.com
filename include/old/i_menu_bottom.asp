}
 	
function fMenu6() {
	oCMenu.makeMenu('oper','','operations')
		oCMenu.makeMenu('files','oper','Downloads / Files','files')
		oCMenu.makeMenu('staff','oper','Staff','staff.asp')
		oCMenu.makeMenu('java_irc','oper','Java IRC','jirc/')
		oCMenu.makeMenu('winners','oper','Prize Winners','winners.asp')
	
		oCMenu.makeMenu('demos','oper','Demo Library', 'demos')
		oCMenu.makeMenu('Voting','oper','Voting Booth', 'ballot/')
		oCMenu.makeMenu('Contrib','oper','TWL Contributors', 'contributors.asp')
}

var oCMenu;

fMenu1();
fMenu2();
fMenu3();
fMenu4();
fMenu5();
fMenu6();
fMenu7();

oCMenu.construct()	
