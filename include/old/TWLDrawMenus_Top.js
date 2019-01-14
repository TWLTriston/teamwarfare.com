function fMyTWLLoggedIn(strName) {
	fMakeMenu(arrParents, 'my_twl','',strName,'default.asp')
	fMakeMenu(arrMyTWL, 'my_twl','',strName,'default.asp')
			fMakeMenu(arrMyTWL, 'home','my_twl','Home', 'default.asp')
			fMakeMenu(arrMyTWL, 'account','my_twl','Account Maintenance')
			fMakeMenu(arrMyTWL, 'profile','account','Profile','viewplayer.asp?player=' + escape(strName))
			fMakeMenu(arrMyTWL, 'edit_profile','account','Edit Profile','addplayer.asp?IsEdit=true')
			fMakeMenu(arrMyTWL, 'register_team','account','Register Team','addteam.asp')
			fMakeMenu(arrMyTWL, 'logout','account','Logout', '', 'fPopLogin();')
			fMakeMenu(arrMyTWL, 'preferences','account','Preferences','preferences.asp')
			fMakeMenu(arrMyTWL, 'requestnamechange','account','Request Name Change','request/ReqNameChange.asp?player=' + escape(strName))
	
		fMakeMenu(arrMyTWL, 'news_archive','my_twl','News Archive','newsarchive.asp')
		fMakeMenu(arrMyTWL, 'search','my_twl','Search','')
			fMakeMenu(arrMyTWL, 'search1','search','Find Player By Name','searchPlayerByName.asp')
			fMakeMenu(arrMyTWL, 'search4','search','Find Player By In Game Identifier','searchPlayerByIdentifier.asp')
			fMakeMenu(arrMyTWL, 'search2','search','Find Team By Name','searchTeamByName.asp')
			fMakeMenu(arrMyTWL, 'search3','search','Find Team By Tag','searchTeamByTag.asp')
		
		fMakeMenu(arrMyTWL, 'myteams','my_twl','My Teams')

}

function fMyTWL() {
	fMakeMenu(arrParents, 'my_twl','','my twl', 'default.asp')
		fMakeMenu(arrMyTWL, 'my_twl','','my twl', 'default.asp')
		fMakeMenu(arrMyTWL, 'home','my_twl','Home', 'default.asp')
		fMakeMenu(arrMyTWL, 'login','my_twl','Login', '', 'fPopLogin();')
		fMakeMenu(arrMyTWL, 'forgot_password','my_twl','Forgot Password','forgotpassword.asp')
		fMakeMenu(arrMyTWL, 'activate','my_twl','Deactivated Account?','activate.asp')
		fMakeMenu(arrMyTWL, 'register','my_twl','Register','addplayer.asp')
		fMakeMenu(arrMyTWL, 'contactus','my_twl','Contact Us','staff.asp')
}

function fOperations() {				
	fMakeMenu(arrParents, 'oper','','operations')
		fMakeMenu(arrOperations, 'oper','','operations')
		fMakeMenu(arrOperations, 'files','oper','Downloads / Files','files')
		fMakeMenu(arrOperations, 'stuff','oper','TWL Stuff Store','stuff/')
		fMakeMenu(arrOperations, 'staff','oper','Contact Us / Staff','staff.asp')
		fMakeMenu(arrOperations, 'winners','oper','Prize Winners','winners.asp')
	
		fMakeMenu(arrOperations, 'demos','oper','Demo Library', 'demos')
		fMakeMenu(arrOperations, 'Voting','oper','Voting Booth', 'ballot/')
		fMakeMenu(arrOperations, 'Contrib','oper','TWL Contributors', 'contributors.asp')
}