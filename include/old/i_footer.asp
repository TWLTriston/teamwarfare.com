<%
Call ShowFooter()
%>
<div name="divContainer" id="divContainer" style="position: absolute; top: 95px; left: 0px; visibility: hidden; ">
<script language="javascript" type="text/javascript">
<% If Session("LoggedIn") = True Then %>
	fMyTWLLoggedIn('<%=Server.HTMLEncode(Replace(Replace(Session("uName"), "\", "\\"), "'", "\'"))%>');
	<%	
	Dim intNumber, strCurrentTeam
	Dim strLName, strLAbbr, strTName, strTTag, strMenuTeamName
	intNumber = 0
	strMenuTeamName = -1

	strSQL = "EXECUTE PlayerGetTeams '" & Session("PlayerID") & "'"
	If (oRS.State <> 0) THen
		Set oRs = nothing
		Set oRs = Server.CreateObject("ADODB.RecordSet")
	End If
	oRS.Open strSQL, oConn
	If Not (oRS.EOF AND oRS.BOF) Then
		Do While Not(oRS.EOF)
			If strMenuTeamName <> oRS.Fields("TeamName").Value Then 
					intNumber = intNumber + 1
					Response.Write "fMakeMenu(arrMyTWL, 'myteams" & intNumber & "','myteams','" & Server.HTMLEncode(Replace(Replace(ors.fields("TeamTag").value, "\", "\\"), "'", "\'")) & "','viewteam.asp?team=" & Server.URLEncode(ors.fields("TeamName").value) & "')" & vbcrlf
					strMenuTeamName = oRS.Fields("TeamName").Value
			End If
			oRS.MoveNext
		Loop
	Else
		Response.Write "fMakeMenu(arrMyTWL, 'myteams" & intNumber & "','myteams','No teams found.')" & vbcrlf
	End If
	oRS.NextRecordSet

	%>
	fMakeMenu(arrMyTWL, 'contactus','my_twl','Contact Us','staff.asp')
<% Else %>
	fMyTWL();
<% End If %>
fForumsAll();
<% If bSysAdmin Or bAnyLadderAdmin Then %>
fForumsStaff();
fForumsStaff2();
	<% If IsSysAdminLevel2() Then %>
		fMakeMenu(arrForums, 'sysAdminForums','dbm553','SysAdmin Forums','forums/forumdisplay.asp?forumid=8')
		fMakeMenu(arrForums, 'forum8','sysAdminForums','SysAdmin Forum','forums/forumdisplay.asp?forumid=8')
		fMakeMenu(arrForums, 'forum44','sysAdminForums','Quality Control','forums/forumdisplay.asp?forumid=44')
		fMakeMenu(arrForums, 'forum96','sysAdminForums','TWLHosting.com','forums/forumdisplay.asp?forumid=96')
	<% End If %>
<% End If %>
fCompetition();
fRules();
fOperations();
<% If bSysAdmin Or bAnyLadderAdmin Then %>
	fMakeMenu(arrParents, 'admin','','admin')
	fMakeMenu(arrHelp, 'admin','','admin')
		fMakeMenu(arrHelp, 'admain','admin','Menu','adminmenu.asp')
		fMakeMenu(arrHelp, 'adopsnew','admin','News','newsdesk.asp')
		fMakeMenu(arrHelp, 'teamLadMenu','admin','Team Ladders ','')
					fMakeMenu(arrHelp, 'adopsmat','teamLadMenu','Match','adminops.asp?aType=Match')
			fMakeMenu(arrHelp, 'adopsfor','teamLadMenu','Forfeit','adminops.asp?aType=Forfeit')
			fMakeMenu(arrHelp, 'adopshis','teamLadMenu','History','adminops.asp?aType=History')
			fMakeMenu(arrHelp, 'adopslad','teamLadMenu','Ladder','adminops.asp?aType=Ladder')
			fMakeMenu(arrHelp, 'adopsrank','teamLadMenu','Rank','adminops.asp?aType=Rank')
			fMakeMenu(arrHelp, 'adopslidrep','teamLadMenu','GUID Report','reports/identifierreport.asp')
			<% If bSysadmin Then %>
				<% If IsSysAdminLevel2() Then %>
					fMakeMenu(arrHelp, 'laAdmins','teamLadMenu','Assign Admins','assignadmin.asp')
				<% End If %>
			fMakeMenu(arrHelp, 'ladmatchoptions','teamLadMenu','Match Options','ladder/ladderoptions.asp')
			fMakeMenu(arrHelp, 'addladder','teamLadMenu','Add Ladder','addladder.asp')
			<% End If %>
		<% If bSysadmin Then %>
			fMakeMenu(arrHelp, 'scrimlad','admin','Scrim Ladders','scrim/generaladmin.asp')
				fMakeMenu(arrHelp, 'scrimladgen','scrimlad','General Admin','scrim/generaladmin.asp')
				fMakeMenu(arrHelp, 'scrimladadm','scrimlad','Assign Admins','scrim/assignadmin.asp')
				fMakeMenu(arrHelp, 'scrimguidrep','scrimlad','GUID Report','reports/scrimidentifierreport.asp')
		<% End If %>
		fMakeMenu(arrHelp, 'playerLadMenu','admin','Player Ladders ','')
			fMakeMenu(arrHelp, 'adopspmat','playerLadMenu','Match','adminops.asp?aType=PMatch')
			fMakeMenu(arrHelp, 'adopspfor','playerLadMenu','Forfeit','adminops.asp?aType=PForfeit')
			fMakeMenu(arrHelp, 'adopsplad','playerLadMenu','Ladder Admin','adminops.asp?aType=PLadder')
			fMakeMenu(arrHelp, 'adopsplrank','playerLadMenu','Player Rank','editplayerrank.asp')
			fMakeMenu(arrHelp, 'adopsplhistory','playerLadMenu','History','edit1v1history.asp')
			<% if bSysAdmin Then %>
			fMakeMenu(arrHelp, 'addpladmin','playerLadMenu','Add Player Ladder','addplayerladder.asp')
			fMakeMenu(arrHelp, 'listpladmin','playerLadMenu','List Player Ladders','playerladderlist.asp')
			<% end if %>
		fMakeMenu(arrHelp, 'leagueadmin','admin','Leagues','leagueadmin.asp')
			fMakeMenu(arrHelp, 'leaguegen','leagueadmin','General Admin','leagueadmin.asp')
			fMakeMenu(arrHelp, 'leagueidrep','leagueadmin','GUID Report','reports/leagueidentifierreport.asp')
		fMakeMenu(arrHelp, 'helprules','admin','Help/Rules','help/admin')
		fMakeMenu(arrHelp, 'reports','admin','Reports ','')
			fMakeMenu(arrHelp, 'plfor','reports','Player Forfeit Report','reports/playerforfietreport.asp')
			fMakeMenu(arrHelp, 'plrr','reports','Player Roster Report','reports/playerrosterreport.asp')
			fMakeMenu(arrHelp, 'rr','reports','Roster Report','reports/rosterreport.asp')
			fMakeMenu(arrHelp, 'act','reports','Ladder Activity','reports/activity.asp')
			fMakeMenu(arrHelp, 'tournyrosterreport','reports','Tournament Roster Report','reports/tournamentrosterreport.asp')
			fMakeMenu(arrHelp, 'leaguerosterreport','reports','League Roster Report','reports/leaguerosterreport.asp')
	<% If bSysAdmin Then %>
			fMakeMenu(arrHelp, 'leagueadd','leagueadmin','Add League','leagueadd.asp')
			<% If IsSysAdminLevel2() Then %>
			fMakeMenu(arrHelp, 'leagueaa','leagueadmin','League Assign Admin','leagueassignadmin.asp')
			<% End If %>
		fMakeMenu(arrHelp, 'votingadmin','admin','Voting Booth ','')
			fMakeMenu(arrHelp, 'addballot','votingadmin','Add Ballot','ballot/addballot.asp')
			fMakeMenu(arrHelp, 'actballot','votingadmin','Activate Ballot','ballot/activateballot.asp')
			fMakeMenu(arrHelp, 'ballotresults','votingadmin','Old Ballot Results','ballot/results.asp')
		fMakeMenu(arrHelp, 'sysadminstuff','admin','Sysadmin Tools ','')
			<% If IsSysAdminLevel2() THen %>
			fMakeMenu(arrHelp, 'menus','sysadminstuff','Update Menus','menu/')
			fMakeMenu(arrHelp, 'player','sysadminstuff','Delete / Sysadmin Player ','adminops.asp?aType=Player')
			fMakeMenu(arrHelp, 'team','sysadminstuff','Delete Team','adminops.asp?aType=Team')
			fMakeMenu(arrHelp, 'rostrans','sysadminstuff','Transfer Rosters','rostertransfer.asp')
			<% End If %>
			fMakeMenu(arrHelp, 'mm','sysadminstuff','Mass Mail','massmail.asp')
			fMakeMenu(arrHelp, 'emailsearch','sysadminstuff','Email Search','emailsearch.asp')
			<% If IsSysAdminLevel2() Then %>
			fMakeMenu(arrHelp, 'status','sysadminstuff','Server Status','reports/server_status.asp')
			<% End If %>
			fMakeMenu(arrHelp, 'tracker','sysadminstuff','IP Tracker','tracker.asp')
			fMakeMenu(arrHelp, 'systourny','admin','Tournaments','')
			fMakeMenu(arrHelp, 'tournygenera','systourny','General Admin','tournament/admintournament.asp')
			fMakeMenu(arrHelp, 'tournyassign','systourny','Assign Admin','tournament/AssignAdmin.asp')
			<% If IsSysAdminLevel2() Then %>
			fMakeMenu(arrHelp, 'addtourny','systourny','Add Tournament','tournament/createtourny.asp')
			fMakeMenu(arrHelp, 'ipban','sysadminstuff','IP Banner','ipban.asp')
			fMakeMenu(arrHelp, 'gamelist','sysadminstuff','Game List','gamelist.asp')
			fMakeMenu(arrHelp, 'newgame','sysadminstuff','New Game','addgame.asp')
			fMakeMenu(arrHelp, 'forum','sysadminstuff','Forum','forums/admin')
			<% End If %>
	<% End If %>
<% Else %>
	fMakeMenu(arrParents, 'help','','help', '', 'fPopHelp();')
<% End If %>
<% 
Select Case Session("StyleID")
	Case 3, 5, 6, 7, 8, 9
		%>
		fDrawMenus(arrParents, '', 0, 0);
		fDrawMenus(arrMyTWL, 'mytwl', 0, 0);
		fDrawMenus(arrForums, 'forums', 165);
		fDrawMenus(arrOperations, 'operations', 660);
		fDrawMenus(arrHelp, 'help', 825);
		<%
	Case Else
		%>
		fDrawMenus(arrParents, '', 0, 0);
		fDrawMenus(arrMyTWL, 'mytwl', 0, 0);
		fDrawMenus(arrForums, 'forums', 130);
		fDrawMenus(arrComp, 'comp', 260);
		fDrawMenus(arrRules, 'rules', 390);
		fDrawMenus(arrOperations, 'operations', 520);
		fDrawMenus(arrHelp, 'help', 650);
		<%
End Select
%>
fCenterMenus();
</script>
</div>

<% If Session("StyleID") = 7 Or Session("StyleID") = 8 Or Session("StyleID") = 9 Then 
	%>
	<% If False Then %>
		<div id="divAdSpace1Cache" style="display: none;">
				
		</div>
	<% End if %>
	<div id="divAdSpace2Cache" style="display: none;">
	<%
	If LCase(Request.ServerVariables("PATH_INFO")) = "/default.asp" AND True Then 
		' Ghost Recon
		%>
<script language="javascript" type="text/javascript">
Ads_kid=0;Ads_bid=0;Ads_xl=0;Ads_yl=0;Ads_xp='';Ads_yp='';Ads_xp1='';Ads_yp1='';Ads_opt=0;Ads_par='';Ads_cnturl='';
</script>
<script type="text/javascript" language="javascript" src="http://a.as-us.falkag.net/dat/cjf/00/14/54/90.js"></script>		
<%
	ElseIf False Then
		'' TWL Hosting
		Dim iRandomBanner
		Randomize
		iRandomBanner = Int((3) * Rnd)
		If iRandomBanner = 0 Then
			%><a href="https://secure.trinitygames.com/twl_hosting/signup.php?package=677"><img src="/images/ads/twlhosting_bf2.jpg" height="600" width="120" alt="" border="0" /></a><%
		ElseIf iRandomBanner = 1 Then
			%><a href="http://twlhosting.teamwarfare.com/"><img src="/images/ads/120x600_AmericasArmy.jpg" height="600" width="120" alt="" border="0" /></a><%
		Else
			%><a href="http://twlhosting.teamwarfare.com/"><img src="/images/ads/120x600_CounterStrikeSource.jpg" height="600" width="120" alt="" border="0" /></a><%
		End If
	ElseIf True Then 
		' Kawasaki
		%>
		<script language="javascript" type="text/javascript">
		Ads_kid=0;Ads_bid=0;Ads_xl=0;Ads_yl=0;
		</script>
		<script type="text/javascript" language="javascript" src="http://a.as-us.falkag.net/dat/cjf/00/09/33/53.js"></script>
		<%
	End If
	%>
	</div>
	
	<script language="javascript" type="text/javascript">
	<% If False Then %>
		document.getElementById("divAdSpace1").innerHTML = document.getElementById("divAdSpace1Cache").innerHTML ;
	<% End If %>
	document.getElementById("divAdSpace2").innerHTML = document.getElementById("divAdSpace2Cache").innerHTML ;
	</script>
	<%
End If
%>
 <script src="http://www.google-analytics.com/urchin.js" type="text/javascript"> 
 </script> 
 <script type="text/javascript"> 
 _uacct = "UA-271929-1"; 
 urchinTracker(); 
 </script>

</body>
</html>
