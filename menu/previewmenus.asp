<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Update TWL Menus"

Dim strSQL, oConn, oRS, oRs2, oRs3, oRs4
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")
Set oRS3 = Server.CreateObject("ADODB.RecordSet")
Set oRS4 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

if not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title><%=strPageTitle%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<%
	Select Case Session("StyleID")
		Case 2
			bgcone = "#00163D"
			bgctwo = "#00102B"
			bgcheader = "#002C73"
			bgcblack = "#000000"
			%>
			<link rel=stylesheet href="/core/style2.css" type="text/css">
			<script language="javascript" type="text/javascript" src="/include/TWLMenus2.js"></script>
			<%
		Case Else
			bgcone = "#3C0000"
			bgctwo = "#2B0000"
			bgcheader = "#2B0000"
			bgcblack = "#000000"
			%>
			<link rel=stylesheet href="/core/style.css" type="text/css">
			<script language="javascript" type="text/javascript" src="/include/TWLMenus.js"></script>
			<%
	End Select
	%>
	<script language="javascript" type="text/javascript">
	<!-- #include virtual="/include/TWLDrawMenus_Top.js" -->
	function fForumsAll() {
		fMakeMenu(arrParents, 'dbm2574','','forums','forums/')
		fMakeMenu(arrForums, 'dbm2574','','forums','forums/')
		<% Call fWriteMenu("arrForums", 2574, true) %>
	}
	function fForumsStaff() {
		fMakeMenu(arrForums, 'dbm2573','dbm2573','Staff Forums','')
		<% Call fWriteMenu("arrForums", 2573, true) %>
	}
	function fForumsStaff2() {
		fMakeMenu(arrForums, 'dbm2572','dbm2572','Staff Gaming Forums','')
		<% Call fWriteMenu("arrForums", 2572, true) %>
	}
	function fCompetition() {
		fMakeMenu(arrParents, 'dbm1','','competition','ladderlist.asp')
		fMakeMenu(arrComp, 'dbm1','','competition','ladderlist.asp')
		<% Call fWriteMenu("arrComp", 1, true) %>
	
	}	
	
	function fRules() {
		fMakeMenu(arrParents, 'dbm2','','rules','viewrules.asp?ruleset=2')
		fMakeMenu(arrRules, 'dbm2','','rules','viewrules.asp?ruleset=2')
		<% Call fWriteMenu("arrRules", 2, false) %>
	}
	</script>
</head>

<body bgcolor="#000000" text="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" ONLOAD="fCenterMenus();preload();" onresize="fCenterMenus();">
<%
Call ShowBanner()
Call ShowAbsTop()
%>

<%
Call ContentStart("Preview TWL Menus")
%>
How do they look?<br />
<a href="updatemenus.asp">Click here to change something</a><br />
<a href="makemenus.asp">Click here to make these the current menus</a><br />
<%
Call ContentEnd()
Call ShowFooter()
%>
<div name="divContainer" id="divContainer" style="position: absolute; top: 0px; left: 0px; visibility: hidden; ">
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
	<% If bSysAdmin Then %>
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
			<% If bSysadmin Then %>
			fMakeMenu(arrHelp, 'laAdmins','teamLadMenu','Assign Admins','assignadmin.asp')
			fMakeMenu(arrHelp, 'ladmatchoptions','teamLadMenu','Match Options','ladder/ladderoptions.asp')
			fMakeMenu(arrHelp, 'addladder','teamLadMenu','Add Ladder','addladder.asp')
			<% End If %>
		fMakeMenu(arrHelp, 'playerLadMenu','admin','Player Ladders ','')
			fMakeMenu(arrHelp, 'adopspmat','playerLadMenu','Match','adminops.asp?aType=PMatch')
			fMakeMenu(arrHelp, 'adopspfor','playerLadMenu','Forfeit','adminops.asp?aType=PForfeit')
			fMakeMenu(arrHelp, 'adopsplad','playerLadMenu','Ladder Admin','adminops.asp?aType=PLadder')
			fMakeMenu(arrHelp, 'adopsplrank','playerLadMenu','Player Rank','editplayerrank.asp')
			<% if bSysAdmin Then %>
			fMakeMenu(arrHelp, 'addpladmin','playerLadMenu','Add Player Ladder','addplayerladder.asp')
			fMakeMenu(arrHelp, 'listpladmin','playerLadMenu','List Player Ladders','playerladderlist.asp')
			<% end if %>
		fMakeMenu(arrHelp, 'leagueadmin','admin','Leagues','leagueadmin.asp')
			fMakeMenu(arrHelp, 'leaguegen','leagueadmin','General Admin','leagueadmin.asp')
		fMakeMenu(arrHelp, 'helprules','admin','Help/Rules','help/admin')
		fMakeMenu(arrHelp, 'reports','admin','Reports ','')
			fMakeMenu(arrHelp, 'plfor','reports','Player Forfeit Report','reports/playerforfietreport.asp')
			fMakeMenu(arrHelp, 'plrr','reports','Player Roster Report','reports/playerrosterreport.asp')
			fMakeMenu(arrHelp, 'rr','reports','Roster Report','reports/rosterreport.asp')
			fMakeMenu(arrHelp, 'act','reports','Ladder Activity','reports/activity.asp')
	<% If bSysAdmin Then %>
			fMakeMenu(arrHelp, 'leagueadd','leagueadmin','Add League','leagueadd.asp')
			fMakeMenu(arrHelp, 'leagueaa','leagueadmin','League Assign Admin','leagueassignadmin.asp')
		fMakeMenu(arrHelp, 'votingadmin','admin','Voting Booth ','')
			fMakeMenu(arrHelp, 'addballot','votingadmin','Add Ballot','ballot/addballot.asp')
			fMakeMenu(arrHelp, 'actballot','votingadmin','Activate Ballot','ballot/activateballot.asp')
			fMakeMenu(arrHelp, 'ballotresults','votingadmin','Old Ballot Results','ballot/results.asp')
		fMakeMenu(arrHelp, 'sysadminstuff','admin','Sysadmin Tools ','')
			fMakeMenu(arrHelp, 'menus','sysadminstuff','Update Menus','menu/')
			fMakeMenu(arrHelp, 'player','sysadminstuff','Delete / Sysadmin Player ','adminops.asp?aType=Player')
			fMakeMenu(arrHelp, 'team','sysadminstuff','Delete Team','adminops.asp?aType=Team')
			fMakeMenu(arrHelp, 'mm','sysadminstuff','Mass Mail','massmail.asp')
			fMakeMenu(arrHelp, 'emailsearch','sysadminstuff','Email Search','emailsearch.asp')
			fMakeMenu(arrHelp, 'status','sysadminstuff','Server Status','reports/server_status.asp')
			fMakeMenu(arrHelp, 'tracker','sysadminstuff','IP Tracker','tracker.asp')
			fMakeMenu(arrHelp, 'ipban','sysadminstuff','IP Banner','ipban.asp')
			fMakeMenu(arrHelp, 'gamelist','sysadminstuff','Game List','gamelist.asp')
			fMakeMenu(arrHelp, 'newgame','sysadminstuff','New Game','addgame.asp')
			fMakeMenu(arrHelp, 'forum','sysadminstuff','Forum','forums/admin')
			fMakeMenu(arrHelp, 'addtourny','sysadminstuff','Add Tournament','tournament/createtourny.asp')
	<% End If %>
<% Else %>
	fMakeMenu(arrParents, 'help','','help')
<% End If %>fDrawMenus(arrParents, '', 0, 0);
fDrawMenus(arrMyTWL, 'mytwl', 0, 0);
fDrawMenus(arrForums, 'forums', 130);
fDrawMenus(arrComp, 'comp', 260);
fDrawMenus(arrRules, 'rules', 390);
fDrawMenus(arrOperations, 'operations', 520);
fDrawMenus(arrHelp, 'help', 650);
</script>
</div>
</body>
</html>

<%
Function DisplayMenu(strArray, strMenuName, intMenuID, intParentMenuID, byVal linkURL, spec, spec2)
	Dim strText
	if Not(IsNull(linkURL)) AND Len(Trim(linkURL)) > 0 Then
		linkURL = Replace(linkURL, "http://www.teamwarfare.com/", "")
		If Left(linkURL, 1) = "/" Then
			linkUrl = Right(linkURL, Len(linkURL - 1))
		End If
	end if
	if intParentMenuID = "0" Then
		intParentMenuID = ""
	Else
		intParentMenuID = "dbm" & intParentMenuID
	End If
	strText = spec2 & "fMakeMenu(" & strArray & ", 'dbm" & intMenuID & "','" & intParentMenuID & "','" & Replace(Server.HTMLEncode(strMenuName & ""), "'", "\'") & "','" & Replace(Server.HTMLEncode(linkURL & ""), "'", "\'") & "')" & vbCrLf
	DisplayMenu = strText
End Function

Function fWriteMenu(strArray, iParentMenuID, blnTier2) 
	strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & iParentMenuID & " ORDER BY SortOrder, ShowMenuName "
	oRs2.Open strSQL, oConn
	If Not(oRs2.EOF AND oRs2.BOF) Then
		Do While Not(oRs2.EOF)
			Response.Write DisplayMenu(strArray, oRs2.Fields("ShowMenuName").Value, oRs2.Fields("MenuID").Value, oRs2.Fields("ParentMenuID").Value, oRs2.Fields("LinkURL").Value, "", vbTab & vbTab)
			If (blnTier2) Then
				strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & oRs2.Fields("MenuID").Value & " ORDER BY SortOrder, ShowMenuName "
				oRs3.Open strSQL, oConn
				If Not(oRs3.EOF AND oRs3.BOF) Then
					Do While Not(oRs3.EOF)
						Response.Write DisplayMenu(strArray, ors3.Fields("ShowMenuName").Value, ors3.Fields("MenuID").Value, ors3.Fields("ParentMenuID").Value, ors3.Fields("LinkURL").Value, "", vbTab & vbTab & vbTab)
						strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & oRs3.Fields("MenuID").Value & " ORDER BY SortOrder, ShowMenuName "
						oRs4.Open strSQL, oConn
						If Not(oRs4.EOF AND oRs4.BOF) Then
							Do While Not(oRs4.EOF)
								Response.Write DisplayMenu(strArray, ors4.Fields("ShowMenuName").Value, ors4.Fields("MenuID").Value, ors4.Fields("ParentMenuID").Value, ors4.Fields("LinkURL").Value, "", vbTab & vbTab & vbTab & vbTab)
								oRs4.MoveNext
							Loop
						End If
						oRs4.NextRecordSet
						oRs3.MoveNext
					Loop
				End If
				oRs3.NextRecordSet
			End If
			oRs2.MoveNext
		Loop
	End If
	oRs2.NextRecordSet
End Function
%>

