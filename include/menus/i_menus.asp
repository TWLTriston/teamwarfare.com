<%
'------------------------------------------------------------------------------------
' Created on 6/3/2006
' This is the file that will draw the entire menu structure
'------------------------------------------------------------------------------------
Dim blnMenuLoggedIn
blnMenuLoggedIn = Session("LoggedIn")
%>
<div id="divNavigation" class="noAd">
	<a href="#divMainWrapper" title="Skip Navigation" class="access">Skip Navigation</a>
	<ul id="nav">
	<%
'------------------------------ 
' my twl menu
'------------------------------ 
%><li><% If blnMenuLoggedIn Then %><a href="/"><%=Server.HTMLEncode(Session("uName"))%></a><% Else %><a href="/">my twl</a><% End If %><ul><li><a href="/">Home</a></li><% If blnMenuLoggedIn Then %><li><a href="/viewplayer.asp?player=<%=Server.URLEncode(Session("uName"))%>" class="haschild">Account Maintenance</a><ul><li><a href="/viewplayer.asp?player=<%=Server.URLEncode(Session("uName"))%>">Profile</a></li><li><a href="/addplayer.asp?IsEdit=true">Edit Profile</a></li><li><a href="/addteam.asp">Register Team</a></li><li><a href="javascript:fPopLogin();">Logout</a></li><li><a href="/preferences.asp">Preferences</a></li><li><a href="/request/ReqNameChange.asp?player=<%=Server.URLEncode(Session("uName"))%>">Request Name Change</a></li></ul></li><li><a href="/newsarchive.asp">News Archive</a></li><li><a href="searchPlayerByName.asp" class="haschild">Search</a><ul><li><a href="/searchPlayerByName.asp">Find Player By Name</a></li><li><a href="/searchPlayerByIdentifier.asp">Find Player By In Game Identifier</a></li><li><a href="/searchTeamByName.asp">Find Team By Name</a></li><li><a href="/searchTeamByTag.asp">Find Team By Tag</a></li></ul></li><li><a href="/viewplayer.asp?player=<%=Server.URLEncode(Session("uName"))%>" class="haschild">My Teams</a><ul><%
Dim intNumber, strMenuTeamName
intNumber = 0
strMenuTeamName = -1
strSQL = "EXECUTE PlayerGetTeams '" & Session("PlayerID") & "'"
If (oRS.State <> 0) Then
Set oRs = Nothing
Set oRs = Server.CreateObject("ADODB.RecordSet")
End If
oRS.Open strSQL, oConn
If Not (oRS.EOF AND oRS.BOF) Then
Do While Not(oRS.EOF)
If strMenuTeamName <> oRS.Fields("TeamName").Value Then 
intNumber = intNumber + 1
Response.Write "<li><a href=""/viewteam.asp?team=" & Server.URLEncode(oRs.Fields("TeamName").Value) & """>" & Server.HTMLEncode(oRs.Fields("TeamName").Value) & " " & Server.HTMLEncode(oRs.Fields("TeamTag").Value) & "</a></li>"
strMenuTeamName = oRS.Fields("TeamName").Value
End If
oRS.MoveNext
Loop
Else
Response.Write "<li><a href=""/viewplayer.asp?player=" & Server.URLEncode(Session("uName")) & """>No teams found.</a></li>"
End If
oRS.NextRecordSet
%></ul></li><li><a href="/staff.asp">Contact Us</a></li><% Else %><li><a href="javascript:fPopLogin();">Login</a></li><li><a href="/forgotpassword.asp">Forgot Password</a></li><li><a href="activate.asp">Deactivated Account?</a></li><li><a href="addplayer.asp">Register</a></li><li><a href="/staff.asp">Contact Us</a></li><% End If %></ul></li><%
'------------------------------ 
' forums
'------------------------------ 
%><li><a href="/forums/">forums</a><ul><!-- #include file="i_forums_public.asp" --><% 
If bSysAdmin Or bAnyLadderAdmin Then %><li><a class="haschild" href="#">Staff Forums</a><ul><!-- #include file="i_forums_staff_1.asp" --></ul></li><li><a class="haschild" href="#">Staff Gaming Forums</a><ul><!-- #include file="i_forums_staff_2.asp" --></ul></li><% If IsSysAdminLevel2() Then %><li><a href="/forums/forumdisplay.asp?forumid=8" class="haschild">SysAdmin Forums</a><ul><li><a href="/forums/forumdisplay.asp?forumid=8">SysAdmin Forum</a></li><li><a href="/forums/forumdisplay.asp?forumid=44">Quality Control</a></li><li><a href="/forums/forumdisplay.asp?forumid=96">TWLHosting.com</a></li></ul></li><% End If %><% End If %></ul></li><li><a href="/ladderlist.asp">competition</a></li><li><a href="/rulechooser.asp">rules</a></li><li><a href="#">operations</a><ul><li><a href="/files/">Downloads / Files</a></li>
<li><a href="/staff.asp">Contact Us / Staff</a></li><li><a href="/support/">Site Support</a></li><li><a href="/demos/">Demo Library</a></li><li><a href="/ballot/">Voting Booth</a></li><li><a href="/contributors.asp">TWL Contributors</a></li></ul></li><%
'------------------------------ 
' help / staff menu
'------------------------------ 
If bSysAdmin Or bAnyLadderAdmin Then %><li><a href="/adminmenu.asp">admin</a><ul><li><a href="/adminmenu.asp">Menu</a></li><li><a href="/newsdesk.asp">News</a></li><li><a href="#" class="haschild">Team Ladders</a><ul><li><a href="/adminops.asp?aType=Match">Match</a></li><li><a href="/adminops.asp?aType=Forfeit">Forfeit</a></li><li><a href="/adminops.asp?aType=History">History</a></li><li><a href="/adminops.asp?aType=Ladder">Ladder</a></li><li><a href="/adminops.asp?aType=Rank">Rank</a></li><li><a href="/reports/identifierreport.asp">GUID Report</a></li><% If bSysadmin Then %><% If IsSysAdminLevel2() Then %><li><a href="/assignadmin.asp">Assign Admins</a></li><% End If %><li><a href="/ladder/ladderoptions.asp">Match Options</a></li><li><a href="/addladder.asp">Add Ladder</a></li><% End If %></ul></li><% If bSysadmin Then %><li><a href="/scrim/generaladmin.asp" class="haschild">Scrim Ladders</a><ul><li><a href="/scrim/generaladmin.asp">General Admin</a></li><li><a href="/scrim/assignadmin.asp">Assign Admins</a></li><li><a href="/reports/scrimidentifierreport.asp">GUID Report</a></li></ul></li><% End If %><li><a href="#" class="haschild">Player Ladders</a><ul><li><a href="/adminops.asp?aType=PMatch">Match</a></li><li><a href="/adminops.asp?aType=PForfeit">Forfeit</a></li><li><a href="/adminops.asp?aType=PLadder">Ladder Admin</a></li><li><a href="/editplayerrank.asp">Player Rank</a></li><li><a href="/edit1v1history.asp">History</a></li><% If bSysAdmin Then %><li><a href="/addplayerladder.asp">Add Player Ladder</a></li><li><a href="/playerladderlist.asp">List Player Ladders</a></li><% End If %></ul></li><li><a href="/leagueadmin.asp" class="haschild">Leagues</a><ul><li><a href="/leagueadmin.asp">General Admin</a></li><li><a href="/reports/leagueidentifierreport.asp">GUID Report</a></li><% If bSysAdmin Then %><li><a href="/leagueadd.asp">Add League</a></li><% If IsSysAdminLevel2() Then %><li><a href="/leagueassignadmin.asp">League Assign Admin</a></li><% End If %><% End If %></ul></li><li><a href="/help/admin">Help/Rules</a></li><li><a href="#" class="haschild">Reports</a><ul><li><a href="/reports/playerforfietreport.asp">Player Forfeit Report</a></li><li><a href="/reports/playerrosterreport.asp">Player Roster Report</a></li><li><a href="/reports/rosterreport.asp">Roster Report</a></li><li><a href="/reports/activity.asp">Ladder Activity</a></li><li><a href="/reports/tournamentrosterreport.asp">Tournament Roster Report</a></li><li><a href="/reports/leaguerosterreport.asp">League Roster Report</a></li></ul></li><li><a href="#" class="haschild">Voting Booth</a><ul><li><a href="/ballot/addballot.asp">Add Ballot</a></li><li><a href="/ballot/activateballot.asp">Activate Ballot</a></li><li><a href="/ballot/results.asp">Old Ballot Results</a></li></ul></li><% If bSysAdmin Then %><li><a href="#" class="haschild">Tournaments</a><ul><li><a href="/tournament/admintournament.asp">General Admin</a></li><li><a href="/tournament/AssignAdmin.asp">Assign Admin</a></li><% If IsSysAdminLevel2() Then %><li><a href="/tournament/createtourny.asp">Add Tournament</a></li><% End If %></ul></li><li><a href="#" class="haschild">Sysadmin Tools</a><ul><% If IsSysAdminLevel2() Then %><li><a href="/menu/">Update Menus</a></li><li><a href="/adminops.asp?aType=Player">Delete / Sysadmin Player </a></li><li><a href="/adminops.asp?aType=Team">Delete Team</a></li><li><a href="/rostertransfer.asp">Transfer Rosters</a></li><% End If %><li><a href="/massmail.asp">Mass Mail</a></li><li><a href="/emailsearch.asp">Email Search</a></li><% If IsSysAdminLevel2() Then %><li><a href="/reports/server_status.asp">Server Status</a></li><% End If %><li><a href="/tracker.asp">IP Tracker</a></li><% If IsSysAdminLevel2() Then %><li><a href="/ipban.asp">IP Banner</a></li><li><a href="/gamelist.asp">Game List</a></li><li><a href="/addgame.asp">New Game</a></li><li><a href="/forums/admin">Forum</a></li><li><a href="/securityaudit.asp">Security Audit</a></li><% End If %></ul><% End If %></ul></li><% Else %><li><a href="javascript:fPopHelp();">help</a></li><% End If %>
	</ul>
</div>