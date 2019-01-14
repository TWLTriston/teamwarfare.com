<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("League"), """", "&quot;") 

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo, strHeaderColor
strHeaderColor	= Application("HeaderColor")

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLoggedIn
bLoggedIn = Session("LoggedIn")
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strLeagueName, intLeagueID
strLeagueName = Request.QueryString("League")
If Len(Trim(strLeagueName)) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If

Dim intLeagueMatchID
intLeagueMatchID = Trim(Request.QueryString("LeagueMatchID"))
If Not(IsNumeric(intLeagueMatchID)) OR Len(intLeagueMatchID) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
	
strSQL = "SELECT LeagueID, LeagueName FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "' AND LeagueActive = 1"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
	strLeagueName = oRs.Fields("LeagueName").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
oRs.NextRecordSet'' Get match details
Dim strHomeTeam, strVisitorTeam, strHomeTeamTags, strVisitorTeamTags, intHomeTeamID, intVisitorTeamID
DIm intHomeLinkID, intVisitorLinkID, strMaps(6), i, intHomeVotes, intVisitorVotes
Dim intHomeWins, intHomeLosses, intHomeDraws, intHomeNoShows, intHomeRoundsWon, intHomeRoundsLost, intHomeLeaguePoints, intHomeWinPct
Dim intVisitorWins, intVisitorLosses, intVisitorDraws, intVisitorNoShows, intVisitorRoundsWon, intVisitorRoundsLost, intVisitorLeaguePoints, intVisitorWinPct
strSQL = "EXECUTE LeagueMatchDetails @LeagueMatchID = '" & CheckString(intLeagueMatchID) & "'"
oRs.Open strSQL, oConn
If oRs.State = 1 Then
	If Not(oRs.EOF AND oRs.BOF) Then
		strHomeTeam = oRs.FIelds("HomeTeamName").Value
		strVisitorTeam = oRs.FIelds("VisitorTeamName").Value
		strHomeTeamTags = oRs.FIelds("HomeTeamTag").Value
		strVisitorTeamTags = oRs.FIelds("VisitorTeamTag").Value
		intHomeTeamID = oRs.FIelds("HomeTeamID").Value
		intVisitorTeamID = oRs.FIelds("VisitorTeamID").Value
		intHomeLinkID = oRs.FIelds("HomeTeamLinkID").Value
		intVisitorLinkID = oRs.FIelds("VisitorTeamLinkID").Value
		intHomeVotes = oRs.FIelds("HomeVotes").Value
		intVisitorVotes = oRs.FIelds("VisitorVotes").Value
		intHomeWins = oRs.FIelds("HomeWins").Value
		intHomeLosses = oRs.FIelds("HomeLosses").Value
		intHomeDraws = oRs.FIelds("HomeDraws").Value
		intHomeNoShows = oRs.FIelds("HomeNoShows").Value
		intHomeRoundsWon = oRs.FIelds("HomeRoundsWon").Value
		intHomeRoundsLost = oRs.FIelds("HomeRoundsLost").Value
		intHomeLeaguePoints = oRs.FIelds("HomeLeaguePoints").Value
		intHomeWinPct = oRs.FIelds("HomeWinPct").Value
		intVisitorWins = oRs.FIelds("VisitorWins").Value
		intVisitorLosses = oRs.FIelds("VisitorLosses").Value
		intVisitorDraws = oRs.FIelds("VisitorDraws").Value
		intVisitorNoShows = oRs.FIelds("VisitorNoShows").Value
		intVisitorRoundsWon = oRs.FIelds("VisitorRoundsWon").Value
		intVisitorRoundsLost = oRs.FIelds("VisitorRoundsLost").Value
		intVisitorLeaguePoints = oRs.FIelds("VisitorLeaguePoints").Value
		intVisitorWinPct = oRs.FIelds("VisitorWinPct").Value
	End If
	oRs.nextRecordSet
End if

Dim intHomePct, intVisitorPct
If intHomeVotes + intVisitorVotes > 0 Then
	intHomePct = (intHomeVotes / (intHomeVotes + intVisitorVotes))
	intVisitorPct = (intVisitorVotes / (intHomeVotes + intVisitorVotes))
Else
	intHomePCt = 0
	intVisitorPct = 0
End If
'' find out if user has voted on match
Dim intVotedID
intVotedID = -1
If bLoggedIn Then 
	strSQL = "SELECT lnkLeagueTeamID FROM lnk_match_player_votes WHERE PlayerID = '" & Session("PlayerID") & "' AND LeagueMatchID ='" & CheckString(intLeagueMatchID) & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intVotedID = oRs.Fields("lnkLeagueTeamID").Value
	Else 
		intVotedID = 0
	End If
	oRs.NextRecordSet
End If

Dim bgc
Dim intTimeZoneDifference, strDate, strTime
Dim strCurrentTime, strCurrentDate
Dim strDateMask, bln24HourTime

intTimeZoneDifference = Session("intTimeZoneDifference")
strDateMask = "MM-DD-YYYY"
bln24HourTime = False
Dim intPerPage, intPageNum

intPageNum = Request.querystring("page")
If Len(intPageNum) = 0 Or  Not(IsNumeric(intPageNum)) then
	intPageNum = 1
Else
	intPageNum = cInt(intPageNum)
End If

intPerPage = Request.querystring("PerPage")
If Len(intPerPage) = 0 Or Not(IsNumeric(intPerPage)) then
	intPerPage = 20
Else
	intPerPage = cInt(intPerPage)
End If

Dim intCurrent
intCurrent = 0
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call Content2BoxStart(strLeagueName & " League Match Details")
%>	
		<table border="0" cellspacing="0" cellpadding="0" width="97%" align="center" bgcolor="#444444">
		<tr>
			<td>
			<table border="0" cellspacing="1" cellpadding="4" width="100%">
			<tr>
				<th colspan="2" bgcolor="#000000">Home Team</th>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Team Name:</b></td>
				<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(strHomeTeam & "")%>"><%=Server.HTMLEncode(strHomeTeam & "")%></a></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgctwo%>" align="right"><b>League Points:</b></td>
				<td bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(intHomeLeaguePoints & "")%></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Match Record:</b></td>
				<td bgcolor="<%=bgcone%>"><%=Server.HTMLEncode("W: " & intHomeWins & " L: " & intHomeLosses & " D: " & intHomeDraws & " F: " & intHomeNoShows)%></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgctwo%>" align="right"><b>Round / Map Record:</b></td>
				<td bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode("W: " & intHomeRoundsWon & " L: " & intHomeRoundsLost)%></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Predictions for:</b></td>
				<td bgcolor="<%=bgcone%>"><%=intHomeVotes & " - " & FormatPercent(intHomePct, 2)%></td>
			</tr>
			<% 
			If intVotedID = 0 Then
				%>
				<tr><td colspan="2" bgcolor="#000000" align="center"><a href="saveitem.asp?League=<%=Server.URLEncode(strLeagueName & "")%>&SaveType=LeagueMatchVote&MatchID=<%=intLeagueMatchID%>&VoteFor=<%=intHomeLinkID%>">Click here to predict a win for <%=Server.HTMLEncode(strHomeTeam & "")%></a></td></tr>
				<%
			End If
			%>
			</table>
			</td>
		</tr>
		</table>
<% Call Content2BoxMiddle() %>

		<table border="0" cellspacing="0" cellpadding="0" width="97%" align="center" bgcolor="#444444">
		<tr>
			<td>
			<table border="0" cellspacing="1" cellpadding="4" width="100%">
			<tr>
				<th colspan="2" bgcolor="#000000">Visitor Team</th>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Team Name:</b></td>
				<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(strVisitorTeam & "")%>"><%=Server.HTMLEncode(strVisitorTeam & "")%></a></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgctwo%>" align="right"><b>League Points:</b></td>
				<td bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(intVisitorLeaguePoints & "")%></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Match Record:</b></td>
				<td bgcolor="<%=bgcone%>"><%=Server.HTMLEncode("W: " & intVisitorWins & " L: " & intVisitorLosses & " D: " & intVisitorDraws & " F: " & intVisitorNoShows)%></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgctwo%>" align="right"><b>Round / Map Record:</b></td>
				<td bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode("W: " & intVisitorRoundsWon & " L: " & intVisitorRoundsLost)%></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Predictions for:</b></td>
				<td bgcolor="<%=bgcone%>"><%=intVisitorVotes & " - " & FormatPercent(intVisitorPct, 2)%></td>
			</tr>
			<% 
			If intVotedID = 0 Then
				%>
				<tr><td colspan="2" bgcolor="#000000" align="center"><a href="saveitem.asp?League=<%=Server.URLEncode(strLeagueName & "")%>&SaveType=LeagueMatchVote&MatchID=<%=intLeagueMatchID%>&VoteFor=<%=intVisitorLinkID%>">Click here to predict a win for <%=Server.HTMLEncode(strVisitorTeam & "")%></a></td></tr>
				<%
			End If
			%>
			</table>
			</td>
		</tr>
		</table>

<%
Call Content2BoxEnd()
Call ContentStart("Rant Board")
%>
<div align="left">&nbsp;&nbsp;&nbsp;<a href="viewleaguematches.asp?league=<%=Server.URLEncode(strLEagueName & "")%>">Back to league matches</a></div>
<br /><br />
Warning: TeamWarfare staff will not police these rant boards. They are meant<br />
for teams to get into the "spirit" of their upcoming matches. So have fun.<br /><br />
<%
Dim intPages, intTotalRecords, intContributor
strSQL = "SELECT PlayerHandle, PlayerTitle, p.PlayerID, Contributor, LMRID, Rant, RantTime FROM tbl_league_match_rants lmr INNER JOIN tbl_players p ON p.PlayerID = lmr.PlayerID WHERE LeagueMatchID = '" & CheckString(intLeagueMatchID) & "' ORDER BY LMRID DESC"
' strSQL = "SELECT PlayerHandle, PlayerTitle, Contributor, p.PlayerID, MRID, Rant, RantTime FROM tbl_match_rants mr INNER JOIN tbl_players p ON p.PlayerID = mr.PlayerID WHERE MatchID = '" & CheckString(intMatchID) & "' ORDER BY MRID ASC"
oRS.PageSize = intPerPage
oRS.CacheSize = intPerPage
oRS.CursorLocation = adUseClient
oRs.Open strSQL, oConn, adOpenForwardOnly, adLockReadOnly ', adCmdTableDirect
If (oRs.EOF AND oRs.BOF) Then
	%>
	<a href="LeagueAddRant.asp?League=<%=Server.URLEnCode(strLeagueName & "")%>&LeagueMatchID=<%=intLeagueMatchID%>&cMode=add">Add Rant</a><br />
	There have been no match rants yet.
	<%
Else
	intPages		= oRS.PageCount
	intTotalRecords		= oRS.RecordCount 
	If intPageNum <= intPages Then
		oRS.AbsolutePage	= intPageNum
	Else
		oRs.AbsolutePage = 1
		intPageNum = 1
	End If
	%>
	<table border="0" cellspacing="0" width="97%">
	<% Call ListPages(intPageNum, intPages) %>
	</table>
	<table border="0" cellspacing="0" cellpadding="0" width="97%" class="cssBordered"
	<TR bgcolor=<%=strHeaderColor%>>
		<TH CLASS="columnheader" ALIGN="LEFT">Ranter</TH>
		<TH CLASS="columnheader" >
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH=100%>
			<tr>
				<th align="left">Rant</th>
				<th align="right" class="columnheader"><a href="LeagueAddRant.asp?League=<%=Server.URLEnCode(strLeagueName & "")%>&LeagueMatchID=<%=intLeagueMatchID%>&cMode=add">Add Rant</a></TH>
			</tr>
			</table>
		</th>
	</tr>
	<%
	Do While Not(oRS.EOF) AND oRs.AbsolutePage = intPageNum
		intContributor = oRs.Fields("Contributor").Value
		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		%>
		<tr>
			<td valign="top" bgcolor="<%=bgc%>" width="15%"><b><%=Server.HTMLEncode(oRs.Fields("PlayerHandle").Value & "")%></b><br /><span class="usertitle"><%=oRs.Fields("PlayerTitle").Value%>
			<%
			If intContributor = 1 Then
				Response.Write "<br /><a href=""/contributors.asp"">TWL Contributor</a>"
			End If
			%>
			
			</span></td>
			<td bgcolor="<%=bgc%>">
				<%
					Call FixDate(oRs.Fields("RantTime").Value, intTimeZoneDifference, strDate, strTime, strDateMask, bln24HourTime)
					Response.Write "<span class=""smalldate"">" & strDate & "</span> <span class=""smalltime"">" & strTime & "</span>"
				%> <span class="postoptions"> / <a href="/viewplayer.asp?player=<%=server.URLEncode(oRs.FieldS("PlayerHandle").Value)%>">profile</a> <%
						If bSysAdmin Then
							Response.Write " / <a href=""LeagueAddRant.asp?cmode=edit&lmrid=" & oRs.FIeldS("LMRID").Value & "&League=" & Server.URLEncode(strLeagueName & "") & "&LeagueMatchID=" & intLeagueMatchID & """>edit</a>"
							Response.Write " / <a href=""SaveItem.asp?SaveType=Delete_Rant&lmrid=" & oRs.FIeldS("LMRID").Value & "&League=" & Server.URLEncode(strLeagueName & "") & "&LeagueMatchID=" & intLeagueMatchID & """>delete</a>"
						End If
					%></span>
				<br />
				<hr class="forum" />
				<%=Replace(oRs.Fields("Rant").Value, chr(13), "<br />")%>
				<hr class="forum" />
			</td>
		</tr>
		<%
		oRS.MoveNext
	Loop
	%>
	<TR bgcolor=<%=strHeaderColor%>>
		<TH CLASS="columnheader" COLSPAN=2 ALIGN="RIGHT"><a href="LeagueAddRant.asp?League=<%=Server.URLEnCode(strLeagueName & "")%>&LeagueMatchID=<%=intLeagueMatchID%>&cMode=add">Add Rant</a><br /></TH>
	</tr>
	</table>
	<table border="0" cellspacing="0" width="97%">
	<% Call ListPages(intPageNum, intPages) %>
	</table>
	<%
End If
%>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing

Function ListPages(byVal iPageNum, byVal iTotalPages)
	Dim i
	If iTotalPages > 1 Then
		Response.Write "<TR><TD><IMG SRC=""/images/spacer.gif"" HEIGHT=5></TD></TR>"
		Response.Write "<TR>"
		Response.Write "<TD CLASS=""pagelist"">"
		Response.Write "Pages (" & iTotalPages & "): <B>"
		If iPageNum > 5 Then
			Response.Write " <a alt=""First Page"" href=""viewleaguematch.asp?league=" & Server.URLENcode(strLeagueName & "") & "&LeagueMatchID=" & intLeagueMatchID & "&page=1"">&laquo; First</A> ... "
		End If
		If iPageNum > 1 Then
			Response.Write " <a alt=""Previous Page"" href=""viewleaguematch.asp?league=" & Server.URLENcode(strLeagueName & "") & "&LeagueMatchID=" & intLeagueMatchID & "&page=" & iPageNum - 1 & """>&laquo;</A> "
		End If
		For i = iPageNum - 5 To iPageNum + 5 
			If i > 0 Then
				If i = iPageNum Then
					Response.Write " <span class=""currentpage"">[" & i & "]</span>"
				ElseIf i <= iTotalPages Then
					Response.Write " <a href=""viewleaguematch.asp?league=" & Server.URLENcode(strLeagueName & "") & "&LeagueMatchID=" & intLeagueMatchID & "&page=" & i & """>" & i & "</a>"
				End If				
			End If
		Next
		If iPageNum < iTotalPages Then
			Response.Write " <a alt=""Next Page"" href=""viewleaguematch.asp?league=" & Server.URLENcode(strLeagueName & "") & "&LeagueMatchID=" & intLeagueMatchID & "&page=" & iPageNum + 1 & """>&raquo;</A> "
		End If
		If iPageNum + 5 < iTotalPages Then
			Response.Write " ... <a alt=""Last Page"" href=""viewleaguematch.asp?league=" & Server.URLENcode(strLeagueName & "") & "&LeagueMatchID=" & intLeagueMatchID & "&page=" & iTotalpages & """>Last &raquo;</A>"
		End If
		Response.Write "</B>"
		Response.Write "</TD></TR>"
		Response.Write "<TR><TD><IMG SRC=""/images/spacer.gif"" HEIGHT=5></TD></TR>"
	End If
End Function
%>
