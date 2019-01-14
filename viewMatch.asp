<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle

strPageTitle = "TWL: View Match" 

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

Dim strLadderName, intLadderID
strLadderName = Request.QueryString("Ladder")
If Len(Trim(strLadderName)) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If

Dim intMatchID
intMatchID = Trim(Request.QueryString("MatchID"))
If Not(IsNumeric(intMatchID)) OR Len(intMatchID) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
	
strSQL = "SELECT LadderID, LadderName FROM tbl_ladders WHERE LadderName = '" & CheckString(strLadderName) & "' AND LadderActive = 1"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLadderID = oRs.Fields("LadderID").Value
	strLadderName = oRs.Fields("LadderName").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
oRs.NextRecordSet

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

'' Get match details
Dim strDefenderTeam, strAttackerTeam, strDefenderTeamTags, strAttackerTeamTags, intDefenderTeamID, intAttackerTeamID
DIm intDefenderLinkID, intAttackerLinkID, strMaps(6), i, intDefenderVotes, intAttackerVotes
Dim intDefenderWins, intDefenderLosses, intDefenderForfeits
Dim intAttackerWins, intAttackerLosses, intAttackerForfeits
strSQL = "EXECUTE LadderMatchDetails @LadderMatchID = '" & CheckString(intMatchID) & "'"
oRs.Open strSQL, oConn
If oRs.State = 1 Then
	If Not(oRs.EOF AND oRs.BOF) Then
		strDefenderTeam = oRs.FIelds("DefenderTeamName").Value
		strAttackerTeam = oRs.FIelds("AttackerTeamName").Value
		strDefenderTeamTags = oRs.FIelds("DefenderTeamTag").Value
		strAttackerTeamTags = oRs.FIelds("AttackerTeamTag").Value
		intDefenderTeamID = oRs.FIelds("DefenderTeamID").Value
		intAttackerTeamID = oRs.FIelds("AttackerTeamID").Value
		intDefenderLinkID = oRs.FIelds("DefenderTeamLinkID").Value
		intAttackerLinkID = oRs.FIelds("AttackerTeamLinkID").Value
		intDefenderVotes = oRs.FIelds("DefenderVotes").Value
		intAttackerVotes = oRs.FIelds("AttackerVotes").Value
		intDefenderWins = oRs.FIelds("DefenderWins").Value
		intDefenderLosses = oRs.FIelds("DefenderLosses").Value
		intDefenderForfeits = oRs.FIelds("DefenderForfeits").Value

		intAttackerWins = oRs.FIelds("AttackerWins").Value
		intAttackerLosses = oRs.FIelds("AttackerLosses").Value
		intAttackerForfeits = oRs.FIelds("AttackerForfeits").Value
	End If
	oRs.nextRecordSet
End if

Dim intDefenderPct, intAttackerPct
If intDefenderVotes + intAttackerVotes > 0 Then
	intDefenderPct = (intDefenderVotes / (intDefenderVotes + intAttackerVotes))
	intAttackerPct = (intAttackerVotes / (intDefenderVotes + intAttackerVotes))
Else
	intDefenderPCt = 0
	intAttackerPct = 0
End If
'' find out if user has voted on match
Dim intVotedID
intVotedID = -1
If bLoggedIn Then 
	strSQL = "SELECT TLLinkID FROM lnk_l_p_m_votes WHERE PlayerID = '" & Session("PlayerID") & "' AND MatchID ='" & CheckString(intMatchID) & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		intVotedID = oRs.Fields("TLLinkID").Value
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

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call Content2BoxStart(strLadderName & " Ladder Match Details")
%>	
			<table class="cssBordered" width="100%">
			<tr>
				<th colspan="2" bgcolor="#000000">Defending Team</th>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Team Name:</b></td>
				<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(strDefenderTeam & "")%>"><%=Server.HTMLEncode(strDefenderTeam & "")%></a></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Record:</b></td>
				<td bgcolor="<%=bgcone%>"><%=Server.HTMLEncode("W: " & intDefenderWins & " L: " & intDefenderLosses & " F: " & intDefenderForfeits)%></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Predictions for:</b></td>
				<td bgcolor="<%=bgcone%>"><%=intDefenderVotes & " - " & FormatPercent(intDefenderPct, 2)%></td>
			</tr>
			<% 
			If intVotedID = 0 Then
				%>
				<tr><td colspan="2" bgcolor="#000000" align="center"><a href="saveitem.asp?Ladder=<%=Server.URLEncode(strLadderName & "")%>&SaveType=LadderMatchVote&MatchID=<%=intMatchID%>&VoteFor=<%=intDefenderLinkID%>">Click here to predict a win for <%=Server.HTMLEncode(strDefenderTeam & "")%></a></td></tr>
				<%
			End If
			%>
			</table>

		<% Call Content2BoxMiddle() %>
			<table class="cssBordered" width="100%">
			<tr>
				<th colspan="2" bgcolor="#000000">Attacking Team</th>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Team Name:</b></td>
				<td bgcolor="<%=bgcone%>"><a href="viewteam.asp?team=<%=Server.URLEncode(strAttackerTeam & "")%>"><%=Server.HTMLEncode(strAttackerTeam & "")%></a></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Record:</b></td>
				<td bgcolor="<%=bgcone%>"><%=Server.HTMLEncode("W: " & intAttackerWins & " L: " & intAttackerLosses & " F: " & intAttackerForfeits)%></td>
			</tr>				
			<tr>
				<td bgcolor="<%=bgcone%>" align="right"><b>Predictions for:</b></td>
				<td bgcolor="<%=bgcone%>"><%=intAttackerVotes & " - " & FormatPercent(intAttackerPct, 2)%></td>
			</tr>
			<% 
			If intVotedID = 0 Then
				%>
				<tr><td colspan="2" bgcolor="#000000" align="center"><a href="saveitem.asp?Ladder=<%=Server.URLEncode(strLadderName & "")%>&SaveType=LadderMatchVote&MatchID=<%=intMatchID%>&VoteFor=<%=intAttackerLinkID%>">Click here to predict a win for <%=Server.HTMLEncode(strAttackerTeam & "")%></a></td></tr>
				<%
			End If
			%>
			</table>

<%
Call Content2BoxEnd()
Call ContentStart("Rant Board")
%>
<center>
	
<a href="viewladder.asp?Ladder=<%=Server.URLEncode(strLadderName & "")%>">Jump to <%=Server.HTMLEncode(strLadderName & "")%></a>
<br /><br />
Warning: TeamWarfare staff will not police these rant boards. They are meant<br />
for teams to get into the "spirit" of their upcoming matches. So have fun.<br /><br />
NOTICE: The above rule does <b>NOT</b> allow for a complete disregard of common sense.<br />
Racism, excessive personal attacks, and utter idiocy will reflect badly on your team
and may result in punishment.<br /><br/>
<%
If Request.queryString("E") = "1" Then
	%>
	<font color="#ff0000"><B>You must wait at least 15 seconds between rants.</b></font><br /><br />
	<%
End If
%>
</center>
<%
Dim intPages, intTotalRecords, intContributor
strSQL = "SELECT PlayerHandle, PlayerTitle, Contributor, p.PlayerID, MRID, Rant, RantTime FROM tbl_match_rants mr INNER JOIN tbl_players p ON p.PlayerID = mr.PlayerID WHERE MatchID = '" & CheckString(intMatchID) & "' ORDER BY MRID ASC"
oRS.PageSize = intPerPage
oRS.CacheSize = intPerPage
oRS.CursorLocation = adUseClient
oRs.Open strSQL, oConn, adOpenForwardOnly, adLockReadOnly ', adCmdTableDirect
If (oRs.EOF AND oRs.BOF) Then
	%>
	<center><a href="LadderAddRant.asp?Ladder=<%=Server.URLEnCode(strLadderName & "")%>&MatchID=<%=intMatchID%>&cMode=add">Add Rant</a><br />
	There have been no match rants yet.
</center>
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
	<table border="0" cellspacing="0" width="100%">
	<% Call ListPages(intPageNum, intPages) %>
	</table>
	<table border="0" cellspacing="0" cellpadding="0" width="100%" class="cssBordered">
	<TR bgcolor=<%=strHeaderColor%>>
		<TH CLASS="columnheader" ALIGN="LEFT">Ranter</TH>
		<TH CLASS="columnheader" >
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH=100%>
			<tr>
				<th align="left">Rant</th>
				<th align="right" class="columnheader"><a href="LadderAddRant.asp?Ladder=<%=Server.URLEnCode(strLadderName & "")%>&MatchID=<%=intMatchID%>&cMode=add">Add Rant</a></TH>
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
							Response.Write " / <a href=""LadderAddRant.asp?cmode=edit&mrid=" & oRs.FIeldS("MRID").Value & "&Ladder=" & Server.URLEncode(strLadderName & "") & "&MatchID=" & intMatchID & """>edit</a>"
							Response.Write " / <a href=""SaveItem.asp?SaveType=Delete_Match_Rant&mrid=" & oRs.FIeldS("MRID").Value & "&Ladder=" & Server.URLEncode(strLadderName & "") & "&MatchID=" & intMatchID & """>delete</a>"
							Response.Write " / <a href=""SaveItem.asp?SaveType=Purge_Match_Rant&p=" & oRs.FIeldS("PlayerID").Value & "&Ladder=" & Server.URLEncode(strLadderName & "") & "&MatchID=" & intMatchID & """>purge all</a>"
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
		<TH CLASS="columnheader" COLSPAN=2 ALIGN="RIGHT"><a href="LadderAddRant.asp?Ladder=<%=Server.URLEnCode(strLadderName & "")%>&MatchID=<%=intMatchID%>&cMode=add">Add Rant</a><br /></TH>
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
			Response.Write " <a alt=""First Page"" href=""viewmatch.asp?ladder=" & Server.URLENcode(strLadderName & "") & "&matchid=" & intMatchID & "&page=1"">&laquo; First</A> ... "
		End If
		If iPageNum > 1 Then
			Response.Write " <a alt=""Previous Page"" href=""viewmatch.asp?ladder=" & Server.URLENcode(strLadderName & "") & "&matchid=" & intMatchID &"&page=" & iPageNum - 1 & """>&laquo;</A> "
		End If
		For i = iPageNum - 5 To iPageNum + 5 
			If i > 0 Then
				If i = iPageNum Then
					Response.Write " <span class=""currentpage"">[" & i & "]</span>"
				ElseIf i <= iTotalPages Then
					Response.Write " <a href=""viewmatch.asp?ladder=" & Server.URLENcode(strLadderName & "") & "&matchid=" & intMatchID & "&page=" & i & """>" & i & "</a>"
				End If				
			End If
		Next
		If iPageNum < iTotalPages Then
			Response.Write " <a alt=""Next Page"" href=""viewmatch.asp?ladder=" & Server.URLENcode(strLadderName & "") & "&matchid=" & intMatchID & "&page=" & iPageNum + 1 & """>&raquo;</A> "
		End If
		If iPageNum + 5 < iTotalPages Then
			Response.Write " ... <a alt=""Last Page"" href=""viewmatch.asp?ladder=" & Server.URLENcode(strLadderName & "") & "&matchid=" & intMatchID & "&page=" & iTotalpages & """>Last &raquo;</A>"
		End If
		Response.Write "</B>"
		Response.Write "</TD></TR>"
		Response.Write "<TR><TD><IMG SRC=""/images/spacer.gif"" HEIGHT=5></TD></TR>"
	End If
End Function
%>