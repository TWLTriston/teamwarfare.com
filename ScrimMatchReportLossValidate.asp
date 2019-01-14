<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Validate Loss Report"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
Dim intMatchID, intLinkID
Dim strDefenderName, strAttackerName, intLadderID, strTeamName, intAttackerEloID, intDefenderEloID
Dim strLadderName, strMapArray(6)
Dim i, intMaps
Dim dtmMatchDate

Dim MatchID, TeamID
Dim mLadderName, mMap1, mMap2, mMap3, mMap4, mMap5
Dim mLadderID, mDefenderTeamID, mAttackerTeamID
Dim mBookedDate, mDid, mAid, mMatchDate
Dim mMap1DefScore, mMap2DefScore, mMap3DefScore, mMap4DefScore, mMap5DefScore
Dim mMap1AttScore, mMap2AttScore, mMap3AttScore, mMap4AttScore, mMap5AttScore
Dim mMap1OTWin, mMap2OTWin, mMap3OTWin, mMap4OTWin, mMap5OTWin
Dim mMap1Forfeit, mMap2Forfeit, mMap3Forfeit
Dim dScore, aScore, m1c, m2c, m3c
Dim mWinner1, mWinner2, mWinner3
Dim pScore, mMatchWinnerIsDefender, mWinner, mLoser, mWinnerID, mLoserID
Dim Map1IsOT, Map2IsOT, Map3IsOT
Dim Map1Forfeit, Map2Forfeit, Map3Forfeit

Dim blnIsAWinner, blnMapTie(6), blnMapScored(6), blnMapOT(6), blnMapForfeit(6)
Dim intAttScore, intDefScore
Dim item

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->
<% Call ContentStart("Validate Match Report") %>
<%
intMatchID = Request.Form("matchid")
intLinkID = Request.Form("LinkID")
intLadderID = Request.Form("ladderid")

strSQL = "SELECT 'DefenderName' = d.TeamName, 'AttackerName' = a.TeamName, MatchDate, AttackerEloTeamID, DefenderEloTeamID, Map1, Map2, Map3, l.EloLadderName, l.EloLadderID "
strSQL = strSQL & " FROM tbl_elo_matches em "
strSQL = strSQL & " INNER JOIN tbl_elo_ladders l ON l.EloLadderID = em.EloLadderID "
strSQL = strSQL & " INNER JOIN lnk_elo_team aet ON aet.lnkEloTeamID = em.AttackerEloTeamID "
strSQL = strSQL & " INNER JOIN lnk_elo_team det ON det.lnkEloTeamID = em.DefenderEloTeamID "
strSQL = strSQL & " INNER JOIN tbl_teams a ON a.TeamID = aet.TeamID "
strSQL = strSQL & " INNER JOIN tbl_teams d ON d.TeamID = det.TeamID "
strSQL = strSQL & " WHERE em.EloMatchID = '" & intMatchID & "'"
oRs.Open strSQL, oconn
If Not (oRs.EOF AND oRs.BOF) Then
	strDefenderName = oRs.Fields("DefenderName").Value
	strAttackerName = oRs.Fields("AttackerName").Value
	intAttackerEloID = oRs.Fields("AttackerEloTeamID").Value
	intDefenderEloID = oRs.Fields("DefenderEloTeamID").Value
	strLadderName = oRs.Fields("EloLadderName").Value
	intLadderID = oRs.Fields("EloLadderID").Value
	strMapArray(1) = oRs.Fields("Map1").Value
	strMapArray(2) = oRs.Fields("Map2").Value
	strMapArray(3) = oRs.Fields("Map3").Value
'	dtmMatchDate = oRs.Fields("MatchDate").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.clear
	response.redirect "/errorpage.asp?error=7"
End If
ors.Close

If CInt(intLinkID) = intAttackerEloID THen
	strTeamName = strAttackerName
Else
	strTeamName = strDefenderName
End If

If Not(bSysAdmin or IsEloLadderAdmin(strLadderName) or IsTeamFounder(strTeamName) OR IsEloTeamCaptain(strTeamName, strLadderName) ) Then 
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
end if
dtmMatchDate = Request.Form("matchdate")
intMaps = clng(Request.Form("Maps"))
strMapArray(1) = Request.Form("Map1")
strMapArray(2) = Request.Form("Map2")
strMapArray(3) = Request.Form("Map3")
strMapArray(4) = Request.Form("Map3")
strMapArray(5) = Request.Form("Map3")

intDefScore = 0
intAttScore = 0

blnIsAWinner = False

		' Each map must have a clearly defined winner
		For i = 1 To clng(intMaps)
			' Is this a forfeit win???
			If Len(Request.Form("Map" & i & "Forfeit")) > 0 Then
				' Win by forfeit!
				If Request.Form("Map" & i & "Forfeit") = "Attacker" Then
					intDefScore = intDefScore + 1
				Else
					intAttScore = intAttScore + 1
				End If				
				blnMapScored(i) = True
			ElseIf Len(Request.Form("Map" & i & "OTwin")) > 0 Then
				' Win by OT!!
				If Request.Form("Map" & i & "OTwin") = "Attacker" Then
					intAttScore = intAttScore + 1
				Else
					intDefScore = intDefScore + 1
				End If	
				blnMapScored(i) = True
			ElseIf Len(Request.Form("Map" & i & "AttScore")) > 0 AND Len(Request.Form("Map" & i & "DefScore")) > 0 Then
				' Check scores for a winner
				If clng(Request.Form("Map" & i & "AttScore")) > clng(Request.Form("Map" & i & "DefScore")) Then
					intAttScore = intAttScore + 1
				ElseIf clng(Request.Form("Map" & i & "DefScore")) > clng(Request.Form("Map" & i & "AttScore")) Then
					intDefScore = intDefScore + 1
				Else
					'' This is a tie
					blnMapTie(i) = True
				End If			
				blnMapScored(i) = True
			Else 
				blnMapScored(i) = False
			End If
			If Len(Request.Form("Map" & i & "OTWin")) > 0 Then
				blnMapOT(i) = 1
			Else
				blnMapOT(i) = 0
			End If
			If Len(Request.Form("Map" & i & "Forfeit")) > 0 Then
				If Request.Form("Map" & i & "Forfeit") = "Defender" Then
					intAttScore = intAttScore + 1
				ElseIf Request.Form("Map" & i & "Forfeit") = "Attacker" Then
					intDefScore = intDefScore + 1
				End If
				blnMapForfeit(i) = 1
			Else
				blnMapForfeit(i) = 0
			End If
		Next
		If (intMaps = 0) Then
			blnIsAWinner = True
		ElseIf intDefScore / intMaps > 0.5 Then
			blnIsAWinner = True
		ElseIf intAttScore / intMaps > 0.5 Then
			blnIsAWinner = True
		End If
		If intAttScore < intDefScore Then
	mWinner = strDefenderName
	mLoser = strAttackerName
	pScore = intDefScore & " - " & intAttScore
	mLoserID = intAttackerEloID
	mWinnerID = intDefenderEloID
	mMatchWinnerIsDefender = True
ElseIf intAttScore > intDefScore Then
	mLoserID = intDefenderEloID
	mWinnerID = intAttackerEloID
	mWinner = strAttackerName
	mLoser = strDefenderName
	mMatchWinnerIsDefender = False
	pScore = intAttScore & " - " & intDefScore
Else
	mWinner = strTeamName
End If

%>
<TABLE BORDER=0 CELLSpACING=0 CELLPADDING=0 BGCOLOR="#444444">
<TR><TD>
<table align=center border=0 cellspacing=1 cellpadding=2 width="100%">
	
<%
if mWinner = strTeamName then
	%>
	<tr bgcolor=<%=bgctwo%> height=30><td align=center><%=Server.HTMLEncode (strDefenderName)%> vs. <%=Server.HTMLEncode (strAttackerName)%></td></tr>
	<tr bgcolor=<%=bgcone%> height=30><td align=center>You are not permitted to report as the winner of the match</td></tr>
	<tr bgcolor=<%=bgctwo%> height=30><td align=center><a href="ScrimMatchReportLoss.asp?matchid=<%=intMatchID%>&linkid=<%=intLinkID%>">Click here to return</a></td></tr>
	<%
else
	%>
	<tr bgcolor="#000000"><TH colspan=2><%=Server.HTMLEncode(strDefenderName)%> vs. <%=Server.HTMLEncode(strAttackerName)%> </tH></tr>
	<tr bgcolor=<%=bgcone%> height=30><td align=center colspan=2><b>&nbsp;<%=Server.HTMLEncode(mWinner)%> wins the match <%=pScore%></b></td></tr>
	<TR	BGCOLOR=<%=bgctwo%>><TD ALIGN=CENTER COLSPAN=2>This match is being scored on a per map basis, <BR>each map must have a clearly defined winner.</TD></TR>

	<%	For i = 1 to intMaps %>
		<tr BGCOLOR="#000000"><td colspan=2><img src="/images/spacer.gif" width="1" height="10"></td></tr>
		<tr height=25 bgcolor=<%=bgctwo%>><td colspan=2>&nbsp;<%=Server.HTMLEncode(strMapArray(i))%></td></tr>
		<tr BGCOLOR="#000000"><td colspan=2><img src="/images/spacer.gif" width="1" height="5"></td></tr>
		<tr height=20 bgcolor=<%=bgcone%>><td ALIGN=RIGHT>&nbsp;<%=Server.HTMLEncode(strDefenderName)%> Score:</td><td><%=Request.Form("Map" & i & "DefScore")%></td></tr>
		<tr height=20 bgcolor=<%=bgctwo%>><td ALIGN=RIGHT>&nbsp;<%=Server.HTMLEncode(strAttackerName)%> Score: </td><td><%=Request.Form("Map" & i & "AttScore")%></td></tr>
		<% If Len(Request.Form("Map" & i & "OTWin")) > 0 Then %>
		<tr height=20 bgcolor=<%=bgcone%>><td colspan=2>&nbsp;Map won in over time.</td></tr>
		<% End If %> 
	<% Next %>
	
	<tr BGCOLOR="#000000"><td colspan=2><img src="/images/spacer.gif" width="1" height="9"></td></tr>
	<FORM NAME="frmSaveResults" ACTION="/scrim/SaveItem.asp" METHOD="POST">
	<INPUT TYPE=HIDDEN NAME="SaveType" VALUE="ReportMatch">
	<INPUT TYPE=HIDDEN NAME="matchwinnerdefending" VALUE="<%=cBool(mMatchWinnerIsDefender)%>">
	<INPUT TYPE=HIDDEN NAME="matchlosername" VALUE="<%=Server.HTMLEncode(mLoser)%>">
	<INPUT TYPE=HIDDEN NAME="matchWinnerName" VALUE="<%=Server.HTMLEncode(mWinner)%>">
	<INPUT TYPE=HIDDEN NAME="LadderName" VALUE="<%=Server.HTMLEncode(strLadderName)%>">
	<INPUT TYPE=HIDDEN NAME="MatchLoserID" VALUE="<%=Server.HTMLEncode(mLoserID)%>">
	<INPUT TYPE=HIDDEN NAME="MatchWinnerID" VALUE="<%=Server.HTMLEncode(mWinnerID)%>">
	<% For i = 1 To intMaps %>
	<INPUT TYPE=HIDDEN NAME="intMap<%=i%>OT" VALUE="<%=blnMapOT(i)%>">
	<% Next %>	
	<%
	For Each Item in Request.Form
		Response.Write "<INPUT TYPE=HIDDEN NAME=""" & item & """ VALUE=""" & Server.HTMLEncode(Request.Form(item)) & """>" & vbCrLf
	Next
	%>
	<tr bgcolor=<%=bgcone%> height=30 valign=center><td colspan=2 align=center><input type=button value="Back" onclick="window.location.href='ScrimMatchReportLoss.asp?matchid=<%=intMatchID%>&linkid=<%=intLinkID%>';" class=bright>
	<input type=SUBMIT name=saveresults value="Save Results" class=bright>
	</td></tr>
	</FORM>
<%
end if
%>
</TABLE>
</TD></TR>
</tABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

