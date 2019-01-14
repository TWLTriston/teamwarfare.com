<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Team League Administration"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
Dim blnFirstItem
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strTeamTag, strTeamName, strLeagueName, intTeamLinkID, intPlayerID, intTeamFounderID
Dim intTeamID, intLeagueID, strInfo, intMatchID
strTeamName = Request.QueryString("team")
strLeagueName = Request.QueryString("league")

intPlayerID = Session("PlayerID")

strSQL = "SELECT LeagueID, LeagueName FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
	strLeagueName = oRs.Fields("LeagueName").Value
Else
	oRs.Close
	oConn.Close
	Set oRs = Nothing
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
oRs.NextRecordSet
bTeamFounder = IsTeamFounder(strTeamName)
bLeagueAdmin = IsLeagueAdminByID(intLeagueID)
strSQL = "SELECT TeamID, TeamFounderID, TeamName, TeamTag FROM tbl_teams WHERE TeamName= '" & CheckString(strTeamName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intTeamID = oRs.Fields("TeamID").Value
	intTeamFounderID = oRs.Fields("TeamFounderID").Value
	strTeamName = oRs.Fields("TeamName").Value
	strTeamTag = oRs.Fields("TeamTag").Value
Else
	oRs.Close
	oConn.Close
	Set oRs = Nothing
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
oRs.NextRecordSet

strSQL = "SELECT lnkLeagueTeamID FROM lnk_league_team WHERE LeagueID='" & intLeagueID & "' AND TeamID='" & intTeamID & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intTeamLinkID = oRs.Fields("lnkLeagueTeamID").Value
Else
	oRs.Close
	oConn.Close
	Set oRs = Nothing
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
oRs.NextRecordSet

bTeamCaptain = IsLeagueTeamCaptainByID(intTeamID, intLeagueID)
If Not(bSysAdmin OR bTeamFounder OR bLeagueAdmin OR bTeamCaptain) Then
	oConn.Close
	Set oRs = Nothing
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->
<% Call ContentStart("Team League Administration") %>
<% If bSysAdmin Or bLeagueAdmin Then %>
<script language="javascript" type"text/javascript">
<!-- 
<% If bSysAdmiN Then %>
function PopChange() {
	var objChangeRecord = window.open("TeamLeagueRecordChange.asp?team=<%=Server.URLEncode(strTeamName)%>&league=<%=Server.URLEncode(strLeagueName)%>", "RecordChange",  "width=375,height=300,toolbar=0,scrollbars=0,status=0,location=0,menubar=0,resizable=0");
	objChangeRecord.focus();
}
<% End If %>
function DeleteLeagueMatch(intMatchID) {
	if (confirm("Are you sure you want to delete this match?")) {
		window.location = "saveitem.asp?league=<%=Server.URLEncode(strLeagueName)%>&team=<%=Server.URLEncode(strTeamName)%>&savetype=LeagueDeleteMatch&matchid="+intMatchID;
	}
}
//-->
</script>
<% End If %>
<%
if (bsysadmin) Then
	Response.write "<br /><center><a href=""javascript:PopChange();"">Change this team's record</a></center><br />"
End If
%>
<table border="0" cellspacing="0" cellpadding="0" width="97%" bgcolor="#444444" align="center">
<tr><td><table border="0" cellspacing="1" cellpadding="0" width="100%">
<tr>
	<th colspan="8" bgcolor="#000000">Pending Matches</th>
</tr>
<%
strSQL = "EXECUTE LeagueTeamMatches @LeagueTeamID = '" & intTeamLinkID & "'"
oRs.Open strSQL, oConn
If (oRs.State = 1) Then
	If Not(oRs.EOF AND oRs.BOF) Then
		%>
		<tr>
			<th bgcolor="#000000">Date</th>
			<th bgcolor="#000000">Opponent</th>
			<th bgcolor="#000000">Status</th>
			<th bgcolor="#000000">Maps</th>
			<th bgcolor="#000000">Comms</th>
			<th bgcolor="#000000">Last Comm</th>
			<th bgcolor="#000000">Details</th>
			<% If bLeagueAdmin or bSysAdmin Then %>
			<th bgcolor="#000000">Delete</th>
			<% End If %>
		</tr>
		<%
		blnFirstItem = True
		Do While Not (oRs.EOF)
			%>
			<tr>
				<td valign="top" align="center" bgcolor="<%=bgcone%>"><%=FormatDateTime(oRs.FIelds("MatchDate").Value, 2)%></td>
				<td valign="top"  bgcolor="<%=bgctwo%>"><b><%
					If oRs.Fields("LeagueDivisionID").Value <> 0 Then
						Response.Write "Division"
					ElseIf oRs.Fields("LeagueConferenceID").Value <> 0 Then
						Response.Write "Conference"
					Else
						Response.Write "League"
					End If
					%></b> &raquo; <a href="viewteam.asp?team=<%=Server.URLEncode(oRs.Fields("OpponentName").Value)%>"><%=Server.HTMLEncode(oRs.Fields("OpponentName").Value)%></a></td>
				<% If cLng(oRs.Fields("HomeTeamLinkID").Value) = intTeamLinkID Then %>
				<td valign="top" align="center" bgcolor="<%=bgcone%>">Home</td>
				<% Else %>
				<td valign="top" align="center" bgcolor="<%=bgcone%>">Visitor</td>
				<% End If %>
				<td valign="top"  bgcolor="<%=bgctwo%>" align="center">
					<%
					If Len(oRs.Fields("Map1").Value) > 0 Then
						Response.Write oRs.Fields("Map1").Value
					End If
					If Len(oRs.Fields("Map2").Value) > 0 Then
						Response.Write "<br /> " & oRs.Fields("Map2").Value
					End If
					If Len(oRs.Fields("Map3").Value) > 0 Then
						Response.Write "<br /> " & oRs.Fields("Map3").Value
					End If
					If Len(oRs.Fields("Map4").Value) > 0 Then
						Response.Write "<br /> " & oRs.Fields("Map4").Value
					End If
					If Len(oRs.Fields("Map5").Value) > 0 Then
						Response.Write "<br /> " & oRs.Fields("Map5").Value
					End If
					%></td>
				<td valign="top"  align="center" bgcolor="<%=bgctwo%>"><%=oRs.FIelds("CommsCount").Value%></td>
				<td valign="top"  align="center" bgcolor="<%=bgcone%>"><%
					if Not(IsNull(oRs.Fields("LastCommDate").Value)) Then
						Response.Write oRs.Fields("LastCommDate").Value & "<br />" & Server.HTMLEncode(oRs.Fields("LastCommAuthor").Value)
					Else
						Response.Write "none"
					End If
					%></td>
				<% If Request.QueryString("MatchID") = cStr(oRs.Fields("LeagueMatchID").Value & "") OR (blnFirstItem AND len(Request.QueryString("MatchID") = 0)) Then %>
				<td valign="top" align="center" bgcolor="<%=bgctwo%>"><a href="teamleagueadmin.asp?league=<%=Server.URLEncode(strLeagueName)%>&team=<%=Server.URLEncode(strTeamName)%>">Hide</a></td>
				<% Else %>
				<td valign="top" align="center" bgcolor="<%=bgctwo%>"><a href="teamleagueadmin.asp?league=<%=Server.URLEncode(strLeagueName)%>&team=<%=Server.URLEncode(strTeamName)%>&matchid=<%=oRs.Fields("LeagueMatchID")%>">View</a></td>
				<% End If %>
				<% If bSysAdmin or bLeagueAdmin Then %>
				<td valign="top" align="center" bgcolor="<%=bgcone%>"><a href="javascript:DeleteLeagueMatch(<%=oRs.Fields("LeagueMatchID").Value%>)">Delete</a></td>
				<% End If %>
			</tr>
			<%
			If Request.QueryString("MatchID") = cStr(oRs.Fields("LeagueMatchID").Value & "") OR (blnFirstItem AND Len(Request.QueryString("MatchID")) = 0) Then
				' Match Details..
				intMatchID = oRs.Fields("LeagueMatchID").Value
				%>
				<tr>
					<td colspan="8" bgcolor="#000000" align="center">
						<br />
						<%
						if bTeamCaptain or bTeamFounder then
							Response.Write "<br><center><input type=button name=matchcomm value=""Add Match Communication"" onclick=""window.location.href='leaguematchcomms.asp?matchid=" & intMatchID & "&mode=add&tag=" & server.urlencode(strTeamTag) & "&league=" & server.urlencode(strLeagueName) & "&team=" & server.urlencode(strTeamName) & "';"">"
						else
							Response.Write "<br><center><input type=button name=matchcomm value=""Add Match Communication"" onclick=""window.location.href='leaguematchcomms.asp?matchid=" & intMatchID & "&mode=add&tag=TWLAdmin&league=" & server.urlencode(strLeagueName) & "&team=" & server.urlencode(strTeamName) & "';"">"
						end if
						strSQL = "select * from tbl_league_comms where LeagueMatchID=" & intMatchID & " ORDER BY LeagueCommID desc"
						ors2.Open strSQL, oconn
						%>
						<table align=center width=580 border=0 cellspacing=0 cellpadding=1>
						<%
						bgc=bgcone
						if not (ors2.EOF and ors2.bof) then
							do while not ors2.EOF
								Response.Write "<tr bgcolor="& bgc & "><td colspan=2><hr></td></tr><tr bgcolor="& bgc & "><td>Author: <b>" & oRs2.Fields("CommAuthor").Value & " - Posted: " & FormatDateTime(oRs2.Fields("CommDate").Value, 0) & "</td>"
								if bSysAdmin then
									Response.Write "<td align=right><a href=leaguematchcomms.asp?matchid=" & intMatchID & "&mode=edit&league=" & server.urlencode(strLeagueName) & "&team=" & server.urlencode(strTeamName) & "&leaguecommid=" & oRs2.Fields("LeagueCommID").Value & ">Edit</a> - <a href=SaveItem.asp?leaguecommid=" & oRs2.Fields("LeagueCommID").Value & "&matchid=" & intMatchID & "&SaveType=DeleteLeagueCommunications&league=" & server.urlencode(strLeagueName) & "&team=" & server.urlencode(strTeamName) & ">Delete&nbsp;&nbsp;</a></td></tr>"
								else
									Response.Write "<td>&nbsp;</td></tr>"
								end if
								Response.write "<tr bgcolor="& bgc &"><td colspan=2>&nbsp;" & oRs2.Fields("Comms").Value & "</td></tr>"
								if bgc = bgcone then
									bgc=bgctwo
								else
									bgc=bgcone
								end if
								oRs2.MoveNext
							loop
						end if
						oRs2.Close 
						%>
						</table>
						<br />
						<% If cDate(oRs.Fields("MatchDate").Value) <= cDate(Now() + 1) OR bSysAdmin Then %>
						<a href="DisputeMatchLeague.asp?MatchId=<%=intMatchID%>&DisputeTeamID=<%=intTeamLinkID%>&League=<%=Server.URLEncode(strLeagueName)%>">Dispute Match</a>
						<% End If %>
						<br />
						<% If cDate(oRs.Fields("MatchDate").Value) <= cDate(Now() + 1) OR bSysAdmin Then %>
						<a href="leaguereportmatch.asp?matchid=<%=intMatchID%>&f=<%=Server.URLEncode("teamleagueadmin.asp?league=" & Server.URLEncode(strLeagueName) & "&team=" & Server.URLEncode(strTeamName))%>">Report Match Results</a>
						<% End If %>
						<br /><br />
					</td>
				</tr>
				<%
			End If
			oRs.MoveNext
			blnFirstItem = False
		Loop
	Else
		%>
		<tr>
			<td colspan="8" bgcolor="<%=bgctwo%>">No pending matches scheduled.</td>
		</tr>
		<%
	End If
	oRs.NextRecordSet
End If
%>
</table></td></tr>
</table>
<%
Call ContentEnd()

Call Content3BoxStart("League Captain Management")
strSQL = "select p.playerhandle, p.PlayeriD from tbl_players p inner join lnk_league_team_player lnk on lnk.PlayerID=p.playerid where (lnk.lnkLeagueTeamID='" & intTeamLinkID & "' and lnk.isadmin=1) ORDER BY p.playerhandle"
Response.Write "<table border=0 align=center width=97% cellspacing=0><tr height=30 bgcolor="&bgcone&"><td align=center><b>Current Captains</b></td></tr>"
ors.Open strSQL, oconn
if not (ors.EOF and ors.BOF) then
	bgc=bgcone
	do while not ors.EOF
		strInfo=""
		if ors.fields(1).value = intTeamFounderID then 
			strInfo=" (founder)"
		end if	
		if bgc=bgctwo then
			bgc=bgcone
		else bgc=bgctwo
		end if
		Response.Write "<tr height=30 bgcolor="&bgc&"><td align=center>" & Server.HTMLEncode(ors.Fields(0).Value) & strInfo & "</td></tr>"
		ors.MoveNext			
	loop
end if
response.write "</table>"
ors.close

Call Content3BoxMiddle1()

strSQL = "select p.playerhandle, lnk.lnkLeagueTeamPlayerID from tbl_players p inner join lnk_league_team_player lnk on lnk.PlayerID=p.playerid where (lnk.lnkLeagueTeamID='" & intTeamLinkID & "' and lnk.isadmin=0) ORDER BY p.playerhandle"
ors.Open strSQL, oconn
if not(ors.EOF and ors.BOF) then
	Response.Write "<form name=promote action=saveitem.asp method=post><table align=center border=0 width=97% cellspacing=0>"
	Response.Write "<tr bgcolor="&bgcone&" height=30><td align=center><b>Promote Player to Captain</b></td></tr>"
	Response.Write "<tr bgcolor="&bgctwo&"><td height=30 align=center><select name=playerlist style='width:150'>"
	do while not ors.EOF
		Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value)
		ors.MoveNext
	loop
	Response.Write "</select></td></tr><tr height=30 bgcolor="&bgcone&"><td align=center>"
	%>
	<input class=bright type=submit value="Promote">
	<input type=hidden name=SaveType value="PromoteLeagueCaptain">
	<input type=hidden name=League value="<%=Server.HTMLEncode(strLeagueName)%>">
	<input type=hidden name=Team value="<%=Server.HTMLEncode(strTeamName)%>">
	</td></tr></table></form>
	<%
end if
ors.Close

Call Content3BoxMiddle2()

strSQL = "select p.playerhandle, lnk.lnkLeagueTeamPlayerID, p.PlayerID from tbl_players p inner join lnk_league_team_player lnk on lnk.PlayerID=p.playerid where (lnk.lnkLeagueTeamID='" & intTeamLinkID & "' and lnk.isadmin=1) ORDER BY p.playerhandle"
ors.Open strSQL, oconn
if not(ors.EOF and ors.BOF) then
	Response.Write "<form name=demote action=saveitem.asp method=post><table align=center border=0 width=97% cellspacing=0><tr bgcolor="&bgcone&" height=30><td align=center><b>Demote Captain</b></td></tr>"
	Response.Write "<tr bgcolor="&bgctwo&" height=30><td align=center><select name=playerlist style='width:150'>"
	do while not ors.EOF
		if ors.fields(2).value <> intTeamFounderID then
			Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value)
		end if
		ors.MoveNext
	loop
	Response.Write "</select></td></tr><tr bgcolor="&bgcone&" height=30><td align=center><input class=bright type=submit id=submit2 name=submit2 value=Demote><input type=hidden name=SaveType value=DemoteLeagueCaptain><input type=hidden name=League value=""" & Server.HTMLEncode(strLeagueName) & """><input type=hidden name=team value=""" & Server.HTMLEncode(strTeamName) & """></td></tr></table></form>"	
End If
ors.Close

Call Content3BoxEnd()
Call ContentStart(Server.HTMLEncode(strLeagueName) & "League Roster Management")

	strSQL = "select p.playerhandle, lnk.lnkLeagueTeamPlayerID, p.PlayerID from tbl_players p inner join lnk_league_team_player lnk on lnk.PlayerID=p.playerid where (lnk.lnkLeagueTeamID='" & intTeamLinkID & "') ORDER BY p.playerhandle"
	ors.open strsql,oconn
	if not (ors.eof and ors.bof) then
		response.write "<form name=BootPlayer method=post action=saveitem.asp><table width=50% align=center border=0 cellspacing=0 cellpadding=0>"
		Response.Write "<input type=hidden name=League value=""" & Server.HTMLEncode(strLeagueName) & """><input type=hidden name=team value=""" & Server.HTMLEncode(strTeamName) & """>"	
		response.write "<tr bgcolor="&bgcone&" height=125 valign=center><td align=center><input type=hidden name=savetype value=""DropLeaguePlayer""><select name=PlayerID size=5 class=brightred style='width:200'>"
		do while not ors.eof
			if ors.fields(2).value <> intTeamFounderID then 
				response.write "<option value=" & ors.fields(1).value & ">" & Server.HTMLEncode(ors.fields(0).value) & "</option>" & vbCrLf
			end if
			ors.movenext
		loop
		response.write "</select></td></tr><tr bgcolor="&bgctwo&" height=35><td align=center><input type=hidden name=link value=" & intTeamLinkID & "><input type=submit class=bright style='width:75' value='Kick Player'></td></tr></table></form>"
	end if
	ors.close

Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>