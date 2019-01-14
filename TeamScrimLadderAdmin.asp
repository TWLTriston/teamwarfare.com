	<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Team Scrim Ladder Administration"

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strTeamName, strLadderName
strTeamName = Request.QueryString("team")
strLadderName = Request.QueryString("ladder")

bTeamFounder = IsTeamFounder(strTeamName)
bTeamCaptain = IsEloTeamCaptain(strTeamName, strLadderName)
bLadderAdmin = IsEloLadderAdmin(strLadderName)

Dim intTeamID, strTeamTag, intLadderID
Dim blnLocked, intMinPlayer, intMaxRatingDiff
Dim intPlayerID, intEloTeamID, dtmLoginTime
Dim intRating, intFounderID, intMatchID
intPlayerID = Session("PlayerID")

intMatchID = Request.QueryString("MatchID")
If Len(intMatchID) = 0 OR Not(IsNumeric(intMatchID)) Then
	intMatchID = 0
Else
	intMatchID = CDbl(intMatchID)
End If

'' Challenge Variables
Dim strOpponentName, strOpponentTag, intOpponentRanking
Dim bgc2
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->
<%
strSQL = "SELECT teamid, teamtag, teamfounderid from tbl_teams where teamname='" & CheckString(strTeamName) & "'"
ors.Open strSQL, oconn
if not (ors.eof and ors.BOF) then
	intTeamID = ors.Fields("TeamID").Value
	strTeamTag = ors.fields("TeamTag").value
	intFounderID = oRS.Fields("TeamFounderID").Value
Else
	oRs.Close
	oConn.Close 
	Set oRS = Nothing
	Set oConn = Nothing
	Set oRs2 = Nothing
	response.clear
	response.redirect "/errorpage.asp?error=7"
end if
oRs.NextRecordset 

strSQL = "select Eloladderid, EloLocked, EloMaxRatingDiff, EloMinPlayer FROM tbl_elo_ladders where Eloladdername='" & CheckString(strLadderName) & "'"
ors.Open strSQL, oconn
if not (ors.eof and ors.BOF) then
	intLadderID =ors.Fields("EloLadderID").Value
	blnLocked = CBool(ors.fields("EloLocked").value)
	intMaxRatingDiff = ors.fields("EloMaxRatingDiff").value
	intMinPlayer = oRS.Fields("EloMinPlayer").Value
Else
	oRs.Close
	oConn.Close 
	Set oRS = Nothing
	Set oConn = Nothing
	Set oRs2 = Nothing
	response.clear
	response.redirect "/errorpage.asp?error=7"
end if
ors.Close

strSQL="select lnkEloTeamID, Rating, LastLogin FROM lnk_elo_team  where Eloladderid=" & intLadderID & " and teamid=" & intTeamID
ors.Open strSQL, oconn
if not (ors.eof and ors.BOF) then
	intEloTeamID = ors.Fields(0).Value
	dtmLoginTime = ors.Fields("LastLogin").Value
	intRating = oRS.Fields("Rating").Value
else
	oRs.Close
	oConn.Close 
	Set oRS = Nothing
	Set oConn = Nothing
	Set oRs2 = Nothing
	response.clear
	response.redirect "/errorpage.asp?error=7"
end if
oRs.NextRecordset 

Call ContentStart("Team Administration - " &  Server.HTMLEncode(strTeamName) & " on the " & Server.HTMLEncode(strLadderName) & " Ladder")

if not(bSysAdmin or bTeamCaptain or bTeamFounder or bLadderAdmin)  then
	oConn.Close 
	Set oRS = Nothing
	Set oConn = Nothing
	Set oRs2 = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
else
	if (bTeamCaptain or bTeamFounder) then
		strsql = "update lnk_elo_team set LastLogin = GetDate() where lnkEloTeamID= " & intEloTeamID
		oConn.Execute(strSQL)
	End If
	If (bsysadmin or bLadderAdmin) then
		response.write "<center><B>Last Login Time: " & dtmLoginTime & "</b></center>"
	end if

	Response.Write "<center><p class=text><b>Current Rating:</b> " & intRating & "</center><br /><br />"
	
	' Initiate a challenge
	If Not (blnLocked) Then
		
		%>
		<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
		<form name="frmScrimChallenge" id="frmScrimChallenge" action="/scrim/saveitem.asp" method="post">
		<input type="hidden" name="SaveType" id="SaveType" value="ChallengeTeam" />
		<input type="hidden" name="Team" id="Team" value="<%=Server.HTMLEncode(strTeamName)%>" />
		<input type="hidden" name="Ladder" id="Ladder" value="<%=Server.HTMLEncode(strLadderName)%>" />
		<input type="hidden" name="LinkID" id="LinkID" value="<%=Server.HTMLEncode(intEloTeamID)%>" />
		<input type="hidden" name="LadderID" id="LadderID" value="<%=Server.HTMLEncode(intLadderID)%>" />
		<tr>
			<td>
				<table border="0" cellspacing="1" cellpadding="4">
				<tr>
					<th colspan="2" bgcolor="#000000">Initiate a Challenge</th>
				</tr>
				<tr>
					<td bgcolor="<%=bgcone%>">Choose a team to challenge:</td>
					<td bgcolor="<%=bgcone%>">
						<select name="selLinkID" id="selLinkID">
							<%
							strSQL = " SELECT TeamName, t.TeamID, lnkEloTeamID, Rating "
							strSQL = strSQL & " FROM tbl_teams t "
							strSQL = strSQL & " INNER JOIN lnk_elo_team et ON et.TeamID = t.TeamID "
							strSQL = strSQL & " WHERE et.Active = 1 AND et.EloLadderID = '" & intLadderID & "' AND ABS(Rating - " & intRating & ") <= " & intMaxRatingDiff & " AND et.TeamID <> '" & intTeamID & "'"
							strSQL = strSQL & " AND lnkEloTeamID NOT IN ( SELECT DefenderEloTeamID FROM tbl_elo_matches WHERE AttackerEloTeamID = '" & intEloTeamID & "'  AND MatchActive = 1"
							strSQL = strSQL & " UNION SELECT AttackerEloTeamID FROM tbl_elo_matches WHERE DefenderEloTeamID = '" & intEloTeamID & "' AND MatchActive = 1) " 
							strSQL = strSQL & " ORDER BY TeamName ASC, Rating DESC"
							oRs.Open strSQL, oConn
							If Not(oRs.EOF AND oRs.BOF) Then
								Do While Not(oRs.EOF)
									Response.Write "<option value=""" & oRs.Fields("lnkEloTeamID") & """>" & Server.HTMLEncode(oRs.Fields("TeamName").Value & " -- " & oRs.Fields("Rating").Value) & "</option>" & vbCrLf
									oRs.MoveNext
								Loop
							Else
								Response.Write "<option value="""">No teams available to challenge</option>"
							End If
							oRs.NextRecordSet
							%>
						</select>
					</td>
				</tr>
				<tr>
					<td colspan="2" bgcolor="#000000" align="center"><input type="submit" value="Challenge Selected Team" /></td>
				</tr>
				</table>
			</td>
		</tr>
		</form>
		</table>
		<br />

		<table border="0" cellspacing="0" cellpadding="0" width="97%" bgcolor="#444444">
		<tr>
			<td>
				<table border="0" cellspacing="1" cellpadding="4" width="100%">
				<tr>
					<th colspan="9" bgcolor="#000000">Current Matches</th>
				</tr>
				<tr>
					<th bgcolor="#000000">Opponent (Rating)</th>
					<th bgcolor="#000000">Challenge Date</th>
					<th bgcolor="#000000">Match Date</th>
					<th bgcolor="#000000">Maps</th>
					<th bgcolor="#000000">Last Comm</th>
					<th bgcolor="#000000">Manage</th>
				</tr>
				<%	
				strSQL = "SELECT EloMatchID, DefenderEloTeamID, AttackerEloTeamID, ChallengeDate, MatchDate, EloLadderID, Map1, Map2, Map3, Map4, Map5, LastComm = (SELECT TOP 1 CommDate FROM tbl_elo_comms ec WHERE ec.EloMatchID = m.EloMatchID ORDER BY EloCommID DESC) "
				strSQL = strSQL & " FROM tbl_elo_matches m WHERE (DefenderEloTeamID = '" & intEloTeamID & "' OR AttackerEloTeamID = '" & intEloTeamID & "') AND MatchActive = 1 ORDER BY ChallengeDate ASC"
				oRs.Open strSQL, oConn
				If Not(oRs.EOF AND oRs.BOF) Then
					If intMatchID = 0 Then
						intMatchID = oRs.Fields("EloMatchID").Value
					End If
					Do While Not(oRs.EOF)
						If bgc = bgcone Then
							bgc = bgctwo
						Else
							bgc = bgcone
						End If
						
						If oRs.Fields("DefenderEloTeamID").Value = intEloTeamID Then
							strSQL = "SELECT TeamName, TeamTag, Rating FROM tbl_teams t INNER JOIN lnk_elo_team et ON et.TeamID = t.TeamID WHERE et.lnkEloTeamID = '" & oRs.Fields("AttackerEloTeamID").Value & "'"
						Else
							strSQL = "SELECT TeamName, TeamTag, Rating FROM tbl_teams t INNER JOIN lnk_elo_team et ON et.TeamID = t.TeamID WHERE et.lnkEloTeamID = '" & oRs.Fields("DefenderEloTeamID").Value & "'"
						End If
						oRs2.Open strSQL, oConn
						If Not(oRs2.EOF AND oRs2.BOF) Then
							
							strOpponentName =  oRs2.Fields("TeamName").Value
							strOpponentTag = oRs2.Fields("TeamTag").Value
							intOpponentRanking = oRs2.Fields("Rating").Value
						End If
						oRs2.NextRecordSet						
						%>
						<tr>
							<% If intMatchID = oRs.Fields("EloMatchID").Value Then %>
							<td bgcolor="<%=bgc%>" rowspan="2" valign="top"><a href="/viewteam.asp?team=<%=Server.URLEncode(strOpponentName & "")%>"><%=Server.HTMLEncode(strOpponentName & "")%> (<%=intOpponentRanking%>)</a></td>
							<% Else %>
							<td bgcolor="<%=bgc%>"><a href="/viewteam.asp?team=<%=Server.URLEncode(strOpponentName & "")%>"><%=Server.HTMLEncode(strOpponentName & "")%> (<%=intOpponentRanking%>)</a></td>
							<% End If %>
							<td bgcolor="<%=bgc%>" align="center"><%=FormatDateTime(oRs.Fields("ChallengeDate").Value, 2)%></td>
							<td bgcolor="<%=bgc%>"><%
								If IsDate(oRs.FieldS("MatchDate").Value) Then
									Response.Write FormatDateTime(oRs.Fields("MatchDate").Value, 0)
								Else
									Response.Write "Unscheduled"
								End If
								%></td>
							<td bgcolor="<%=bgc%>"><%
								If Not(IsNull(oRs.Fields("Map1").Value) OR Len(oRs.Fields("Map1").Value) = 0) Then
									Response.Write oRs.Fields("Map1").Value
								End If
								If Not(IsNull(oRs.Fields("Map2").Value) OR Len(oRs.Fields("Map2").Value) = 0) Then
									Response.Write ", " & oRs.Fields("Map2").Value
								End If
								If Not(IsNull(oRs.Fields("Map3").Value) OR Len(oRs.Fields("Map3").Value) = 0) Then
									Response.Write ", " & oRs.Fields("Map3").Value
								End If
								If Not(IsNull(oRs.Fields("Map4").Value) OR Len(oRs.Fields("Map4").Value) = 0) Then
									Response.Write ", " & oRs.Fields("Map4").Value
								End If
								If Not(IsNull(oRs.Fields("Map5").Value) OR Len(oRs.Fields("Map5").Value) = 0) Then
									Response.Write ", " & oRs.Fields("Map5").Value
								End If
								%></td>
							<td bgcolor="<%=bgc%>" align="center"><%
								If IsDate(oRs.FieldS("LastComm").Value) Then
									Response.Write FormatDateTime(oRs.Fields("LastComm").Value, 0)
								Else
									Response.Write " never "
								End If
								%></td>
							<% If intMatchID <> oRs.Fields("EloMatchID").Value Then %>
							<td bgcolor="<%=bgc%>" align="center"><a href="teamscrimladderadmin.asp?ladder=<%=Server.URLEncode(strLadderName)%>&team=<%=Server.URLEncode(strTeamName)%>&matchid=<%=oRs.Fields("EloMatchID").Value%>">manage</a></td>
							<% Else %>
							<td bgcolor="<%=bgc%>" align="center"><a href="teamscrimladderadmin.asp?ladder=<%=Server.URLEncode(strLadderName)%>&team=<%=Server.URLEncode(strTeamName)%>&matchid=-1">hide</a></td>
							<% End If %>
						</tr>
						<%
						If intMatchID = oRs.Fields("EloMatchID").Value Then
							%>
							<tr>
								<td colspan="7" bgcolor="#000000" align="center">
									<br />
									<b><a href="ScrimMatchReportLoss.asp?matchid=<%=intMatchID%>&linkid=<%=intEloTeamID%>">Report Loss</a></b><br />
									<br />
									<b><a href="DisputeMatchScrimLadder.asp?MatchID=<%=intMatchID%>&DisputeTeamID=<%=intEloTeamID%>&Ladder=<%=Server.URLEncode(strLadderName)%>">Dispute Match</a></b><br />
									<br />
									<% If oRs.Fields("DefenderEloTeamID").Value = intEloTeamID Then %>
									<b><a href="/scrim/saveitem.asp?savetype=KillMatch&matchid=<%=intMatchID%>&linkid=<%=intEloTeamID%>&team=<%=Server.URLEncode(strTeamName)%>&ladder=<%=Server.URLEncode(strLadderName)%>">Decline Challenge</a></b><br />
									<% Else %>
									<b><a href="/scrim/saveitem.asp?savetype=KillMatch&matchid=<%=intMatchID%>&linkid=<%=intEloTeamID%>&team=<%=Server.URLEncode(strTeamName)%>&ladder=<%=Server.URLEncode(strLadderName)%>">Retract Challenge</a></b><br />
									<% End If %>
									
									<br />
									<table border="0" cellspacing="0" cellpadding="0" width="97%" bgcolor="#444444" align="center">
									<tr>
										<td>
											<table border="0" cellspacing="1" cellpadding="4" width="100%">
											<tr>
												<th bgcolor="#000000">Match Communications</th>
											</tr>
											<%
											If bTeamCaptain or bTeamFounder Then
												%>
												<tr>
													<td bgcolor="#000000" align="center"><b><a href="scrimmatchcomms.asp?MatchID=<%=intMatchID%>&mode=add&tag=<%=server.urlencode(strTeamTag)%>&ladder=<%=server.urlencode(strLadderName)%>&team=<%=server.urlencode(strTeamName)%>">add match communication</a></td>
												</tr>
												<%
											else
												%>
												<tr>
													<td align="center" bgcolor="#000000"><b><a href="scrimmatchcomms.asp?MatchID=<%=intMatchID%>&mode=add&tag=TWLAdmin&ladder=<%=server.urlencode(strLadderName)%>&team=<%=server.urlencode(strTeamName)%>">add match communication</a></td>
												</tr>
												<%
											end if
											strSQL = "select * from tbl_elo_Comms where ((ElomatchID='" & intMatchID & "') and (CommDead=0)) order by EloCommID DESC"
											oRs2.Open strSQL, oconn
											bgc2=bgcone
											if not (ors2.EOF and ors2.bof) then
												do while not ors2.EOF
													%>
													<tr>
														<td bgcolor="<%=bgc2%>">
															Author: <b><%=oRs2.Fields("CommAuthor").Value%> - Posted: <%=FormatDateTime(oRs2.Fields("CommDate").Value)%></b><br />
															<%
															If bSysAdmin Then
																%>
																<a href="scrimmatchcomms.asp?MatchID=<%=intMatchID%>&mode=edit&tag=TWLAdmin&ladder=<%=server.urlencode(strLadderName)%>&commid=<%=oRs2.Fields("EloCommID").Value%>&team=<%=server.urlencode(strTeamName)%>">edit</a> - 
																<a href="/scrim/SaveItem.asp?commid=<%=ors2.Fields("EloCommID").Value%>&SaveType=Delete_Communications&ladder=<%=server.urlencode(strLadderName)%>&team=<%=server.urlencode(strTeamName)%>">Delete</a><br />
																<%
															End If
															Response.Write oRs2.Fields("Comms").Value
															%>
														</td>
													</tr>	
													<%
													if bgc2 = bgcone then
														bgc2=bgctwo
													else
														bgc2=bgcone
													end if
													oRs2.MoveNext
												Loop
											Else
												%>
												<tr>
													<td bgcolor="<%=bgc2%>">No Match Communications Entered</td>
												</tr>
												<%
											End If
											oRs2.NextRecordSet
											%>
											</table>
										</td>
									</tr>
									</table>
									<br />
									
									<script language="javascript" type="text/javascript">
									<!--
										function fCheckValidDate(oForm) {
											var intMonth = oForm.selMonth.options[oForm.selMonth.selectedIndex].value;
											var intDay = oForm.selDay.options[oForm.selDay.selectedIndex].value;
											var intYear = oForm.selYear.options[oForm.selYear.selectedIndex].value;
											var blnValidDate = false;
											
											if (intMonth == 1 || intMonth == 3 || intMonth == 5 || intMonth == 7 || intMonth == 8 || intMonth == 10 || intMonth == 12) {
												blnValidDate = true;											
											} else if (intMonth == 2) {
												if (intDay < 29) {
													blnValidDate = true;
												} else if ((intDay == 29) && (intYear % 4 == 0)) {
													blnValidDate = true;
												}
											} else {
												if (intDay < 31) {
													blnValidDate = true;
												}
											}
											return blnValidDate;
										}
									//-->
									</script>
									<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center">
									<form name="frmDateChange" id="frmDateChange" method="post" action="/scrim/saveitem.asp" onSubmit="return fCheckValidDate(this.form)">
									<input type="hidden" name="SaveType" id="SaveType" value="ChangeMatchTime" />
									<input type="hidden" name="MatchID" id="MatchID" value="<%=Server.HTMLEncode(intMatchID)%>" />
									<input type="hidden" name="Team" id="Team" value="<%=Server.HTMLEncode(strTeamName)%>" />
									<input type="hidden" name="Ladder" id="Ladder" value="<%=Server.HTMLEncode(strLadderName)%>" />
									<tr>
										<td>
											<table border="0" cellspacing="1" cellpadding="4" width="100%">
											<tr>
												<th bgcolor="#000000" colspan="2">Configure Match Time</th>
											</tr>
											<tr>
												<td bgcolor="<%=bgcone%>">Date:</td>
												<td bgcolor="<%=bgcone%>">
													<select name="selMonth" id="selMonth">
														<%
														Dim i
														For i = 1 to 12
															Response.Write "<option value=""" & i & """"
															If i = Month(Now()) Then
																Response.Write " selected=""selected"""
															End If
															Response.Write ">" & MonthName(i) & "</option>" & vbCrLf
															
														Next
														%>
													</select>

													<select name="selDay" id="selDay">
														<%
														For i = 1 to 31
															Response.Write "<option value=""" & i & """"
															If i = Day(Now()) Then
																Response.Write " selected=""selected"""
															End If
															Response.Write ">" & i & "</option>" & vbCrLf
														Next
														%>
													</select>

													<select name="selYear" id="selYear">
														<%
														For i = 0 to 1
															Response.Write "<option value=""" & Year(now()) + i & """>" & Year(now()) + i & "</option>" & vbCrLf
														Next
														%>
													</select>
												</td>
											</tr>
											<tr>
												<td bgcolor="<%=bgctwo%>">Time:</td>
												<td bgcolor="<%=bgctwo%>">
													<select name="selHour" id="selHour">
														<%
														For i = 1 to 12
															Response.Write "<option value=""" & i & """>" & i & "</option>" & vbCrLf
														Next
														%>
													</select>
													:
													<select name="selMinute" id="selMinute">
														<%
														For i = 0 to 45 Step 15
															If i = 0 Then
																Response.Write "<option value=""0" & i & """>0" & i & "</option>" & vbCrLf
															Else
																Response.Write "<option value=""" & i & """>" & i & "</option>" & vbCrLf
															End If
														Next
														%>
													</select>
													
													<select name="selAMPM" id="selAMPM">
														<option value="PM">PM</option>
														<option value="AM">AM</option>
													</select>
													EST
												</td>
											</tr>
											<tr>
												<td colspan="2" bgcolor="#000000%" align="center">
													<input type="submit" value="Change Match Time" />
												</td>
											</tr>
											</table>
										</tr>
									</tr>
									</form>
									</table>
									
									<br />									
									<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center">
									<form name="frmDateChange" id="frmDateChange" method="post" action="/scrim/saveitem.asp">
									<input type="hidden" name="SaveType" id="SaveType" value="ChangeMaps" />
									<input type="hidden" name="MatchID" id="MatchID" value="<%=Server.HTMLEncode(intMatchID)%>" />
									<input type="hidden" name="Team" id="Team" value="<%=Server.HTMLEncode(strTeamName)%>" />
									<input type="hidden" name="Ladder" id="Ladder" value="<%=Server.HTMLEncode(strLadderName)%>" />
									<tr>
										<td>
											<table border="0" cellspacing="1" cellpadding="4" width="100%">
											<tr>
												<th bgcolor="#000000" colspan="2">Configure Maps</th>
											</tr>
											<%
											strSQL = "SELECT MapName FROM tbl_maps m INNER JOIN lnk_elo_maps em ON em.MapID = m.MapID WHERE em.EloLadderID = '" & intLadderID & "'"
											oRs2.Open strSQL, oConn, 3, 3
											If Not(oRs2.EOF AND oRs2.BOF) Then
												%>
												<tr>
													<td bgcolor="<%=bgcone%>">Map 1:</td>
													<td bgcolor="<%=bgctwo%>"><select name="selMap1" id="selMap1">
														<option value="">No Map 1</option>
														<%
														oRs2.MoveFirst
														Do While Not(ors2.EOF)
															Response.Write "<option value=""" & Server.HTMLEncode(oRs2.Fields("MapName").Value) & """"
															If oRs2.Fields("MapName").Value = oRs.Fields("Map1").Value Then 
																Response.Write " selected=""selected"""
															End If																
															Response.Write ">" & Server.HTMLEncode(oRs2.Fields("MapName").Value) & "</option>"
															oRs2.MoveNext
														Loop
														%>
														</select>
													</td>
												</tr>
												<tr>
													<td bgcolor="<%=bgctwo%>">Map 2:</td>
													<td bgcolor="<%=bgctwo%>"><select name="selMap2" id="selMap2">
														<option value="">No Map 2</option>
														<%
														oRs2.MoveFirst
														Do While Not(ors2.EOF)
															Response.Write "<option value=""" & Server.HTMLEncode(oRs2.Fields("MapName").Value) & """"
															If oRs2.Fields("MapName").Value = oRs.Fields("Map2").Value Then 
																Response.Write " selected=""selected"""
															End If																
															Response.Write ">" & Server.HTMLEncode(oRs2.Fields("MapName").Value) & "</option>"
															oRs2.MoveNext
														Loop
														%>
														</select>
													</td>
												</tr>
												<tr>
													<td bgcolor="<%=bgcone%>">Map 3:</td>
													<td bgcolor="<%=bgctwo%>"><select name="selMap3" id="selMap3">
														<option value="">No Map 3</option>
														<%
														oRs2.MoveFirst
														Do While Not(ors2.EOF)
															Response.Write "<option value=""" & Server.HTMLEncode(oRs2.Fields("MapName").Value) & """"
															If oRs2.Fields("MapName").Value = oRs.Fields("Map3").Value Then 
																Response.Write " selected=""selected"""
															End If																
															Response.Write ">" & Server.HTMLEncode(oRs2.Fields("MapName").Value) & "</option>"
															oRs2.MoveNext
														Loop
														%>
														</select>
													</td>
												</tr>
												<tr>
													<td bgcolor="<%=bgcone%>">Map 4:</td>
													<td bgcolor="<%=bgctwo%>"><select name="selMap4" id="selMap4">
														<option value="">No Map 4</option>
														<%
														oRs2.MoveFirst
														Do While Not(ors2.EOF)
															Response.Write "<option value=""" & Server.HTMLEncode(oRs2.Fields("MapName").Value) & """"
															If oRs2.Fields("MapName").Value = oRs.Fields("Map4").Value Then 
																Response.Write " selected=""selected"""
															End If																
															Response.Write ">" & Server.HTMLEncode(oRs2.Fields("MapName").Value) & "</option>"
															oRs2.MoveNext
														Loop
														%>
														</select>
													</td>
												</tr>
												<tr>
													<td bgcolor="<%=bgcone%>">Map 5:</td>
													<td bgcolor="<%=bgctwo%>"><select name="selMap5" id="selMap5">
														<option value="">No Map 5</option>
														<%
														oRs2.MoveFirst
														Do While Not(ors2.EOF)
															Response.Write "<option value=""" & Server.HTMLEncode(oRs2.Fields("MapName").Value) & """"
															If oRs2.Fields("MapName").Value = oRs.Fields("Map5").Value Then 
																Response.Write " selected=""selected"""
															End If																
															Response.Write ">" & Server.HTMLEncode(oRs2.Fields("MapName").Value) & "</option>"
															oRs2.MoveNext
														Loop
														%>
														</select>
													</td>
												</tr>
												<tr>
													<td align="center" colspan="2" bgcolor="#000000"><input type="submit" value="Save Maps" /></td>
												</tr>
											<%
										End If
										oRs2.NextRecordSet
										%>
											</table>
										</td>
									</tr>
									</form>
									</table>
									<br />
									</td>
								</tr>
							<%
						End If
						
						oRs.MoveNext
					Loop
				Else
					%>
					<tr>
						<td colspan="8" bgcolor="#000000">No matches currently</td>
					</tr>
					<%
				End If
				oRs.NextRecordSet
				%>
				</table>
			</td>
		</tr>
		</table>
		<%
	Else
		%>
		Ladder is locked at this time.
		<%
	End If
	

Call ContentEnd()
Call Content3BoxStart("Ladder Captain Management")

strSQL = "select tbl_players.playerhandle, l.PlayerID  from tbl_players inner join lnk_elo_team_player l on l.PlayerID=tbl_Players.playerid where (l.lnkEloTeamID='" & intEloTeamID & "' and L.isadmin=1) order by tbl_players.playerhandle"
ors.Open strSQL, oconn

Dim strInfo
Response.Write "<table border=0 align=center width=97% cellspacing=0><tr height=30 bgcolor="&bgcone&"><td align=center><b>Current Captains</b></td></tr>"
if not (ors.EOF and ors.BOF) then
	bgc=bgcone
	do while not ors.EOF
		strInfo=""
		if ors.fields("playerID").value = intFounderID then 
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

strSQL = "select tbl_players.playerhandle, l.lnkEloTeamPlayerID, l.PlayerID  from tbl_players inner join lnk_elo_team_player l on l.PlayerID=tbl_Players.playerid where (l.lnkEloTeamID='" & intEloTeamID & "' and L.isadmin=0) order by tbl_players.playerhandle"
ors.Open strSQL, oconn
if not(ors.EOF and ors.BOF) then
	%>
	<form name=promote action=/scrim/saveitem.asp method=post><table align=center border=0 width=97% cellspacing=0><tr bgcolor=<%=bgcone%> height=30><td align=center><b>Promote Player to Captain</b></td></tr>
	<tr bgcolor=<%=bgctwo%>><td height=30 align=center><select name=playerlist style='width:150'>
	<%
	do while not ors.EOF
		Response.Write "<option value=""" & ors.Fields(1).Value & """>" & Server.HTMLEncode(ors.Fields(0).Value)
		ors.MoveNext
	loop
	%>
	</select></td></tr><tr height=30 bgcolor=<%=bgcone%>><td align=center>
	<input class=bright type=submit value='Promote'>
	<input type=hidden name=SaveType value=PromoteScrimCaptain>
	<input type=hidden name=ladder value="<%=Server.HTMLEncode(strLadderName)%>">
	<input type=hidden name=team value="<%=Server.HTMLEncode(strTeamName)%>">
	</td></tr></table></form>
	<%
end if
ors.Close

Call Content3BoxMiddle2()

strSQL = "select tbl_players.playerhandle, l.lnkEloTeamPlayerID, l.PlayerID  from tbl_players inner join lnk_elo_team_player l on l.PlayerID=tbl_Players.playerid where (l.lnkEloTeamID='" & intEloTeamID & "' and L.isadmin=1) order by tbl_players.playerhandle"
ors.Open strSQL, oconn
if not(ors.EOF and ors.BOF) then
	%>
	<form name=demote action=/scrim/saveitem.asp method=post><table align=center border=0 width=97% cellspacing=0>
	<tr bgcolor=<%=bgcone%> height=30><td align=center><b>Demote Captain</b></td></tr>
	<tr bgcolor=<%=bgctwo%> height=30><td align=center><select name=playerlist style='width:150'>
	<%
	do while not ors.EOF
		if ors.fields("PlayerID").value <> intFounderID then
			Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value)
		end if
		ors.MoveNext
	loop
	%>
	</select></td></tr><tr bgcolor=<%=bgcone%> height=30><td align=center><input class=bright type=submit id=submit2 name=submit2 value=Demote></td>
	</tr>
		<input type=hidden name=SaveType value=DemoteScrimCaptain>
		<input type=hidden name=ladder value="<%=server.HTMLEncode(strLadderName)%>">
		<input type=hidden name=team value="<%=Server.HTMLEncode(strTeamName)%>"></table>
	</form>
	<%
end if
ors.Close

Call Content3BoxEnd()
Call ContentStart(Server.HTMLEncode(strLadderName) & " Ladder Roster Management")


	strSQL = "select tbl_players.playerhandle, l.PlayerID from tbl_players inner join lnk_elo_team_player l on l.PlayerID=tbl_Players.playerid where (l.lnkEloTeamID='" & intEloTeamID & "') order by tbl_players.playerhandle"
	ors.open strsql,oconn
	if not (ors.eof and ors.bof) then
		%>
		<form name=BootPlayer method=post action=/scrim/saveitem.asp>
		<table width=50% align=center border=0 cellspacing=0 cellpadding=0>
		<tr bgcolor=<%=bgcone%> height=125 valign=center><td align=center>
		<input type=hidden name=savetype value=DropPlayer>
			<select name=PlayerID size=5 class=brightred style='width:200'>
		<%
		do while not ors.eof
			if ors.fields("PlayerID").value <> intFounderID then 
				response.write "<option value=" & ors.fields("PlayerID").value & ">" & Server.HTMLEncode(ors.fields(0).value)
			end if
			ors.movenext
		loop
		response.write "</select></td></tr><tr bgcolor="&bgctwo&" height=35><td align=center><input type=hidden name=link value=" & intEloTeamID & "><input type=submit class=bright style='width:75' value='Kick Player'></td></tr></table></form>"
	end if
	ors.close
end if

Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>