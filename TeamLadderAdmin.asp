<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Team Ladder Administration"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim TeamName, LadderName, m3settings
teamname = Request.QueryString("team")
laddername = Request.QueryString("ladder")

bTeamFounder = IsTeamFounder(TeamName)
bTeamCaptain = IsTeamCaptain(TeamName, LadderName)
bLadderAdmin = IsLadderAdmin(LadderName)

Dim TeamID, TeamTag, PlayerID, ownerID, ownerName, tima
Dim LadderID, LadderLocked, LadderChallenge, LinkID, LoginTime
Dim tRank, mStatus, players, minplayer, MaxJump, lid, rdays, intMaps
Dim enemyID, mDate, xDate, aDate, enemyname, map1, map2, map3, tid, dayprint
Dim matchdate1, matchdate2, matchid, m1side, m2Side, m3Side, info, grammer
Dim CurrentMap, intOptionID, blnOptionSame, intCounter, blnOptionShown
Dim MapConfiguration, Maps, i, map4, map5, strVerbiage, intChallengeDays, rRank
Dim MapArray(6)
PlayerID = Session("PlayerID")
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->
<% IF bSysAdmin Then %>
<script language="javascript" type"text/javascript">
<!-- 
function PopChange() {
	var objChangeRecord = window.open("TeamLadderRecordChange.asp?team=<%=Server.URLEncode(teamname)%>&ladder=<%=Server.URLEncode(LadderName)%>", "RecordChange",  "width=300,height=200,toolbar=0,scrollbars=0,status=0,location=0,menubar=0,resizable=0");
	objChangeRecord.focus();
}
//-->
</script>
<% ENd IF %>
<%
Call ContentStart("Team Administration - " &  Server.HTMLEncode(Request.QueryString("team")) & " on the " & Server.HTMLEncode(Request.QueryString("ladder")) & " Ladder")

strSQL = "SELECT teamid, teamtag, teamfounderid from tbl_teams where teamname='" & CheckString(teamname) & "'"
ors.Open strSQL, oconn
if not (ors.eof and ors.BOF) then
	teamid = ors.Fields(0).Value
	teamtag = ors.fields(1).value
	ownerID = oRS.Fields("TeamFounderID").Value
	
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

strSQL = "select ladderid, LadderLocked, LadderChallenge, MinPlayer, Maps, MapConfiguration, ChallengeDays from tbl_ladders where laddername='" & replace(laddername, "'", "''") & "'"
ors.Open strSQL, oconn
if not (ors.eof and ors.BOF) then
	ladderid=ors.Fields(0).Value
	ladderlocked = ors.fields(1).value
	LadderChallenge = ors.fields(2).value
	MinPlayer = oRS.Fields("MinPlayer").Value
	Maps = oRS.Fields("Maps").Value 
	intMaps = Maps
	MapConfiguration = oRS.Fields("MapCOnfiguration").Value 
	intChallengeDays = oRS.Fields("ChallengeDays").Value
else
	oRs.Close
	oConn.Close 
	Set oRS = Nothing
	Set oConn = Nothing
	Set oRs2 = Nothing
	response.clear
	response.redirect "/errorpage.asp?error=7"
end if
ors.Close

strSQL="select TLLinkID, Rank, LastLogin, Status from lnk_T_L where ladderid=" & ladderid & " and teamid=" & teamid
ors.Open strSQL, oconn
if not (ors.eof and ors.BOF) then
	linkid=ors.Fields(0).Value
	logintime=ors.Fields("LastLogin").Value
	trank = oRS.Fields("Rank").Value
	mStatus = oRS.Fields("Status").Value
	
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

	
	
	
IF bSysAdmin Then 
	%>
	<a href="saveitem.asp?SaveType=ClearStatus&TLLinkID=<%=linkid%>&team=<%=Server.URLEncode(teamname)%>&ladder=<%=Server.URLEncode(LadderName)%>">clear b0rked match (use carefully, ask Triston if you have q's)</a><br />
	<a href="saveitem.asp?SaveType=ResetRest&TLLinkID=<%=linkid%>&team=<%=Server.URLEncode(teamname)%>&ladder=<%=Server.URLEncode(LadderName)%>">Reset Rest Days</a><br />
	<%
End if

strsql="select playerhandle from tbl_players where playerid="&ownerid
oRS.open strsql,oconn
if not (oRS.eof and oRS.bof) then
	ownername = oRS.fields(0).value
else
	ownername="No Owner"
end if
oRS.NextRecordset 

If Right(uCase(teamname), 1) = "s" Then
	strVerbiage = " have "
Else
	strVerbiage = " has "
End IF
if not(bSysAdmin or bTeamCaptain or bTeamFounder or bLadderAdmin)  then
	oConn.Close 
	Set oRS = Nothing
	Set oConn = Nothing
	Set oRs2 = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
else
	if (bTeamCaptain or bTeamFounder) then
		strsql = "update lnk_t_L set LastLogin = '" & now() & "' where TLLinkID= " & linkid 
		oConn.Execute(strSQL)
	End If
	If (bsysadmin or bLadderAdmin) then
		response.write "<center><B>Last Login Time: " & logintime & "</b></center>"
		if (bsysadmin) Then
			Response.write "<br /><center><a href=""javascript:PopChange();"">Change this team's record</a></center><br />"
		End If
	end if

	Response.Write "<center><b>Current Rung:</b> " & tRank & "</center>"
	
	if mstatus="Attacking" then
		strSQL = "select * FROM vMatches where MatchAttackerID = " & linkid & " and matchladderid=" & ladderid
		ors.Open strSQL, oconn
		if not (ors.EOF and ors.BOF) then
			enemyID = ors.Fields("MatchDefenderID").Value
			mDate = ors.Fields("MatchDate").Value 
			xDate = ors.Fields("MatchChallengeDate").Value 
			aDate = ors.Fields("MatchAcceptanceDate").Value 
			enemyname = ors.Fields("DefenderName").Value
			map1=ors.Fields("MatchMap1ID").value
			map2=ors.Fields("MatchMap2ID").value
			map3=ors.Fields("MatchMap3ID").value
			map4=ors.Fields("MatchMap4ID").value
			map5=ors.Fields("MatchMap5ID").value
			MapArray(1) = ors.Fields("MatchMap1ID").Value 
			MapArray(2) = ors.Fields("MatchMap2ID").Value 
			MapArray(3) = ors.Fields("MatchMap3ID").Value 
			MapArray(4) = ors.Fields("MatchMap4ID").Value 
			MapArray(5) = ors.Fields("MatchMap5ID").Value 
			matchdate1 = ors.Fields("MatchSelDate1").Value
			matchdate2= ors.Fields("MatchSelDate2").Value 
			matchid = ors.Fields("MatchID").Value 
		end if
		Response.Write "<center><b>Current Status:</b> Attacking <a href=viewteam.asp?team=" & server.urlencode(enemyname) & ">"&Server.HTMLEncode(enemyname)&"</a> <br>Maps: "
		For i = 1 to Maps
			If i > 1 Then
				Response.Write ", "
			End If
			Response.Write Server.HTMLEncode(MapArray(i))
		Next
		response.write "<BR>Match Date: " & mdate & "</center>"
		if (bsysadmin or bLadderAdmin) and mdate <> "TBD" then %>
			<br><center><input value="Clear Match Date" type=button class=brightgold style="width:200" onclick="window.location.href='ladder/ladderengine.asp?clearDate=true&matchid=<%=matchid%>';" id=clearDate name=clearDate />
			<form name=changeDate action=ladder/ladderengine.asp method=post>
				<input type="hidden" name="saveType" value="changeDate">
				<input type="hidden" name="matchid" value="<%=matchid%>">
				<input type="hidden" name="timezone" value="<%=RIght(mdate, 3)%>">
				Month: <input type="Text" name="newMonth" value="" size="2" maxlength="2">&nbsp;Day:<input type="Text" name="newDay" value="" size="2" maxlength="2">&nbsp;
				Year: <input type="text" name="newYear" id="newYear" value="<%=year(now)%>" maxlength="4" size="4"><br />
				New Time: <input type="Text" name="newHour" value="" size="2" maxlength="2">:<input type="Text" name="newMinute" value="" size="2" maxlength="2">:00 PM<br>
				<input type="submit" class=brightgold style="width:150" value="Change Date" id=submit4 name=submit4>
			</form></center><%
		end if
		response.write "<center><BR><b>Challenge Date:</b> " & xDate & ""
		response.write "<BR><b>Accept Date:</b> " & aDate & "</center>"
		%>
		<!-- #include virtual="/include/incruleinsert.asp" -->
		<%
		if matchdate1="TBD" then
			Response.Write "<center><b>Awaiting match acceptance from " & enemyname & " (Challenged on " & xDate & ")</b></center>"
		elseif mDate = "TBD" then
			if right(ucase(enemyname),1)="S" then
				grammer=" have "
			else
				grammer=" has "
			end if
			%>
				<form name=frmAccept action=saveitem.asp method=post>
				<table width="75%" class="cssBordered" align="center">
				<TR BGCOLOR="#000000"><TH><%="Challenge accepted by " & enemyname & " on " &  aDate%></TH></TR>
				<TR BGCOLOR="<%=bgctwo%>"><TD><%=Server.HTMLEncode(enemyname) & grammer%> selected the match dates listed below. Confirm your choice below</TD></TR>
				<tr BGCOLOR="<%=bgcone%>"><td align=center>Chosen Dates: <select name=matchdate class=brightred><option selected><%=matchdate1%><option><%=matchdate2%></select></td></tr>
				<tr BGCOLOR="<%=bgcone%>"><td align=center>Approved for Shoutcasting: <input type=checkbox name=scApproved value=true checked></td></tr>
				<%
				For i = 1 to Len(MapConfiguration)
					If mid(MapConfiguration, i, 1) = "A" Then
						'' Allow them to choose a map	
						%>
						<TR>
							<TD ALIGN=CENTER BGCOLOR=<%=bgcone%>>Choose Map <%=i%>: <SELECT Name=Map<%=i%> CLASS=bright>
							<%
							strSQL = "EXEC GetMapList '" & matchid & "', " & i
							oRS2.Open strsql, oconn
							if not (oRS2.EOF and oRS2.BOF) then
								do while not oRS2.EOF
									Response.Write "<option VALUE=""" & oRS2.Fields("MapName").Value & """>" & oRS2.Fields("MapName").Value & "</OPTION>" & vbCrLf
									oRS2.MoveNext 
								loop
							End If
							oRS2.Close 
							%>
							</TD>
						</TR>
						<%
					End If
				Next
				%>				
				<tr BGCOLOR="<%=bgcone%>"><td align=center>
					<INPUT TYPE=HIDDEN NAME=MC VALUE="<%=MapConfiguration%>">
					<input type=hidden name=matchid value=<%=matchid%>>
					<input type=hidden name=SaveType value=AcceptMatchDate>
					<input type=hidden name=Ladder value="<%=Server.HTMLEncode(LadderName)%>">
					<input type=hidden name=Team value="<%=Server.HTMLEncode(TeamName)%>">
					<input type=submit name=submit1 value="Confirm Date and Time" class=bright>
				</td></tr>
			</table>
				</form>
			<%
	else
		strSQL = "SELECT * FROM vLadderOptions "
		strSQL = strSQL & " WHERE SelectedBy <> 'R' AND "
		strSQL = strSQL & " LadderID = '" & LadderID & "' AND "
		strSQL = strSQL & " OptionID NOT IN (SELECT mo.OptionID FROM lnk_match_options mo WHERE mo.MatchID = '" & matchid & "')"
		'Response.Write strSQL
		oRs2.Open strSQL, oConn
		If Not(oRS2.EOF AND oRS2.BOF) Then
			blnOptionShown = False
			%>
			<FORM NAME=frm_map_options ACTION="/ladder/option_saveitem.asp" METHOD="POST">
			<INPUT TYPE=HIDDEN NAME="MatchID" VALUE="<%=matchid%>">
			<INPUT TYPE=HIDDEN NAME="SaveType" VALUE="SaveMapOptions">
			<INPUT TYPE=HIDDEN NAME="LadderName" VALUE="<%=Server.HTMLEncode(LadderName & "")%>">
			<INPUT TYPE=HIDDEN NAME="TeamName" VALUE="<%=Server.HTMLEncode(TeamName & "")%>">
				
				<table width="75%" class="cssBordered" align="center">
			<TR BGCOLOR="#000000">
				<TH COLSPAN=2>Select Match Options</TH>
			</TR>
			<%
			intCounter = 0
			Do While Not(oRS2.EOF) And intCounter < 100
				intCounter = intCounter + 1
				intOptionID = oRS2.Fields("OptionID").Value 
				If bgc = bgcone Then
					bgc = bgctwo
				Else
					bgc = bgcone
				End If
				Response.Write "<TR BGCOLOR=""" & bgc & """>"
				If oRs2.Fields("MapNumber").Value <> 0 Then
					CurrentMap = MapArray(cint(oRs2.Fields("MapNumber").Value))
				Else 
					CurrentMap = "the match"
				End If
				Select Case(oRS2.Fields("SelectedBy").Value)
					Case "A", "C"
						blnOptionSame = True
						Response.Write "<TD ALIGN=RIGHT>Choose your " & lCase(oRS2.Fields("OptionName").Value) & " for " & CurrentMap & ": </TD>"
						Response.Write "<TD><SELECT NAME=MO_" & oRS2.Fields("OptionID").Value & ">"
						While Not(oRS2.EOF) AND blnOptionSame
							Response.Write "<OPTION VALUE=""" & oRS2.Fields("OptionValueID").Value & """>" & oRS2.Fields("ValueName").Value & "</OPTION>" & vbCrLf
							oRS2.MoveNext
							If Not(oRS2.EOF) Then
								If oRS2.Fields("OptionID").Value = intOptionID Then
									blnOptionSame = True
								Else
									blnOptionSame = False
								End If
							End If
						Wend
						Response.Write "</SELECT></TD>"
						blnOptionShown = True
					Case "D"
						blnOptionSame = True
						Response.Write "<TD COLSPAN=2>" & EnemyName & " will choose " & lcase(oRs2.Fields("OptionName").Value) & " for " & CurrentMap & "</TD>"
						While Not(oRS2.EOF) AND blnOptionSame
							oRS2.MoveNext
							If Not(oRS2.EOF) Then
								If oRS2.Fields("OptionID").Value = intOptionID Then
									blnOptionSame = True
								Else
									blnOptionSame = False
								End If
							End If
						Wend
					Case Else
						Response.Write "Error"
						oRS2.MoveNext 						
				End Select
				Response.Write "</TR>"
'					oRS2.MoveNext
			Loop
			If blnOptionShown Then
				Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=2 ALIGN=center><INPUT TYPE=SUBMIT VALUE=""Confirm Match Options""></TD></TR>"
			End If
			%>
			</table>
			</FORM>
			<%
		End If
		oRS2.NextRecordset 
		' End Map Selection Options

		' Show Current Selected Options
		strSQL = "SELECT * FROM vMatchOptions "
		strSQL = strSQL & " WHERE MatchID = '" & matchid & "'" 
		'Response.Write strSQL
		oRs2.Open strSQL, oConn
		If Not(oRS2.EOF AND oRS2.BOF) Then
			blnOptionShown = False
			%>
			<table width="75%" class="cssBordered" align="center">
			<TR BGCOLOR="#000000">
				<TH COLSPAN=2>Current Match Options</TH>
			</TR>
			<%
			Do While Not(oRS2.EOF)
				If oRs2.Fields("MapNumber").Value <= intMaps Then
					If bgc = bgcone Then
						bgc = bgctwo
					Else
						bgc = bgcone
					End If
					Response.Write "<TR BGCOLOR=""" & bgc & """>"
					CurrentMap = ""
					If oRs2.Fields("MapNumber").Value <> 0 Then
						CurrentMap = " on " & MapArray(cInt(oRs2.Fields("MapNumber").Value))
					Else 
						CurrentMap = ""
					End If
						
					Select Case(oRS2.Fields("SelectedBy").Value)
						Case "A", "C"
							Response.Write "<TD>You choose " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
						Case "D"
							If oRS2.Fields("SideChoice").Value = "Y" Then
								Response.Write "<TD>You have " & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
							Else
								Response.Write "<TD>" & enemyName & " selected " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
							End If
						Case "R"
							If oRS2.Fields("SideChoice").Value = "Y" Then
								Response.Write "<TD>" & TeamName & strVerbiage & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & "</TD>"
							Else
								Response.Write "<TD>TWL choose " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
							End If
						Case Else
							Response.Write "Error"
					End Select
					Response.Write "</TR>"
				End if
				oRS2.MoveNext
			Loop
			%>
			</TABLE>
			<%
		End If
		oRS2.NextRecordset 
		' End Map Selection Options
		Response.Write "<br><table width=45% align=center ><tr bgcolor="&bgcone&" height=35 valign=middle><td align=center>[ <a href=MatchReportLoss.asp?matchid=" & matchid & "&teamid=" & teamid & ">Report Loss</a> ]&nbsp;&nbsp;&nbsp;&nbsp; "
		Response.Write " [ <a href=DisputeMatch.asp?matchid=" & matchid & "&DisputeTeamID=" & linkid & "&ladder=" & Server.URLEncode(LadderName) & ">Dispute Match Results</a> ]</td></tr></table>"
	end if	
	ors.Close 	

elseif mstatus="Defending" then
		strSQL = "select * FROM vMatches where MatchDefenderID = " & linkid & " and matchladderid=" & ladderid
		ors.Open strSQL, oconn
		if not (ors.EOF and ors.BOF) then
			enemyID = ors.Fields("MatchAttackerID").Value
			mDate = ors.Fields("MatchDate").Value 
			xDate = ors.Fields("MatchChallengeDate").Value 
			aDate = ors.Fields("MatchAcceptanceDate").Value 
			enemyname = ors.Fields("Attackername").Value
			map1=ors.Fields("MatchMap1ID").value
			map2=ors.Fields("MatchMap2ID").value
			map3=ors.Fields("MatchMap3ID").value
			MapArray(1) = ors.Fields("MatchMap1ID").Value 
			MapArray(2) = ors.Fields("MatchMap2ID").Value 
			MapArray(3) = ors.Fields("MatchMap3ID").Value 
			MapArray(4) = ors.Fields("MatchMap4ID").Value 
			MapArray(5) = ors.Fields("MatchMap5ID").Value 
			matchdate1 = ors.Fields("MatchSelDate1").Value
			matchdate2= ors.Fields("MatchSelDate2").Value 
			matchid = ors.Fields("MatchID").Value 
		end if

		Response.Write "<center><b>Current Status:</b> Defending vs <a href=viewteam.asp?team=" & server.urlencode(enemyname) & ">"&Server.HTMLEncode(enemyname)&"</a><br />"
		Response.Write "Maps: "
		For i = 1 to Maps
			If i > 1 Then
				Response.Write ", "
			end if
			Response.Write MapArray(i)
		Next
		response.write " <br>Match Date: " & mdate & "</center>"
		if (bsysadmin or bLadderAdmin) and mdate <> "TBD" then %>
			<br><center><input value="Clear Match Date" type=button class=brightgold style="width:200" onclick="window.location.href='ladder/ladderengine.asp?clearDate=true&matchid=<%=matchid%>';" id=clearDate name=clearDate />
			<form name=changeDate action=ladder/ladderengine.asp method=post>
				<input type="hidden" name="saveType" value="changeDate">
				<input type="hidden" name="timezone" value="<%=RIght(mdate, 3)%>">
				<input type="hidden" name="matchid" value="<%=matchid%>">
				Month: <input type="Text" name="newMonth" value="" size="2" maxlength="2">&nbsp;Day:<input type="Text" name="newDay" value="" size="2" maxlength="2">&nbsp;<%=year(now)%><br>
				Year: <input type="text" name="newYear" id="newYear" value="<%=year(now)%>" maxlength="4" size="4"><br />
				New Time: <input type="Text" name="newHour" value="" size="2" maxlength="2">:<input type="Text" name="newMinute" value="" size="2" maxlength="2">:00 PM<br>
				<input type="submit" class=brightgold style="width:150" value="Change Date" id=submit4 name=submit4>
			</form></center><%
		end if
		response.write "<center><BR><b>Challenge Date:</b> " & xDate & ""
		response.write "<BR><b>Accept Date:</b> " & aDate & "</center>"
		%>
		<!-- #include virtual="/include/incruleinsert.asp" -->
		<%	
		if  (matchdate1 = "TBD") then
			Response.Write "<center><a href=acceptmatch.asp?team=" & server.urlencode(TeamName) & "&ladder=" & server.urlencode(LadderName) & "&matchid=" & matchID & "&enemy=" & server.urlencode(enemyname) & ">Accept the Challenge from " & Server.HTMLEncode(enemyname) & "</a>(You were challenged on " & xDate & ")</center>"
		else if mDate="TBD" then
			Response.Write "<center>You have chosen <b>" & matchdate1 & "</b> and <b>" & matchdate2 & "</b> for match dates.</center>"
			Response.Write "<center>Awaiting acceptance from " & Server.HTMLEncode(enemyname) & "(You were challenged on " & xDate & ")</center>"
		else
		
			strSQL = "SELECT * FROM vLadderOptions "
			strSQL = strSQL & " WHERE SelectedBy <> 'R' AND "
			strSQL = strSQL & " LadderID = '" & LadderID & "' AND "
			strSQL = strSQL & " OptionID NOT IN (SELECT mo.OptionID FROM lnk_match_options mo WHERE mo.MatchID = '" & matchid & "')"
			'Response.Write strSQL
			oRs2.Open strSQL, oConn
			If Not(oRS2.EOF AND oRS2.BOF) Then
				blnOptionShown = False
				%>
				<FORM NAME=frm_map_options ACTION="/ladder/option_saveitem.asp" METHOD="POST">
				<INPUT TYPE=HIDDEN NAME="MatchID" VALUE="<%=matchid%>">
				<INPUT TYPE=HIDDEN NAME="SaveType" VALUE="SaveMapOptions">
				<INPUT TYPE=HIDDEN NAME="LadderName" VALUE="<%=Server.HTMLEncode(LadderName & "")%>">
				<INPUT TYPE=HIDDEN NAME="TeamName" VALUE="<%=Server.HTMLEncode(TeamName & "")%>">
				
				<table width="75%" class="cssBordered" align="center">
				<TR BGCOLOR="#000000">
					<TH COLSPAN=2>Select Match Options</TH>
				</TR>
				<%
				intCounter = 0
				Do While Not(oRS2.EOF) And intCounter < 100
					intCounter = intCounter + 1
					intOptionID = oRS2.Fields("OptionID").Value 
					If bgc = bgcone Then
						bgc = bgctwo
					Else
						bgc = bgcone
					End If
					Response.Write "<TR BGCOLOR=""" & bgc & """>"
					CurrentMap = ""
					If oRs2.Fields("MapNumber").Value <> 0 Then
						CurrentMap = MapArray(cInt(oRs2.Fields("MapNumber").Value))
					Else 
						CurrentMap = "the match"
					End If
					
					Select Case(oRS2.Fields("SelectedBy").Value)
						Case "D"
							blnOptionSame = True
							Response.Write "<TD ALIGN=RIGHT>Choose your " & lCase(oRS2.Fields("OptionName").Value) & " for " & CurrentMap & ": </TD>"
							Response.Write "<TD><SELECT NAME=MO_" & oRS2.Fields("OptionID").Value & ">"
							While Not(oRS2.EOF) AND blnOptionSame
								Response.Write "<OPTION VALUE=""" & oRS2.Fields("OptionValueID").Value & """>" & oRS2.Fields("ValueName").Value & "</OPTION>" & vbCrLf
								oRS2.MoveNext
								If Not(oRS2.EOF) Then
									If oRS2.Fields("OptionID").Value = intOptionID Then
										blnOptionSame = True
									Else
										blnOptionSame = False
									End If
								End If
							Wend
							Response.Write "</SELECT></TD>"
							blnOptionShown = True
						Case "A", "C"
							blnOptionSame = True
							Response.Write "<TD COLSPAN=2>" & EnemyName & " will choose " & lcase(oRs2.Fields("OptionName").Value) & " for " & CurrentMap & "</TD>"
							While Not(oRS2.EOF) AND blnOptionSame
								oRS2.MoveNext
								If Not(oRS2.EOF) Then
									If oRS2.Fields("OptionID").Value = intOptionID Then
										blnOptionSame = True
									Else
										blnOptionSame = False
									End If
								End If
							Wend
						Case Else
							Response.Write "Error"
							oRS2.MoveNext 						
					End Select
					Response.Write "</TR>"
'					oRS2.MoveNext
				Loop
				If blnOptionShown Then
					Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=2 ALIGN=center><INPUT TYPE=SUBMIT VALUE=""Confirm Match Options""></TD></TR>"
				End If
				%>
				</TABLE>
				</FORM>
				<%
			End If
			oRS2.NextRecordset 
			' End Map Selection Options

			' Show Current Selected Options

			strSQL = "SELECT * FROM vMatchOptions "
			strSQL = strSQL & " WHERE MatchID = '" & matchid & "'" 
			'Response.Write strSQL
			oRs2.Open strSQL, oConn
			If Not(oRS2.EOF AND oRS2.BOF) Then
				blnOptionShown = False
				%>
				<table width="75%" class="cssBordered" align="center">
				<TR BGCOLOR="#000000">
					<TH COLSPAN=2>Current Match Options</TH>
				</TR>
				<%
				Do While Not(oRS2.EOF)
					If oRs2.Fields("MapNumber").Value <= intMaps Then
						If bgc = bgcone Then
							bgc = bgctwo
						Else
							bgc = bgcone
						End If
						Response.Write "<TR BGCOLOR=""" & bgc & """>"
						CurrentMap = ""
						If oRs2.Fields("MapNumber").Value <> 0 Then
							CurrentMap = " on " & MapArray(cInt(oRs2.Fields("MapNumber").Value))
						Else 
							CurrentMap = ""
						End If
						
						Select Case(oRS2.Fields("SelectedBy").Value)
							Case "D"
								Response.Write "<TD>You choose " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
							Case "A", "C"
								If oRS2.Fields("SideChoice").Value = "Y" Then
									Response.Write "<TD>You have " & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
								Else
									Response.Write "<TD>" & enemyName & " selected " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
								End If
							Case "R"
								If oRS2.Fields("SideChoice").Value = "Y" Then
									Response.Write "<TD>" & TeamName & strVerbiage & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & "</TD>" 
								Else
									Response.Write "<TD>TWL choose " & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value) & CurrentMap & ".</TD>"
								End If
							Case Else
								Response.Write "Error"
						End Select
						Response.Write "</TR>"
					End If
					oRS2.MoveNext
				Loop
				%>
				</TABLE>
				<%
			End If
			oRS2.NextRecordset 
			' End Map Selection Options
			
		Response.Write "<br><table width=45% align=center ><tr bgcolor="&bgcone&" height=35 valign=middle><td align=center>[ <a href=MatchReportLoss.asp?matchid=" & matchid & "&teamid=" & teamid & ">Report Loss</a> ]&nbsp;&nbsp;&nbsp;&nbsp; "
		Response.Write " [ <a href=DisputeMatch.asp?matchid=" & matchid & "&DisputeTeamID=" & linkid & "&ladder=" & Server.URLEncode(LadderName) & ">Dispute Match Results</a> ]</td></tr></table>"
		end if
	end if
	ors.Close 
elseif (mstatus="Available" or left(mStatus,6)="Immune") then
	strsql= "select count(*) from lnk_T_P_L where TLLinkID ='" & linkid & "'"
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		players = ors.fields(0).value
	end if
	ors.close
	If IsNull(MinPlayer) then
		MinPlayer = 0
	End If

	if (LadderLocked = 0) and players >= minplayer AND ((intChallengeDays AND 2^(WeekDay(Now()))) = 2^(WeekDay(Now()))) then
		MaxJump=int(tRank/2) + 1
		if MaxJump < (LadderChallenge + 1) then
			MaxJump = (LadderChallenge + 1)	
		end if
		strSQL = "SELECT lnk_T_L.Status, tbl_Teams.TeamName, lnk_T_L.Rank, tbl_Ladders.LadderName, lnk_T_L.LadderMode "
		strSQL = strSQL & "FROM tbl_Ladders INNER JOIN (tbl_Teams INNER JOIN lnk_T_L ON tbl_Teams.TeamID = lnk_T_L.TeamID) ON "
		strSQL = strSQL & "tbl_Ladders.LadderID = lnk_T_L.LadderID WHERE (((lnk_T_L.Status='Available' or left(lnk_T_L.Status,8)='Defeated') AND "
		strSQL = strSQL & "((lnk_T_L.Rank)<" & tRank &") AND ((lnk_T_L.Rank)>" & (tRank - MaxJump) &") AND ((tbl_Ladders.LadderName)='" & Request.QueryString("ladder") & "') and lnk_T_L.LadderMode<2 and lnk_t_l.isactive=1)) order by lnk_T_L.Rank"

		strSQL = "select lnk.Status, t.TeamName, lnk.Rank, l.LadderName, lnk.LadderMode, lnk.TLLinkID "
		strSQL = strSQL & "from lnk_t_l lnk, tbl_ladders l, tbl_teams t "
		strSQL = strSQL & "where t.teamid = lnk.teamid AND "
		strSQL = strSQL & "lnk.ladderid = l.ladderid AND "
		strSQL = strSQL & "(lnk.status = 'Available' or left(lnk.Status,8)='Defeated') AND "
		strSQL = strSQL & "lnk.rank < " & tRank & " AND lnk.rank > " & (tRank - MaxJump) & " AND l.LadderName = '" & CheckString(Request.QueryString("ladder")) & "' AND "
		strSQL = strSQL & "lnk.LadderMode < 2  AND lnk.isActive = 1  "
		strSQL = strSQL & "AND " 
		strSQL = strSQL & "lnk.TLLinkID NOT in ( "
		strSQL = strSQL & "select TLLinkID = MatchWinnerID from tbl_history where MatchLoserID = '" & linkid & "' AND DateDiff(d, matchDate, GetDate()) < 7 "
		strSQL = strSQL & "union "
		strSQL = strSQL & "select TLLinkID = MatchLoserID from tbl_history where MatchWinnerID= '" & linkid & "' AND DateDiff(d, matchdate, GetDate()) < 7 "
		strSQL = strSQL & ") "
		strSQL = strSQL & "order by lnk.rank "
		ors.Open strSQL, oconn
		Response.Write "<table align=center border=0 cellspacing=0><tr><td bgcolor="&bgcone&" align=center height=22 width=250><b>" & mstatus & " - Click to Challenge</b></td></tr>"
		bgc=bgctwo
		if not (ors.EOF and ors.BOF) then
			do while not ors.EOF
				Response.Write "<tr bgcolor=" & bgc & " height=20><td align=center height=22 width=250><a href=challenge.asp?opponent=" & server.urlencode(ors.Fields(1).Value) & "&ladder=" & server.urlencode(Request.QueryString("ladder")) & "&team=" & server.urlencode(Request.QueryString("team")) & ">" & Server.HTMLEncode(ors.Fields(1).Value) & "</a> - Rung: " & ors.Fields(2).Value & "</td></tr>"		
				if bgc = bgcone then
					bgc=bgctwo
				else
					bgc=bgcone
				end if
				ors.MoveNext
			loop
		end if
		ors.Close 
		Response.Write "</table>"
	else
		if LadderLocked = 1 then 
			Response.Write "<center<b><i>This ladder is not open for challenging at this time.</i></b></center>"
		end if
		if Players <= minplayer then
			Response.Write "<center><b><i>Your roster is smaller than " & minplayer & " people, you cannot challenge another team at this time.</i></b></center>"
		end if
		If NOT((intChallengeDays AND 2^(WeekDay(Now()))) = 2^(WeekDay(Now()))) Then
			Response.Write "<center><b><i>Today is not a valid day of the week for challenging on this ladder. <br />Valid day(s) are: </i></b></center>"
			For i = 1 to 7
				If  intChallengeDays AND 2^i Then
					Response.Write WeekDayName(i) & " " 
				End If
			Next
		End If
	end if
		
		strSQL="select LadderID, restrank from tbl_ladders where laddername='" & replace(Request.QueryString("ladder"), "'", "''") & "'"
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			lid=ors.Fields(0).Value
			rRank = oRs.Fields("RestRank").Value
		end if
		ors.Close
		strsql="select restdays, tbl_teams.teamid from lnk_T_L inner join tbl_teams on lnk_T_L.teamid=tbl_teams.teamid where tbl_teams.teamname='" & replace(Request.QueryString("team"), "'", "''") & "' and lnk_t_l.ladderid=" & lid
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			rdays=ors.Fields(0).Value
			tid=ors.Fields(1).Value 
		end if
		Response.Write "<table align=center border=0 cellspacing=0 width=540 >"
		ors.Close 
		if rdays=1 then
			dayprint=" day "
		else
			dayprint=" days "
		end if
		if rdays < 1 then 
			rdays = 0
		end if
		Response.Write "<tr><td>&nbsp;</td></tr><tr bgcolor=" & bgc & "><td align=center>You have used " & rdays & dayprint & "of rest and have " & (7 - rdays) & " remaining this quarter</td></tr>"
		if bgc = bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
		
		if (rdays < 7  and trank > cint(rRank)) then
			Response.Write "<tr bgcolor=" & bgc & "><td align=center><a href=saveitem.asp?SaveType=Rest&teamid=" & tid & "&ladderid=" & lid & "&" & Request.QueryString & ">Go on rest with " & Server.HTMLEncode(Request.QueryString("team")) & " on the " & Server.HTMLEncode(Request.QueryString("ladder")) & " ladder</a></td></tr>"
		else
			if rdays > 6 then
				Response.Write "<tr bgcolor=" & bgc & "><td align=center>You are unable to rest at this time because you have used up the allocated days on the " & Server.HTMLEncode(Request.QueryString("ladder")) & " ladder</a></td></tr>"
				if bgc = bgcone then
					bgc=bgctwo
				else
					bgc=bgcone
				end if
			end if
			if trank < cint(rRank) then
				Response.Write "<tr bgcolor=" & bgc & "><td align=center>You are unable to rest at this time because you are ranked #" & trank & " on the " & Server.HTMLEncode(Request.QueryString("ladder")) & " ladder</a></td></tr>"
			end if
		end if
		Response.Write "</table>"
elseif  mstatus="Resting" then
		strSQL="select LadderID from tbl_ladders where laddername='" & replace(Request.QueryString("ladder"), "'", "''") & "'"
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			lid=ors.Fields(0).Value
		end if
		ors.Close
		strsql="select restdays, tbl_teams.teamid from lnk_T_L inner join tbl_teams on lnk_T_L.teamid=tbl_teams.teamid where tbl_teams.teamname='" & replace(Request.QueryString("team"), "'", "''") & "' and lnk_t_l.ladderid=" & lid
		ors.Open strsql, oconn
		'Response.write strsql
		if not (ors.EOF and ors.BOF) then
			rdays=ors.Fields(0).Value
			tid=ors.Fields(1).Value 
		end if
		ors.Close 
		Response.Write "<table align=center border=0 cellspacing=0><tr><td bgcolor="&bgcone&" align=center height=22 width=250><b>" & Server.HTMLEncode(Request.QueryString("team")) & " are currently " & mstatus & "</b></td></tr></table>"
		Response.Write "<table align=center border=0 cellspacing=0 width=540>"
		if rdays=1 then
			dayprint=" day "
		else
			dayprint=" days "
		end if
		if bgc = bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
		if rdays<1 then
			rdays=0
		end if
		Response.Write "<tr><td>&nbsp;</td></tr><tr bgcolor=" & bgc & "><td align=center>You have used " & rdays & dayprint & "of rest and have " & (7 - rdays) & " remaining this quarter</td></tr>"
		if bgc = bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
		Response.Write "<tr bgcolor=" & bgc & "><td align=center><a href=saveitem.asp?SaveType=UnRest&teamid=" & tid & "&ladderid=" & lid & "&" & Request.QueryString & ">Come off rest with " & Server.HTMLEncode(Request.QueryString("team")) & " on the " & Server.HTMLEncode(Request.QueryString("ladder")) & " ladder</a></td></tr>"
		Response.Write "</table>"
elseif left(mStatus,6)="Defeat" then
	strsql = "select modeflagtime from lnk_T_L where tllinkid=" & linkID
	ors.open strsql, oconn
	tima = ors.fields(0).value
	ors.close		
		Response.Write "<center><b>Current Status:</b> " & mstatus & "</center>"
		response.write "<center><b>Loss reported at: " & tima & "</center>"
		response.write "<center><b>Able to challenge up 24-hours after report time.</center>"
else
		Response.Write "<center><b>Current Status:</b> " & mstatus & "</center>"
end if
if mStatus = "Attacking" or mStatus="Defending" then
	Response.Write "<a name=""matchcomms""></a><br><br><center><b><u>Match Communications</u></b></center>"
	strSQL = "select MatchID from tbl_Matches where MatchDefenderID = " & LinkID& " or MatchAttackerID=" & LinkID
	'ors.Close 
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then 
		matchID = ors.Fields(0).Value 
	end if
	ors.Close
	if bTeamCaptain or bTeamFounder then
		Response.Write "<br><center><input type=button name=matchcomm value='Add Match Communication' onclick=""window.location.href='matchcomms.asp?matchid=" & matchid & "&mode=add&tag=" & server.urlencode(teamtag) & "&ladder=" & server.urlencode(LadderName) & "&team=" & server.urlencode(TeamName) & "';"">"
	else
		Response.Write "<br><center><input type=button name=matchcomm value='Add Match Communication' onclick=""window.location.href='matchcomms.asp?matchid=" & matchid & "&mode=add&tag=TWLAdmin&ladder=" & server.urlencode(LadderName) & "&team=" & server.urlencode(TeamName) & "';"">"
	end if
	strSQL = "select * from tbl_Comms where ((matchID='" & matchID & "') and (CommDead=0)) order by CommID desc"
	ors.Open strSQL, oconn
	%>
	<table align=center width=580 border=0 cellspacing=0 cellpadding=1>
	<%
	bgc=bgcone
	if not (ors.EOF and ors.bof) then
		do while not ors.EOF
			Response.Write "<tr bgcolor="& bgc & "><td colspan=2><hr></td></tr><tr bgcolor="& bgc & "><td>Author: <b>" & ors.Fields(2).Value & " - Posted: " & FormatDateTime(ors.Fields(1).Value, 0) & "</td>"
			if bSysAdmin then
				Response.Write "<td align=right><a href=matchcomms.asp?matchid=" & matchID & "&mode=edit&ladder=" & server.urlencode(LadderName) & "&team=" & server.urlencode(TeamName) & "&commid=" & ors.Fields(3).Value & ">Edit</a> - <a href=SaveItem.asp?commid=" & ors.Fields(3).Value & "&SaveType=Delete_Communications&ladder=" & server.urlencode(LadderName) & "&team=" & server.urlencode(TeamName) & ">Delete&nbsp;&nbsp;</a></td></tr>"
			else
				Response.Write "<td>&nbsp;</td></tr>"
			end if
			Response.write "<tr bgcolor="& bgc &"><td colspan=2>" & Replace(ors.Fields(4).Value, vbCrLf, "<br />" & vbCrLf) & "&nbsp;</td></tr>"
			if bgc = bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.MoveNext
		loop
	end if
	%>
	</table>
	<%
	ors.Close 
end if

Call ContentEnd()
Call Content3BoxStart("Ladder Captain Management")
strSQL = "select tbl_players.playerhandle from tbl_players inner join lnk_T_P_L on lnk_T_P_L.PlayerID=tbl_Players.playerid where (lnk_T_P_L.TLLinkID='" & linkid & "' and lnk_T_P_L.isadmin=1) order by tbl_players.playerhandle"
ors.Open strSQL, oconn
Response.Write "<table border=0 align=center width=97% cellspacing=0><tr height=30 bgcolor="&bgcone&"><td align=center><b>Current Captains</b></td></tr>"
if not (ors.EOF and ors.BOF) then
	bgc=bgcone
	do while not ors.EOF
		info=""
		if ors.fields(0).value = ownername then 
			info=" (founder)"
		end if	
		if bgc=bgctwo then
			bgc=bgcone
		else bgc=bgctwo
		end if
		Response.Write "<tr height=30 bgcolor="&bgc&"><td align=center>" & Server.HTMLEncode(ors.Fields(0).Value) & info & "</td></tr>"
		ors.MoveNext			
	loop
end if
Response.Write "</table>"
Call Content3BoxMiddle1()
ors.close
strSQL = "select tbl_players.playerhandle, lnk_T_P_L.TPLLinkID from tbl_players inner join lnk_T_P_L on lnk_T_P_L.PlayerID=tbl_Players.playerid where (lnk_T_P_L.TLLinkID='" & linkid & "' and lnk_T_P_L.isadmin=0) order by tbl_players.playerhandle"
ors.Open strSQL, oconn
if not(ors.EOF and ors.BOF) then
	Response.Write "<form name=promote action=saveitem.asp method=post><table align=center border=0 width=97% cellspacing=0><tr bgcolor="&bgcone&" height=30><td align=center><b>Promote Player to Captain</b></td></tr>"
	Response.Write "<tr bgcolor="&bgctwo&"><td height=30 align=center><select name=playerlist style='width:150'>"
	do while not ors.EOF
		Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value)
		ors.MoveNext
	loop
	Response.Write "</select></td></tr><tr height=30 bgcolor="&bgcone&"><td align=center>"
	%>
	<input class=bright type=submit value='Promote'>
	<input type=hidden name=SaveType value=PromoteCaptain>
	<input type=hidden name=ladder value="<%=Server.HTMLEncode(Request.QueryString("ladder"))%>">
	<input type=hidden name=team value="<%=Server.HTMLEncode(teamname)%>">
	</td></tr></table></form>
	<%
end if
ors.Close
Call Content3BoxMiddle2()
strSQL = "select tbl_players.playerhandle, lnk_T_P_L.TPLLinkID from tbl_players inner join lnk_T_P_L on lnk_T_P_L.PlayerID=tbl_Players.playerid where (lnk_T_P_L.TLLinkID='" & linkid & "' and lnk_T_P_L.isadmin=1 and tbl_players.playerhandle <> '" & CheckString(session("uName")) & "') order by tbl_players.playerhandle"
ors.Open strSQL, oconn
if not(ors.EOF and ors.BOF) then
	Response.Write "<form name=demote action=saveitem.asp method=post><table align=center border=0 width=97% cellspacing=0><tr bgcolor="&bgcone&" height=30><td align=center><b>Demote Captain</b></td></tr>"
	Response.Write "<tr bgcolor="&bgctwo&" height=30><td align=center><select name=playerlist style='width:150'>"
	do while not ors.EOF
		if ors.fields(0).value <> ownername then
			Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value)
		end if
		ors.MoveNext
	loop
	Response.Write "</select></td></tr><tr bgcolor="&bgcone&" height=30><td align=center><input class=bright type=submit id=submit2 name=submit2 value=Demote><input type=hidden name=SaveType value=DemoteCaptain><input type=hidden name=ladder value=""" & server.HTMLEncode(Request.QueryString("ladder") & "") & """><input type=hidden name=team value=""" & Server.HTMLEncode(Request.QueryString("team")) & """></td></tr></table></form>"	
end if
ors.Close

Call Content3BoxEnd()

Call ContentStart(Server.HTMLEncode(request("ladder")) & "Ladder Roster Management")

	strSQL="select PlayerHandle, lnk_T_P_L.DateJoined, lnk_T_P_L.IsAdmin, tbl_players.PlayerID from tbl_Players inner join "
	strsql= strsql & "lnk_T_P_L on lnk_T_P_L.PlayerID=tbl_players.playerid where lnk_T_P_L.TLLinkID=" 
	strsql= strsql & LinkID & " order by PlayerHandle"			
	ors.open strsql,oconn
	if not (ors.eof and ors.bof) then
		response.write "<form name=BootPlayer method=post action=saveitem.asp><table width=50% align=center border=0 cellspacing=0 cellpadding=0>"
		response.write "<tr bgcolor="&bgcone&" height=125 valign=center><td align=center><input type=hidden name=savetype value=DropPlayer><select name=PlayerID size=5 class=brightred style='width:200'>"
		do while not ors.eof
			if ors.fields(3).value <> ownerid then 
				response.write "<option value=" & ors.fields(3).value & ">" & Server.HTMLEncode(ors.fields(0).value)
			end if
			ors.movenext
		loop
		response.write "</select></td></tr><tr bgcolor="&bgctwo&" height=35><td align=center><input type=hidden name=link value=" & linkid & "><input type=submit class=bright style='width:75' value='Kick Player'></td></tr></table></form>"
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