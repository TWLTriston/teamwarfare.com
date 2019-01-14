<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Roster - " & Replace(Request.Querystring("team"), """", "&quot;") 

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin, bTournamentAdmin 
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strTeamName, strLeagueName, intTeamID, intLeagueID, blnError, strError, intLeagueTeamID, intIdentifierID, intCols
Dim strIdentifierName, intTeamFounderID, strDateJoined, strStatus, strDateAdded, intPlayerID, intThisPlayerID
Dim strIDs
strTeamName = Trim(Request.QueryString("Team"))
strLeagueName = Trim(Request.QueryString("League"))
blnError = False
strError = "Error:<br />"
intCols = 4
strSQL = "SELECT TeamFounderID, TeamID FROM tbl_teams WHERE TeamName = '" & CheckString(strTeamName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) THen
	intTeamID = oRs.Fields("TeamID").Value
	intTeamFounderID = oRs.Fields("TeamFounderID").Value
Else
	strError = strError & "Unable to find specified team.<br />"
	blnError = True
End If
oRs.NextRecordSet

strSQL = "SELECT LeagueID, IdentifierID FROM tbl_Leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) THen
	intLeagueID = oRs.Fields("LeagueID").Value
	intIdentifierID = oRs.Fields("IdentifierID").Value
	intCols = intCols + 2
Else
	strError = strError & "Unable to find specified League.<br />"
	blnError = True
End If
oRs.NextRecordSet

If Not(blnError) Then
	strSQL = "SELECT lnkLeagueTeamID FROM lnk_league_team WHERE LeagueID = '" & intLeagueID & "' AND TeamID = '" & intTeamID & "'"
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) THen
		intLeagueTeamID = oRs.Fields("lnkLeagueTeamID").Value
	Else
		strError = strError & "Unable to find specified information.<br />"
		blnError = True
	End If
	oRs.NextRecordSet
	
	If intIdentifierID > 0 Then
		strSQL = "SELECT IdentifierName FROM tbl_identifiers WHERE IdentifierID = '" & intIdentifierID & "'"
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRs.BOF) THen
			strIdentifierName = oRs.Fields("IdentifierName").Value
		End If
		oRs.NextRecordSet
	End If
End if
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("")
If blnError Then
	Response.Write strError
Else
	%>
	<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
	<tr>
		<td>
			<table border="0" cellspacing="1" cellpadding="4" width="100%">
			<tr>
				<td colspan="<%=intCols%>" bgcolor="#000000"><a href="viewteam.asp?team=<%=Server.URLEncode(strTeamName & "")%>">&laquo; back to team</a></td>
			</tr>
			<tr>
				<th colspan="<%=intCols%>" bgcolor="#000000"><%=Server.HTMLEncode(strTeamName & " Roster on the " & strLeagueName & " League")%></th>
			</tr>
			<tr>
				<th bgcolor="#000000">Player</th>
				<th bgcolor="#000000">Email Address</th>
				<th bgcolor="#000000">Status</th>
				<th bgcolor="#000000">Join Date</th>
				<% If intIdentifierID > 0 Then %>
				<th bgcolor="#000000"><%=strIdentifierName%></th>
				<% End If %>
			</tr>
			<%
			strSQL = "SELECT p.PlayerID, PlayerHandle, PlayerEmail, PlayerHideEmail, JoinDate, iSAdmin "
			If intIdentifierID > 0 Then
				strSQL = strSQL & ", IdentifierValue, DateAdded "
			End If
			strSQL = strSQL & " FROM lnk_league_team_player l, tbl_players p "
			If intIdentifierID > 0 Then
				strSQL = strSQL & ", lnk_player_identifier pi "
			End If
			strSQL = strSQL & " WHERE lnkLeagueTeamID = '" & intLeagueTeamID & "' AND l.PlayerID = p.PlayerID "
			If intIdentifierID > 0 Then
				strSQL = strSQL & " AND pi.IdentifierID = '" & intIdentifierID & "' AND l.PlayerID *= pi.PlayerID  AND pi.IdentifierActive = 1 "
			End If
			strSQL = strSQL & " ORDER BY PlayerHandle ASC "
			'Response.Write strSQL
			oRs.Open strSQL, oConn, 1, 3
			If Not(oRS.EOF AND oRs.BOF) THen
				bgc = bgcone
				Do While Not(oRs.EOF)
					If CStr(oRs.Fields("PlayerID").Value) = CStr(intTeamFounderID) Then
						strStatus = "Team Founder"
					ElseIf oRs.Fields("IsAdmin").value Then
						strStatus = "Team Captain"
					Else
						strStatus = "&nbsp;"
					End If
					If Len(oRs.Fields("JoinDate").Value) < 8 Then
						strDateJoined = "&nbsp;"
					Else
						strDateJoined = Formatdatetime(oRs.Fields("JoinDate").Value,2)
					End If

					%>
					<tr>
						<td bgcolor="<%=bgc%>" valign="top"><a href="viewplayer.asp?player=<%=Server.URLEncode(oRs.Fields("PlayerHandle").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("PlayerHandle").Value & "")%></a></td>
						<% If Session("LoggedIn") AND (oRs.Fields("IsAdmin").Value OR bSysAdmin) Then %>
						<td bgcolor="<%=bgc%>" valign="top"><%=Replace(Replace("" & oRs.Fields("PlayerEmail").Value, "@", " at "), ".", " dot ")%></td>
						<% Else %>
						<td bgcolor="<%=bgc%>" valign="top">not available</td>
						<% End If %>
						<td bgcolor="<%=bgc%>" valign="top"><%=strStatus%></td>
						<td bgcolor="<%=bgc%>" align="right" valign="top"><%=strDateJoined%></td>
						<% If intIdentifierID > 0 Then 
							intPlayerID = oRs.FieldS("PlayerID").Value
							If Not(IsNull(oRs.Fields("DateAdded").Value)) Then
								strDateAdded = " (" & FormatDateTime(oRs.Fields("DateAdded").Value, 2) & ")"
							Else 
								strDateAdded = ""
							End If
							strIDs = oRs.Fields("IdentifierValue").Value & strDateAdded
							oRs.MoveNext
							If Not(oRs.EOF) Then
								If CStr(oRs.Fields("PlayerID").Value) = CStr(intPlayerID) Then
									intThisPlayerID = oRs.Fields("PlayerID").Value
									Do While Not(oRs.EOF) AND intThisPlayerID = intPlayerID
										If Not(IsNull(oRs.Fields("DateAdded").Value)) Then
											strDateAdded = " (" & FormatDateTime(oRs.Fields("DateAdded").Value, 2) & ")"
										Else 
											strDateAdded = ""
										End If
										strIDs = strIDs & "<br />" & oRs.Fields("IdentifierValue").Value & strDateAdded
										oRs.MoveNext
										If Not(oRs.EOF) Then
											intThisPlayerID = oRs.Fields("PlayerID").Value
										End If
									Loop
								End If
							End If
							%>
						<td bgcolor="<%=bgc%>" align="right"><%=strIDs%></td>
						<% End If %>
					</tr>
					<%
					If bgc = bgcone then
						bgc = bgctwo
					Else
						bgc = bgcone
					End if
					If intIdentifierID = 0 OR IsNull(intIdentifierID) Then 
						oRs.MoveNext
					End If
				Loop
			End If
			oRs.nextRecordSet
			%>
			</table>
		</td>
	</tr>
	</table>
				
	<%	
End If
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>