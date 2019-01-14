<%
Option Explicit
%>
<!-- #include virtual="/include/xml.asp" -->
<%
Server.ScriptTimeout = 45
Dim strTeamName
Dim oRS, strSQL, oConn, oRS2

Dim intMatchID, strEnemyName, strMatchDate
Dim strStatus, intMaps, i
Dim intCurrentMapNumber, strCurrentMap
Dim strMapArray(6)

'Leagues
Dim strLeagueName, intLeagueID, intConferenceID, strConferenceName, intDivisionID, strDivisionName
Dim intLeagueTeamID, intLeagueWins, intLeagueLosses, intTeamLinkID
Dim intLeagueDraws, intLeaguePoints, intLeagueWinPct, intLeagueRank, intLeagueRoundsOne, intLeagueRoundsLost

Dim strTeamStatus, strURL, strFounderName ' Team Information
Dim arrMapOptions(6), itemMapOption ' Scripting dictionary to produce cool xml output

strTeamName = Request.Form("Team")

if Len(strTeamName) = 0 then
	strTeamName = Request.QueryString("Team")
End if
If (strTeamName = "") Then
	Response.Write "You must specify a team in the querystring, in the form of: filename.asp?team=Cabal+Kabob"
	Response.End 
End If

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

Dim intTeamID
For i = lBound(arrMapOptions) To uBound(arrMapOptions)
	Set arrMapOptions(i) = Server.CreateObject("Scripting.Dictionary")
Next

Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf

	Response.Write "<team>" & vbCrLf
	strSQL = "SELECT t.*, p.PlayerHandle "
	strSQL = strSQL & " FROM tbl_teams t, tbl_players p "
	strSQL = strSQL & " WHERE t.TeamName = '" & Replace(strTeamName, "'", "''") & "' AND t.TeamFounderID *= p.PlayerID "
	oRS.Open strSQL, oConn
	If Not(ors.eof and ors.BOF) Then
		' Give a synopsis of the team's profile
		strFounderName = oRS.Fields("PlayerHandle").Value
		strURL = oRS.Fields("TeamURL").Value
		intTeamID = oRs.Fields("teamID").Value
		If ucase(left(strURL,4))<> "HTTP" AND Len(strURL) > 0 Then
			strURL = "http://" & strURL
		End If

		If ors.Fields("TeamActive").Value= "1" then
			strTeamStatus = "Active"
		Else
			strTeamStatus = "Inactive"
		End If

		Response.Write vbTab & "<teaminformation>" & vbCrLf
		Response.Write vbTab & vbTab & "<name>" & XMLEncode(Server.HTMLEncode(oRS.Fields("TeamName").Value & "")) & "</name>" & vbCrLF 
		Response.Write vbTab & vbTab & "<url>" & XMLEncode(Server.HTMLEncode(strURL & "")) & "</url>" & vbCrLF 
		Response.Write vbTab & vbTab & "<email>" & XMLEncode(Server.HTMLEncode(oRS.Fields("TeamEmail").Value & "")) & "</email>" & vbCrLF 
		Response.Write vbTab & vbTab & "<irc channel=""" & XMLEncode(Server.HTMLEncode(oRS.Fields("TeamIRC").Value & "")) & """>" & XMLEncode(Server.HTMLEncode(oRS.Fields("TeamIRCServer").Value & "")) & "</irc>" & vbCrLF 
		Response.Write vbTab & vbTab & "<status>" & XMLEncode(Server.HTMLEncode(strTeamStatus & "")) & "</status>" & vbCrLF 
		Response.Write vbTab & vbTab & "<founder>" & XMLEncode(Server.HTMLEncode(strFounderName & "")) & "</founder>" & vbCrLF 
		Response.Write vbTab & vbTab & "<description>" & XMLEncode(Server.HTMLEncode(oRS.Fields("TeamDesc").Value & "")) & "</description>" & vbCrLF 
		Response.Write vbTab & "</teaminformation>" & vbCrLf
	End If
	ors.NextRecordset
	
	' Start competition search
	Response.Write vbTab & "<competitioninformation>" & vbCrLf
	
	' Ladders
	strSQL = "SELECT * FROM vLadder WHERE TeamName = '" & Replace(strTeamName, "'", "''") & "'"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF And oRS.BOF) Then
		Do While Not(oRS.EOF)
			' Clear the options dictionaries
			For i = lBound(arrMapOptions) To uBound(arrMapOptions)
				Call arrMapOptions(i).RemoveAll()
			Next

			Response.Write vbTab & vbTab & "<ladder name=""" & XMLEncode(Server.HTMLEncode(oRS.Fields("LadderName").Value)) & """>" & vbCrLf
'			Response.Write vbTab & vbTab & "<ladder>" & vbCrLf

			' Find their current status, so we can show match options accordingly		
			strStatus	= oRS.Fields("Status").Value 
			' Number of maps definded for this ladder, used to find out how many maps to display
			intMaps		= oRS.Fields("Maps").Value 
			
'			Response.Write vbTab & vbTab & vbTab & "<name>" & Server.HTMLEncode(oRS.Fields("LadderName").Value) & "</name>" & vbCrLf
			
			' Add some links to the XML, easier for the parsers
			Response.Write vbTab & vbTab & vbTab & "<ladderlink>" & vbCrLf 
			Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>" & "http://www.teamwarfare.com/viewladder.asp?ladder=" & XMLEncode(Server.UrlEncode(oRS.Fields("LadderName").Value)) & "</httplink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>" & "http://www.teamwarfare.com/xml/viewladder.asp?ladder=" & XMLEncode(Server.UrlEncode(oRS.Fields("LadderName").Value)) & "</xmllink>" & vbCrLf
'			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>n/a</xmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "</ladderlink>" & vbCrLf 
			
			' Current rank, and record
			Response.Write vbTab & vbTab & vbTab & "<rank>" & oRS.Fields("Rank").Value & "</rank>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<wins>" & oRS.Fields("Wins").Value & "</wins>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<losses>" & oRS.Fields("Losses").Value & "</losses>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "<forfeits>" & oRS.Fields("Forfeits").Value & "</forfeits>" & vbCrLf
			
			' Match Information
			' Write out status as an attribute
			Select Case Left(uCase(strStatus), 6)
				Case "DEFEND", "ATTACK"
					If  Left(uCase(strStatus), 6)  = "DEFEND" Then
						Response.Write vbTab & vbTab & vbTab & "<match status=""Defending"">" & vbCrLf
	
						strSQL = "select m.MatchAttackerID, m.MatchID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID, m.MatchMap4ID, m.MatchMap5ID, t.teamname, t.teamtag "
						strSQL = strSQL & " from tbl_Matches m, tbl_teams t, lnk_t_l lnk "
						strSQL = strSQL & " where m.matchdefenderID = " & oRS.Fields("TLLinkID").Value  
						strSQL = strSQL & " AND t.teamid = lnk.teamid "
						strSQL = strSQL & " AND lnk.tllinkid = m.MatchAttackerID "
						oRS2.Open strSQL, oconn
						If not (oRS2.EOF and oRS2.BOF) then
							intMatchID = oRS2.Fields("MatchID").Value 
							strEnemyName = oRS2.Fields("TeamName").Value
							strMatchDate = oRS2.Fields("MatchDate").Value
							strMapArray(1) = oRS2.fields("MatchMap1ID").value
							strMapArray(2) = oRS2.fields("MatchMap2ID").value
							strMapArray(3) = oRS2.fields("MatchMap3ID").value
							strMapArray(4) = oRS2.fields("MatchMap4ID").value
							strMapArray(5) = oRS2.fields("MatchMap5ID").value
						end if
						oRS2.NextRecordset 
					ElseIf Left(uCase(strStatus), 6)  = "ATTACK" Then				
						Response.Write vbTab & vbTab & vbTab & "<match status=""Attacking"">" & vbCrLf

						strsql = "select m.MatchDefenderID, m.MatchID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap4ID, m.MatchMap5ID, m.MatchMap3ID, t.teamname, t.teamtag "
						strsql = strsql & " from tbl_Matches m, tbl_teams t, lnk_t_l lnk "
						strsql = strsql & " where m.MatchAttackerID = " & oRS.Fields("TLLinkID").Value 
						strsql = strsql & " AND t.teamid = lnk.teamid "
						strsql = strsql & " AND lnk.tllinkid = m.matchdefenderID "
						oRS2.Open strSQL, oconn
						if not (oRS2.EOF and oRS2.BOF) then
							intMatchID = oRS2.Fields("MatchID").Value 
							strEnemyName = oRS2.Fields("TeamName").Value
							strMatchDate = oRS2.Fields("MatchDate").Value
							strMapArray(1) = oRS2.fields("MatchMap1ID").value
							strMapArray(2) = oRS2.fields("MatchMap2ID").value
							strMapArray(3) = oRS2.fields("MatchMap3ID").value
							strMapArray(4) = oRS2.fields("MatchMap4ID").value
							strMapArray(5) = oRS2.fields("MatchMap5ID").value
						end if
						oRS2.NextRecordset 
					End If
					' Match quick look
					Response.Write vbTab & vbTab & vbTab & vbTab & "<opponentname>" & XMLEncode(Server.HTMLEncode(strEnemyName)) & "</opponentname>" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & "<opponentlink>" & vbCrLf 
					Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<htmllink>http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLEncode(strEnemyName)) & "</htmllink>" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & XMLEncode(Server.URLEncode(strEnemyName)) & "</xmllink>" & vbCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & "</opponentlink>" & vbCrLf 
					Response.Write vbTab & vbTab & vbTab & vbTab & "<matchdate>" & strMatchDate & "</matchdate>" & vbCrLf
					
					' Display the maps and their options
					If  Left(uCase(strStatus), 6)  = "DEFEND" Then
						strSQL = "SELECT * FROM vMatchOptions "
						strSQL = strSQL & " WHERE MatchID = '" & intMatchID & "'" 
						'Response.Write strSQL
						oRs2.Open strSQL, oConn
						If Not(oRS2.EOF AND oRS2.BOF) Then
							intCurrentMapNumber = -1
							Do While Not(oRS2.EOF)
								intCurrentMapNumber = oRS2.Fields("MapNumber").Value
								On Error Resume Next
								Select Case(oRS2.Fields("SelectedBy").Value)
									Case "D"
										If oRS2.Fields("SideChoice").Value = "Y" Then
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("ValueName").Value)
										Else
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("ValueName").Value)
										End If
									Case "A"
										If oRS2.Fields("SideChoice").Value = "Y" Then
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("Opposite").Value)
										Else
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("ValueName").Value)
										End If
									Case "R"
										If oRS2.Fields("SideChoice").Value = "Y" Then
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("ValueName").Value)
										Else
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("ValueName").Value)
										End If
								End Select
								On Error Goto 0
								oRS2.MoveNext
							Loop
						End If	
						oRS2.NextRecordset 					
					ElseIf Left(uCase(strStatus), 6)  = "ATTACK" Then				
						strSQL = "SELECT * FROM vMatchOptions "
						strSQL = strSQL & " WHERE MatchID = '" & intMatchID & "'" 
						'Response.Write strSQL
						oRs2.Open strSQL, oConn
						If Not(oRS2.EOF AND oRS2.BOF) Then
							intCurrentMapNumber = -1
							Do While Not(oRS2.EOF)
								intCurrentMapNumber = oRS2.Fields("MapNumber").Value
								On Error Resume Next
								Select Case(oRS2.Fields("SelectedBy").Value)
									Case "A"
										If oRS2.Fields("SideChoice").Value = "Y" Then
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("ValueName").Value)
										Else
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("ValueName").Value)
										End If
									Case "D"
										If oRS2.Fields("SideChoice").Value = "Y" Then
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("Opposite").Value)
										Else
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("ValueName").Value)
										End If
									Case "R"
										If oRS2.Fields("SideChoice").Value = "Y" Then
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("Opposite").Value)
										Else
											Call arrMapOptions(intCurrentMapNumber).Add(Replace(oRS2.Fields("OptionName"), " ", ""), oRS2.Fields("ValueName").Value)
										End If
								End Select
								On Error Goto 0
								oRS2.MoveNext
							Loop
						End If
						oRS2.NextRecordset 					
					End If		
					
					' Spit back out the options			
					For i = 1 To intMaps
						If arrMapOptions(i).Count = 0 Then
							Response.Write vbTab & vbTab & vbTab & vbTab & "<map name=""" & XMLEncode(strMapArray(i)) & """ maporder=""" & i & """/>" & vbCrLf
						Else
							Response.Write vbTab & vbTab & vbTab & vbTab & "<map name=""" & XMLEncode(strMapArray(i)) & """ maporder=""" & i & """>" & vbCrLf
							Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<options>" & vbCrLf
							For Each itemMapOption In arrMapOptions(i)
								Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<" & XMLEncode(itemMapOption) & ">" & XMLEncode(Server.HTMLEncode(arrMapOptions(i)(itemMapOption))) & "</" & itemMapOption & ">" & vbCrLf
							Next
							Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "</options>" & vbCrLf
							Response.Write vbTab & vbTab & vbTab & vbTab & "</map>" & vbCrLf
						End If
					Next
					Response.Write vbTab & vbTab & vbTab & "</match>" & vbCrLf					
				' Case not in a match, just show the info
				Case "IMMUNE", "DEFEAT", "RESTIN"
					Response.Write vbTab & vbTab & vbTab & "<match status=""" & XMLEncode(strStatus) & """/>" & vbCrLf
				Case Else
					Response.Write vbTab & vbTab & vbTab & "<match status=""Open""/>" & vbCrLf
			End Select

			strSQL = "SELECT PlayerHandle, lnk_T_P_L.DateJoined, lnk_T_P_L.IsAdmin "
			strSQL = strSQL & " FROM tbl_Players inner join lnk_T_P_L on "
			strSQL = strSQL & " lnk_T_P_L.PlayerID=tbl_players.playerid "
			strSQL = strSQL & " WHERE lnk_T_P_L.TLLinkID=" & ors.Fields("TLLinkID").Value
			strSQL = strSQL & " ORDER BY PlayerHandle"
			oRS2.Open strSQL, oConn
			If Not(oRS2.EOF AND oRS2.BOF) Then
				Response.Write vbTab & vbTab & vbTab & "<roster>" & vbCrLf
				Do While Not(oRS2.EOF)
					Response.Write vbTab & vbTab & vbTab & vbTab & "<player name=""" & XMLEncode(Server.HTMLEncode(oRS2.Fields("PlayerHandle").Value)) & """>" & vbCrLf
					if len(ors2.Fields("DateJoined").Value) < 8 then
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<joindate>n/a</joindate>" & vbCrLf
					else
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<joindate>" & FormatDateTime(oRS2.Fields("DateJoined").Value, 2) & "</joindate>" & vbCrLf
					end if

					If Trim(oRS2.Fields("PlayerHandle").Value) = Trim(strFounderName) then
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<position>Team Founder</position>" & vbCrLf
					ElseIf ors2.Fields("IsAdmin").Value=1 then
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<position>Team Captain</position>" & vbCrLf
					Else
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<position>Player</position>" & vbCrLf
					End If
					Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<links>" & vBCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<htmllink>" & "http://www.teamwarfare.com/viewplayer.asp?player=" & XMLEncode(Server.URLEncode(oRS2.Fields("PlayerHandle").Value)) & "</htmllink>" & vBCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "</links>" & vBCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & "</player>" & vBCrLf
					oRS2.MoveNext 
				Loop
				Response.Write vbTab & vbTab & vbTab & "</roster>" & vbCrLf
			Else
				Response.Write vbTab & vbTab & vbTab & "<roster />" & vbCrLf
			End If
			oRS2.NextRecordset
						
			Response.Write vbTab & vbTab & "</ladder>" & vbCrLf
			oRS.MoveNext
		Loop
	End If
	oRS.NextRecordSet
	
	strSQL = "SELECT lnk.lnkLeagueTeamID, LeagueName, RosterLock, c.LeagueID, c.LeagueConferenceID, ConferenceName, lnk.LeagueDivisionID, "
	strSQL = strSQL & " Wins, Losses, Draws, LeaguePoints, WinPct, Rank, RoundsWon, RoundsLost "
	strSQL = strSQL & " FROM lnk_league_team lnk "
	strSQL = strSQL & " INNER JOIN tbl_leagues l "
	strSQL = strSQL & " ON l.LeagueID = lnk.LeagueID "
	strSQL = strSQL & " INNER JOIN tbl_league_conferences c "
	strSQL = strSQL & " ON c.LeagueConferenceID = lnk.LeagueConferenceID "
	strSQL = strSQL & " WHERE lnk.TeamID = '" & intTeamID & "' AND Active = 1 AND LeagueActive = 1 "
	strSQL = strSQL & " ORDER BY LeagueName ASC"
	'Response.Write strSQL
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		Do While Not(oRs.EOF)
			intTeamLinkID = oRs.Fields("lnkLeagueTeamID").Value
			strLeagueName = oRs.FieldS("LeagueName").Value
			intLeagueID = oRs.FieldS("LeagueID").Value
			intConferenceID = oRs.FieldS("LeagueConferenceID").Value
			strConferenceName = oRs.FieldS("ConferenceName").Value
			intDivisionID = oRs.FieldS("LeagueDivisionID").Value
			intLeagueTeamID = oRs.FieldS("lnkLeagueTeamID").Value
			intLeagueWins = oRs.FieldS("Wins").Value
			intLeagueLosses = oRs.FieldS("Losses").Value
			intLeagueDraws = oRs.FieldS("Draws").Value
			intLeaguePoints = oRs.FieldS("LeaguePoints").Value
			intLeagueWinPct = oRs.FieldS("WinPct").Value
			intLeagueRank = oRs.FieldS("Rank").Value
			intLeagueRoundsOne = oRs.FieldS("RoundsWon").Value
			intLeagueRoundsLost = oRs.FieldS("RoundsLost").Value
			Response.Write vbTab & vbTab & "<league name=""" & XMLEncode(Server.HTMLEncode(strLeagueName & "")) & """>" & vBCrLf
			Response.Write vbTab & vbTab & vbTab & "<leaguelink>" & vBCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<htmllink>http://www.teamwarfare.com/viewleague.asp?league=" & XMLEncode(Server.URLEncode(strLeagueName & "")) & "</htmllink>" & vBCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>n/a</xmllink>" & vBCrLf
			Response.Write vbTab & vbTab & vbTab & "</leaguelink>" & vBCrLf
			Response.Write vbTab & vbTab & vbTab & "<conference>" & XMLEncode(Server.HTMLEncode(strConferenceName & "")) & "</conference>" &  vBCrLf
			If Cint(intDivisionID) = 0 Then
				Response.Write vbTab & vbTab & vbTab & "<division>Pending division assignment</division>" &  vBCrLf
			Else
				strSQL = "SELECT DivisionName FROM tbl_league_divisions WHERE LeagueDivisionID = '" & intDivisionID & "'"
				oRs2.Open strSQL, oConn
				If Not(oRs2.EOF AND oRs2.BOF) Then
					strDivisionName = oRs2.Fields("DivisionName").Value
				End If
				oRs2.NextRecordSet
				Response.Write vbTab & vbTab & vbTab & "<division>" & XMLEncode(Server.HTMLEncode(strDivisionName & "")) & "</division>" & vBCrLf
				Response.Write vbTab & vbTab & vbTab & "<matchinformation>" & vBCrLf
				strSQL = "EXECUTE LeagueTeamMatches @LeagueTeamID = '" & intLeagueTeamID & "'"
				oRs2.Open strSQL, oConn
				If (oRs2.State = 1) Then
					If Not(oRs2.EOF AND oRs2.BOF) Then
						Do While Not (oRs2.EOF)
							If cLng(oRs2.Fields("HomeTeamLinkID").Value) = intTeamLinkID Then
								Response.Write vbTab & vbTab & vbTab & vbTab & "<match status=""Home"">" & vBCrLf
							Else
								Response.Write vbTab & vbTab & vbTab & vbTab & "<match status=""Visitor"">" & vBCrLf
							End If
							Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<opponentname>" & XMLEncode(Server.HTMLEncode(oRs2.Fields("OpponentName").Value & "")) & "</opponentname>" & vBCrLf
							Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<opponentlink>" & vBCrLf
							Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<htmllink>http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLEncode(oRs2.Fields("OpponentName").Value & "")) & "</htmllink>" & vBCrLf
							Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & XMLEncode(Server.URLEncode(oRs2.Fields("OpponentName").Value & "")) & "</xmllink>" & vBCrLf
							Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "</opponentlink>" & vBCrLf
							If Len(oRs2.Fields("Map1").Value) > 0 Then
								Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<map name=""" & XMLEncode(oRs2.Fields("Map1").Value) & """ maporder=""1"" />" & vBCrLf
							End If
							If Len(oRs2.Fields("Map2").Value) > 0 Then
								Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<map name=""" & XMLEncode(oRs2.Fields("Map2").Value) & """ maporder=""2"" />" & vBCrLf
							End If
							If Len(oRs2.Fields("Map3").Value) > 0 Then
								Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<map name=""" & XMLEncode(oRs2.Fields("Map3").Value) & """ maporder=""3"" />" & vBCrLf
							End If
							If Len(oRs2.Fields("Map4").Value) > 0 Then
								Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<map name=""" & XMLEncode(oRs2.Fields("Map4").Value) & """ maporder=""4"" />" & vBCrLf
							End If
							If Len(oRs2.Fields("Map5").Value) > 0 Then
								Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<map name=""" & XMLEncode(oRs2.Fields("Map5").Value) & """ maporder=""5"" />" & vBCrLf
							End If
							Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<matchdate>" & oRs2.Fields("MatchDate").Value & "</matchdate>" & vBCrLf
							Response.Write vbTab & vbTab & vbTab & vbTab & "</match>" & vBCrLf
							oRs2.MoveNext
						Loop
					End If
				End If
				Response.Write vbTab & vbTab & vbTab & "</matchinformation>" & vBCrLf
				oRS2.NextRecordSet
			End If

			strSQL="select PlayerHandle, lnk_league_team_player.JoinDate, lnk_league_team_player.IsAdmin "
			strSQL = strSQL & " from tbl_Players inner join lnk_league_team_player on "
			strSQL = strSQL & "lnk_league_team_player.PlayerID=tbl_players.playerid where "
			strSQL = strSQL & " lnk_league_team_player.lnkLeagueTeamID=" & intLeagueTeamID & " ORDER BY PlayerHandle"
			ors2.Open strSQL, oconn
			If Not(oRS2.EOF AND oRS2.BOF) Then
				Response.Write vbTab & vbTab & vbTab & "<roster>" & vbCrLf
				Do While Not(oRS2.EOF)
					Response.Write vbTab & vbTab & vbTab & vbTab & "<player name=""" & XMLEncode(Server.HTMLEncode(oRS2.Fields("PlayerHandle").Value)) & """>" & vbCrLf
					if len(ors2.Fields("JoinDate").Value) < 8 then
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<joindate>n/a</joindate>" & vbCrLf
					else
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<joindate>" & FormatDateTime(oRS2.Fields("JoinDate").Value, 2) & "</joindate>" & vbCrLf
					end if

					If Trim(oRS2.Fields("PlayerHandle").Value) = Trim(strFounderName) then
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<position>Team Founder</position>" & vbCrLf
					ElseIf ors2.Fields("IsAdmin").Value=1 then
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<position>Team Captain</position>" & vbCrLf
					Else
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<position>Player</position>" & vbCrLf
					End If
					Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<links>" & vBCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<htmllink>" & "http://www.teamwarfare.com/viewplayer.asp?player=" & XMLEncode(Server.URLEncode(oRS2.Fields("PlayerHandle").Value)) & "</htmllink>" & vBCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "</links>" & vBCrLf
					Response.Write vbTab & vbTab & vbTab & vbTab & "</player>" & vBCrLf
					oRS2.MoveNext 
				Loop
				Response.Write vbTab & vbTab & vbTab & "</roster>" & vbCrLf
			Else
				Response.Write vbTab & vbTab & vbTab & "<roster />" & vbCrLf
			End If
			oRs2.NextRecordSet
			Response.Write vbTab & vbTab & "</league>" & vbCrLf
			oRs.MoveNext
		Loop
	End If
	oRs.NextRecordSet
	Response.Write vbTab & "</competitioninformation>" & vbCrLf

	Response.Write "</team>" & vbCrLf
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing
For i = lBound(arrMapOptions) To uBound(arrMapOptions)
	Set arrMapOptions(i) = Nothing
Next
%>