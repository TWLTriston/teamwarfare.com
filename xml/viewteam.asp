<%
Option Explicit

Dim strTeamName
Dim oRS, strSQL, oConn, oRS2

Dim intMatchID, strEnemyName, strMatchDate
Dim strMap1, strMap2, strMap3, strStatus
Dim intCurrentMapNumber, strCurrentMap, strVerbiage, strEnemyVerbiage
Dim strMapArray(6)

strTeamName = Request.Form("Team")
if Len(strTeamName) = 0 then
	strTeamName = Request.QueryString("Team")
End if
If (strTeamName = "") Then
	Response.Write "You must specify a team in the form post."
	Response.End 
End If

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf
	Response.Write "<TeamInformation>" & vbCrLf
	strSQL = "SELECT * FROM vLadder WHERE TeamName = '" & Replace(strTeamName, "'", "''") & "'"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF And oRS.BOF) Then
		Do While Not(oRS.EOF)
			strStatus = oRS.Fields("Status").Value 
			Response.Write vbTab & "<LadderInformation>" & vbCrLf
			Response.Write vbTab & vbTab & "<LadderName>" & Server.HTMLEncode(oRS.Fields("LadderName").Value) & "</LadderName>" & vbCrLf
			Response.Write vbTab & vbTab & "<LadderLink>" & "http://www.teamwarfare.com/viewladder.asp?ladder=" & Server.UrlEncode(oRS.Fields("LadderName").Value) & "</LadderLink>" & vbCrLf
			Response.Write vbTab & vbTab & "<Rank>" & oRS.Fields("Rank").Value & "</Rank>" & vbCrLf
			Response.Write vbTab & vbTab & "<Wins>" & oRS.Fields("Wins").Value & "</Wins>" & vbCrLf
			Response.Write vbTab & vbTab & "<Losses>" & oRS.Fields("Losses").Value & "</Losses>" & vbCrLf
			Response.Write vbTab & vbTab & "<Forfeits>" & oRS.Fields("Forfeits").Value & "</Forfeits>" & vbCrLf
			Response.Write vbTab & vbTab & "<CurrentStatus>" & strStatus & "</CurrentStatus>" & vbCrLf
			Response.Write vbTab & vbTab & "<DetailedStatus>" & vbCrLf
			If  Left(uCase(strStatus), 6)  = "DEFEND" Then
				strSQL = "select m.MatchAttackerID, m.MatchID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID, t.teamname, t.teamtag "
				strSQL = strSQL & " from tbl_Matches m, tbl_teams t, lnk_t_l lnk "
				strSQL = strSQL & " where m.matchdefenderID = " & oRS.Fields("TLLinkID").Value  
				strSQL = strSQL & " AND t.teamid = lnk.teamid "
				strSQL = strSQL & " AND lnk.tllinkid = m.MatchAttackerID "
				oRS2.Open strSQL, oconn
				if not (oRS2.EOF and oRS2.BOF) then
					intMatchID = oRS2.Fields("MatchID").Value 
					strEnemyName = oRS2.Fields("TeamName").Value
					strMatchDate = oRS2.Fields("MatchDate").Value
					strMap1 = oRS2.fields("MatchMap1ID").value
					strMap2 = oRS2.fields("MatchMap2ID").value
					strMap3 = oRS2.fields("MatchMap3ID").value
					strMapArray(1) = oRS2.fields("MatchMap1ID").value
					strMapArray(2) = oRS2.fields("MatchMap2ID").value
					strMapArray(3) = oRS2.fields("MatchMap3ID").value
				end if
				oRS2.NextRecordset 
								
				If Right(uCase(strTeamName), 1) = "s" Then
					strVerbiage = " have "
				Else
					strVerbiage = " has "
				End IF
								
				If Right(uCase(strEnemyName), 1) = "s" Then
					strEnemyVerbiage = " have "
				Else
					strEnemyVerbiage = " has "
				End IF
												
				Response.Write vbTab & vbTab & vbTab & "<Opponent>" & Server.HTMLEncode(strEnemyName) & "</Opponent>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<OpponentLink>" & "http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode(strEnemyName) & "</OpponentLink>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<MatchDate>" & strMatchDate & "</MatchDate>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<Map1>" & strMap1 & "</Map1>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<Map2>" & strMap2 & "</Map2>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<Map3>" & strMap3 & "</Map3>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<MatchOptions>" & vbCrLf
				
				' Show Current Selected Options
				strSQL = "SELECT * FROM vMatchOptions "
				strSQL = strSQL & " WHERE MatchID = '" & intMatchID & "'" 
				'Response.Write strSQL
				oRs2.Open strSQL, oConn
				If Not(oRS2.EOF AND oRS2.BOF) Then
					strCurrentMap = ""
					intCurrentMapNumber = -1
					Do While Not(oRS2.EOF)
						If intCurrentMapNumber <> oRS2.Fields("MapNumber").Value Then
							intCurrentMapNumber = oRS2.Fields("MapNumber").Value
							strCurrentMap = strMapArray(intCurrentMapNumber)
						End If
						Response.Write vbTab & vbTab & vbTab & vbTab & "<Option>"
						Select Case(oRS2.Fields("SelectedBy").Value)
							Case "D"
								If oRS2.Fields("SideChoice").Value = "Y" Then
									Response.Write strTeamName & strVerbiage & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value)
								Else
									Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value
								End If
							Case "A"
								If oRS2.Fields("SideChoice").Value = "Y" Then
									Response.Write strTeamName & strVerbiage & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value)
								Else
									Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value
								End If
							Case "R"
								If oRS2.Fields("SideChoice").Value = "Y" Then
									Response.Write strTeamName & strVerbiage & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value)
								Else
									Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value
								End If
							Case Else
								Response.Write "ERROR"
						End Select
						oRS2.MoveNext
						Response.Write " on " & strCurrentMap & "</Option>" & vBCrLf
					Loop
				End If	
				oRS2.NextRecordset 					
				Response.Write vbTab & vbTab & vbTab & "</MatchOptions>" & vbCrLf
			ElseIf Left(uCase(strStatus), 6)  = "ATTACK" Then
				strsql = "select m.MatchDefenderID, m.MatchID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID, t.teamname, t.teamtag "
				strsql = strsql & " from tbl_Matches m, tbl_teams t, lnk_t_l lnk "
				strsql = strsql & " where m.MatchAttackerID = " & oRS.Fields("TLLinkID").Value 
				strsql = strsql & " AND t.teamid = lnk.teamid "
				strsql = strsql & " AND lnk.tllinkid = m.matchdefenderID "
				oRS2.Open strSQL, oconn
				if not (oRS2.EOF and oRS2.BOF) then
					intMatchID = oRS2.Fields("MatchID").Value 
					strEnemyName = oRS2.Fields("TeamName").Value
					strMatchDate = oRS2.Fields("MatchDate").Value
					strMap1 = oRS2.fields("MatchMap1ID").value
					strMap2 = oRS2.fields("MatchMap2ID").value
					strMap3 = oRS2.fields("MatchMap3ID").value
				end if
				oRS2.NextRecordset 

				Response.Write vbTab & vbTab & vbTab & "<Opponent>" & Server.HTMLencode(strEnemyName) & "</Opponent>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<OpponentLink>" & "http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode(strEnemyName) & "</OpponentLink>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<MatchDate>" & strMatchDate & "</MatchDate>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<Map1>" & strMap1 & "</Map1>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<Map2>" & strMap2 & "</Map2>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<Map3>" & strMap3 & "</Map3>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<MatchOptions>" & vbCrLf
				

				If Right(uCase(strTeamName), 1) = "s" Then
					strVerbiage = " have "
				Else
					strVerbiage = " has "
				End IF
								
				If Right(uCase(strEnemyName), 1) = "s" Then
					strEnemyVerbiage = " have "
				Else
					strEnemyVerbiage = " has "
				End IF
																
				' Show Current Selected Options
				strSQL = "SELECT * FROM vMatchOptions "
				strSQL = strSQL & " WHERE MatchID = '" & intMatchID & "'" 
				'Response.Write strSQL
				oRs2.Open strSQL, oConn
				If Not(oRS2.EOF AND oRS2.BOF) Then
					strCurrentMap = ""
					intCurrentMapNumber = -1
					Do While Not(oRS2.EOF)
						If intCurrentMapNumber <> oRS2.Fields("MapNumber").Value Then
							intCurrentMapNumber = oRS2.Fields("MapNumber").Value
							strCurrentMap = strMapArray(intCurrentMapNumber)
						End If
						Response.Write vbTab & vbTab & vbTab & vbTab & "<Option>"
						Select Case(oRS2.Fields("SelectedBy").Value)
							Case "A"
								If oRS2.Fields("SideChoice").Value = "Y" Then
									Response.Write strTeamName & strVerbiage & oRS2.Fields("ValueName").Value & " for " & lCase(oRS2.Fields("OptionName").Value)
								Else
									Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value
								End If
							Case "D"
								If oRS2.Fields("SideChoice").Value = "Y" Then
									Response.Write strTeamName & strVerbiage & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value)
								Else
									Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value
								End If
							Case "R"
								If oRS2.Fields("SideChoice").Value = "Y" Then
									Response.Write strTeamName & strVerbiage & oRS2.Fields("Opposite").Value & " for " & lCase(oRS2.Fields("OptionName").Value)
								Else
									Response.Write oRS2.Fields("OptionName").Value & ": " & oRS2.Fields("ValueName").Value
								End If
							Case Else
								Response.Write "Error"
						End Select
						oRS2.MoveNext
						Response.Write " on " & strCurrentMap & "</Option>" & vBCrLf
					Loop
				Else
					Response.Write "<Option>No Options Found</Option>"
				End If
				oRS2.NextRecordset 					
				Response.Write vbTab & vbTab & vbTab & "</MatchOptions>" & vbCrLf
 			Else
				Response.Write vbTab & vbTab & vbTab & "<NoStatus>No Detailed Status Available</NoStatus>" & vbCrLf
			
			End If

			Response.Write vbTab & vbTab & "</DetailedStatus>" & vbCrLf
			Response.Write vbTab & "</LadderInformation>" & vbCrLf
			oRS.MoveNext
		Loop
	End If
	oRS.Close	

	Response.Write "</TeamInformation>" & vbCrLf
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing

%>