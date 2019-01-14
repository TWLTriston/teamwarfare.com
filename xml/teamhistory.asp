<%
Option Explicit
%>
<!-- #include virtual="/include/xml.asp" -->
<%
Server.ScriptTimeout = 45
Dim oRS, strSQL, oConn, oRS2

Dim intTeamID
Dim strTeamName, strFounderName, strURL, strTeamStatus
strTeamName = Request.Form("Team")

if Len(strTeamName) = 0 then
	strTeamName = Request.QueryString("Team")
End if
If (strTeamName = "") Then
	Response.Write "You must specify a team in the querystring, in the form of: filename.asp?team=Cabal+kaBob"
	Response.End 
End If

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

Dim intLadderID
Dim strResult, strEnemyName, bDefender, i, intMaps, strStatus 
Dim map1, map1usScore, Map1ThemScore, map1OT, map1FT
Dim map2, map2usScore, Map2ThemScore, map2OT, map2FT
Dim map3, map3usScore, Map3ThemScore, map3OT, map3FT

Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf

	Response.Write "<historyinformation>" & vbCrLf
	strSQL = "SELECT t.*, p.PlayerHandle "
	strSQL = strSQL & " FROM tbl_teams t, tbl_players p "
	strSQL = strSQL & " WHERE t.TeamName = '" & Replace(strTeamName, "'", "''") & "' AND t.TeamFounderID *= p.PlayerID "
	oRS.Open strSQL, oConn
	If Not(ors.eof and ors.BOF) Then
		' Give a synopsis of the team's profile
		strFounderName = oRS.Fields("PlayerHandle").Value
		strURL = oRS.Fields("TeamURL").Value
		intTeamID = oRS.Fields("TeamID").Value
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
	
	Response.Write vbTab & "<ladderinformation>" & vbCrLf
	strSQL = "EXECUTE GetTeamHistory @TeamName = '" & Replace(strTeamName, "'", "''") & "'"
	oRS.Open strSQL, oConn
	if not (ors.eof and ors.BOF) then
		intLadderID = oRS.Fields("MatchLadderID").Value
		Response.Write vbTab & vbTab & "<ladder name=""" & XMLEncode(Server.HTMLEncode(oRs.Fields("LadderName").Value & "")) & """>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<ladderlink>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & vbTab & "<htmllink>http://www.teamwarfare.com/viewladder.asp?ladder=" & XMLEncode(Server.HTMLENcode(oRs.Fields("LadderName").Value & "")) & "</htmllink>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewladder.asp?ladder=" & XMLEncode(Server.HTMLENcode(oRs.Fields("LadderName").Value & "")) & "</xmllink>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "</ladderlink>" & vbCrLf
		do while not ors.EOF
			If intLadderID <> orS.Fields("MatchLadderID").Value Then
				intLadderID = oRS.Fields("MatchLadderID").Value
				Response.Write vbTab & vbTab & "</ladder>" & vbCrLf
				Response.Write vbTab & vbTab & "<ladder name=""" & XMLEncode(Server.HTMLEncode(oRs.Fields("LadderName").Value & "")) & """>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<ladderlink>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & vbTab & "<htmllink>http://www.teamwarfare.com/viewladder.asp?ladder=" & XMLEncode(Server.HTMLENcode(oRs.Fields("LadderName").Value & "")) & "</htmllink>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewladder.asp?ladder=" & XMLEncode(Server.HTMLENcode(oRs.Fields("LadderName").Value & "")) & "</xmllink>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "</ladderlink>" & vbCrLf
			End If					
			bDefender = False
			If oRs.Fields("WinnerName") = strTeamName Then
				strEnemyName = ors.Fields("LoserName").Value 
				strResult = "win"
				If oRS.Fields("WinnerDefending").Value Then
					bDefender = True
				End If
			Else
				strEnemyName = ors.Fields("WinnerName").Value 
				strResult = "loss"
				If oRS.Fields("WinnerDefending").Value Then
					bDefender = True
				End If
			End If
			intMaps = oRs.Fields("Maps").Value 
			
			If bDefender Then
				Response.Write vbTab & vbTab & vbTab & "<match status=""defending"">" & vbCrLf
			Else
				Response.Write vbTab & vbTab & vbTab & "<match status=""attacking"">" & vbCrLf
			End If
			
			Response.Write vbTab & vbTab & vbTab & vbTab & "<opponentname>" & XMLEncode(Server.HTMLEncode(strEnemyName & "")) & "</opponentname>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<opponentlink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<htmllink>http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLEncode(strEnemyName & "")) & "</htmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<xmllink>http://www.teamwarfare.com/xml/viewteam_v2.asp?team=" & XMLEncode(Server.URLEncode(strEnemyName & "")) & "</xmllink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "</opponentlink>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & vbTab & "<matchdate>" & Server.HTMLEncode(oRs.Fields("MatchDate").Value & "") & "</matchdate>" & vbCrLf
			If cBool(oRS.Fields("MatchForfeit").Value) Then
				Response.Write vbTab & vbTab & vbTab & vbTab & "<result>forfeit " & XMLEncode(Server.HTMLEncode(strResult & "")) & "</result>" & vbCrLf
			Else
				Response.Write vbTab & vbTab & vbTab & vbTab & "<result>" & XMLEncode(Server.HTMLEncode(strResult & "")) & "</result>" & vbCrLf
			End If

			Response.Write vbTab & vbTab & vbTab & vbTab & "<maps>" & vbCrLf
			If Not(cBool(oRS.Fields("MatchForfeit").Value)) Then 
				For i = 1 to cInt(intMaps )
					If i < 6 Then
						strStatus = ""
						If oRS.Fields("Map" & i & "OT").Value Then
							strStatus = "overtime" 
						End If
						If oRS.Fields("Map" & i & "ForFeit").Value THen
							strStatus = "forfeit" 
						End If
	
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "<map>" & vbCrLf
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<name>" & XMLEncode(Server.HTMLEncode(oRs.Fields("MatchMap" & i).Value & "")) & "</name>" & vbCrLf
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<status>" & XMLEncode(Server.HTMLEncode(strStatus & "")) & "</status>" & vbCrLf
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<score>" & vbCrLf
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<defender>" & XMLEncode(Server.HTMLEncode(oRS.Fields("MatchMap" & i & "DefenderScore").Value & "")) & "</defender>" & vbCrLf
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<attacker>" & XMLEncode(Server.HTMLEncode(oRS.Fields("MatchMap" & i & "AttackerScore").Value & "")) & "</attacker>" & vbCrLf
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "</score>" & vbCrLf
						Response.Write vbTab & vbTab & vbTab & vbTab & vbTab & "</map>" & vbCrLf
					End If
				Next
			End If
			Response.Write vbTab & vbTab & vbTab & vbTab & "</maps>" & vbCrLf
			Response.Write vbTab & vbTab & vbTab & "</match>" & vbCrLf
			oRs.MoveNext
		loop
		Response.Write vbTab & vbTab & "</ladder>" & vbCrLf
	end if
	ors.Close 
	Response.Write vbTab & "</ladderinformation>" & vbCrLf
	Response.Write "</historyinformation>" & vbCrLf

oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing
%>