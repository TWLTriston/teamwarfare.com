<%
Dim divID, iDivs
If Request.querystring("div") <> "" Then
	If Not(IsNumeric(request.querystring("div"))) Then
		divId = 1
	Else
		divId = Request.querystring("div")
	End if
Else
	divId = 1
End If

iDivs = 0
%>
<table border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#444444">
<tr>
	<td>
		<table border="0" cellspacing="1" cellpadding="4" width="100%">
		<tr><th colspan="4" bgcolor="#000000">Divisions</th></tr>
		<tr>
		<% 
		Dim DivisionName
		strSQL = "select DivisionName, DivisionID from tbl_tdivisions where TournamentID='" & inttournamentID & "' order by DivisionName ASC"
		oRs.Open strSQL, oConn
		if not(oRs.eof and oRs.bof) then
			do while not(oRs.eof)
				if idivs mod 4 = 0 AND iDivs > 0 then
					response.write "</tr><tr>"
				end if
				iDivs = iDivs + 1
				If cint(oRs.fields("DivisionID").value) = cint(divId) Then 
					divisionname = oRs.fields("DivisionName").value%>
					<td align="center" bgcolor="<%=bgcone%>"><%= server.htmlencode(divisionname) %></td>
				<% Else %>
					<td align="center" bgcolor="<%=bgctwo%>"><a href="default.asp?tournament=<%=Server.URLEncode(strTournamentName & "")%>&page=brackets&div=<%=oRs.fields("DivisionID").value%>"><%=server.htmlencode(oRs.fields("DivisionName").value)%></a></td>
				<% 
				End If 
				oRs.movenext
			Loop 
		end If
		oRs.NextRecordSet
		%>
		</tr>
		</table>
	</td>
</tr>
</table>
<br />
<br />

<%
Dim TeamsPerDiv, LoopNumber, LoopNumber1, LoopNumber2
DIM i, y
If cStr(divID) = "0" Then
	intTeamsPerDiv = iDivs
	response.write " Hey hey"
	If intTeamsPerDiv Mod 2 <> 0 Then
		intTeamsPerDiv = intTeamsPerDiv + 1
	End If
End If
'Determine how many vars to declare in each array
DIM roundOneID, roundTwoID, roundThreeID, roundFourID, roundFiveID
DIM roundOneNameString, roundTwoNameString, roundThreeNameString, roundFourNameString, roundFiveNameString
Dim roundSixId, roundSixNameString

TeamsPerDiv = cint(intTeamsPerDiv)

If TeamsPerDiv = 8 Then
	REDIM roundOneId(7)
	REDIM roundOneNameString(7)

	REDIM roundTwoId(3)
	REDIM roundTwoNameString(3)

	REDIM roundThreeId(1)
	REDIM roundThreeNameString(1)
ElseIf TeamsPerDiv = 4 Then
	REDIM roundOneId(3)
	REDIM roundOneNameString(3)

	REDIM roundTwoId(1)
	REDIM roundTwoNameString(1)

Elseif TeamsPerDiv = 16 Then
	REDIM roundOneId(15)
	REDIM roundOneNameString(15)

	REDIM roundTwoId(7)
	REDIM roundTwoNameString(7)

	REDIM roundThreeId(3)
	REDIM roundThreeNameString(3)

	REDIM roundFourId(1)
	REDIM roundFourNameString(1)
Elseif TeamsPerDiv = 32 Then
	REDIM roundOneId(31)
	REDIM roundOneNameString(31)

	REDIM roundTwoId(15)
	REDIM roundTwoNameString(15)

	REDIM roundThreeId(7)
	REDIM roundThreeNameString(7)

	REDIM roundFourId(3)
	REDIM roundFourNameString(3)

	REDIM roundFiveId(1)
	REDIM roundFiveNameString(1)
End If

'Round 1 Stuff

strSQL = "select Team1ID, Team2ID, WinnerID from tbl_rounds where "
strSQL = strSQL & "DivisionID='" & divid & "' and TournamentID='" & intTournamentID & "' and ROUND='1' Order by SeedOrder"
oRS2.open strSQL, oConn

if not(oRS2.eof and oRS2.bof) then
i = cint(0)
	Do While Not (oRS2.EOF)
		roundOneId(i) = oRS2("team1id")
		i = i + 1
		roundOneId(i) = oRS2("team2id")
		i = i + 1
		oRS2.movenext
	Loop
end if
oRS2.nextrecordset

If TeamsPerDiv = 4 Then
	LoopNumber = 3
ElseIf TeamsPerDiv = 8 Then
	LoopNumber = 7
Elseif TeamsPerDiv = 16 Then
	LoopNumber = 15
Elseif TeamsPerDiv = 32 Then
	LoopNumber = 31
End if

For y = 0 to LoopNumber
	If roundOneId(y) = "" Or roundOneId(y) = 0 Then
		roundOneId(y) = 0
		If blnSignUp Then
			roundOneNameString(y) = "Open"
		Else
			roundOneNameString(y) = "Bye"
		End If
	Else
		strSQL = "SELECT TeamName, TeamTag FROM tbl_teams inner join lnk_t_M on lnk_t_m.teamid = tbl_teams.teamid WHERE TMLinkID='" & roundOneId(y) & "'"
		oRS2.open strSQL, oConn
		if not(oRS2.eof and oRS2.bof) then
			if len(server.htmlencode(oRS2("teamname"))) < 16 then
				roundOneNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamname")) & "</a>" & vbcrlf
			elseIf Len(oRS2("TeamTag")) > 0 Then
				roundOneNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamtag")) & "</a>" & vbcrlf
			Else
				roundOneNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(Left(oRS2("teamname"), 16)) & "</a>" & vbcrlf
			end if
		end if
		oRS2.nextrecordset
	End If
Next

'Round 2 Stuff

If TeamsPerDiv = 4 Then
	LoopNumber1 = 0
	LoopNumber2 = 1
ElseIf TeamsPerDiv = 8 Then
	LoopNumber1 = 1
	LoopNumber2 = 3
Elseif TeamsPerDiv = 16 Then
	LoopNumber1 = 3
	LoopNumber2 = 7
Elseif TeamsPerDiv = 32 Then
	LoopNumber1 = 7
	LoopNumber2 = 15
End if

i = cint(0)
For y = 0 to LoopNumber1
	strSQL = "select Team1ID, Team2ID, SeedOrder from tbl_rounds WHERE " &_
			"DivisionID='" & divid & "' and TournamentID='" & intTournamentID & "' and ROUND='2' and SeedOrder='" & y & "'"
	oRS2.open strSQL, oConn

	If not(oRS2.EOF and oRS2.bof) Then
		roundTwoId(i) = oRS2("Team1ID")
		i = i + 1
		roundTwoId(i) = oRS2("Team2ID")
		i = i + 1
	Else
		roundTwoId(i) = 0
		i = i + 1
		roundTwoId(i) = 0
		i = i + 1
	End If
	oRS2.nextrecordset
Next

For y = 0 to LoopNumber2
	If roundTwoId(y) = "" Or roundTwoId(y) = 0 Then
		roundTwoId(y) = 0
		roundTwoNameString(y) = "TBD"
	Else
		strSQL = "SELECT TeamName, TeamTag FROM tbl_teams inner join lnk_t_M on lnk_t_m.teamid = tbl_teams.teamid WHERE TMLinkID='" & roundtwoId(y) & "'"
		oRS2.open strSQL, oConn
		if not(oRS2.eof and oRS2.bof) then
			if len(server.htmlencode(oRS2("teamname"))) < 16 then
				roundTwoNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamname")) & "</a>" & vbcrlf
			elseIf Len(oRS2("TeamTag")) > 0 Then
				roundTwoNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamtag")) & "</a>" & vbcrlf
			Else
				roundTwoNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(Left(oRS2("teamname"), 16)) & "</a>" & vbcrlf
			end if
		end if
		oRS2.nextrecordset
	End If
Next

'Round 3 Stuff

If TeamsPerDiv = 4 Then
	LoopNumber1 = false
	LoopNumber2 = false
ElseIf TeamsPerDiv = 8 Then
	LoopNumber1 = false
	LoopNumber2 = 1
Elseif TeamsPerDiv = 16 Then
	LoopNumber1 = 1
	LoopNumber2 = 3
Elseif TeamsPerDiv = 32 Then
	LoopNumber1 = 3
	LoopNumber2 = 7
End ifIf TeamsPerDiv = 4 Then
	strSQL = "select WinnerID from tbl_rounds WHERE "
	strSQL = strSQL & "DivisionID='" & divid & "' and TournamentID='" & intTournamentID  & "' and ROUND='2'"
	oRS2.open strSQL, oConn

	If Not(oRS2.eof and oRS2.bof) Then
		If Not oRS2.EOF Then
			roundThreeId = oRS2("WinnerID")
		Else
			roundThreeId = 0
		End If
	end if
	oRS2.nextrecordset

	If roundThreeId = "" Or roundThreeId = 0 Then
		roundThreeId = 0
		roundThreeNameString = "TBD"
	Else
		strSQL = "SELECT TeamName, TeamTag FROM tbl_teams inner join lnk_t_M on lnk_t_m.teamid = tbl_teams.teamid WHERE TMLinkID='" & roundThreeId & "'"
		oRS2.open strSQL, oConn
		if not(oRS2.eof and oRS2.bof) then
			if len(server.htmlencode(oRS2("teamname"))) < 16 then
				roundThreeNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamname")) & "</a>" & vbcrlf
			elseIf Len(oRS2("TeamTag")) > 0 Then
				roundThreeNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamtag")) & "</a>" & vbcrlf
			Else
				roundThreeNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(Left(oRS2("teamname"), 16)) & "</a>" & vbcrlf
			end if
		end if
		oRS2.nextrecordset
	End If
ElseIf LoopNumber1 <> False Then
	i = cint(0)
	For y = 0 to LoopNumber1
		strSQL = "select Team1ID, Team2ID, SeedOrder from tbl_rounds WHERE " &_
				"DivisionID='" & divid & "' and TournamentID='" & intTournamentID & "' and ROUND='3' and SeedOrder='" & y & "'"
		oRS2.open strSQL, oConn

		If not(oRS2.EOF and oRS2.bof) Then
			roundThreeId(i) = oRS2("Team1ID")
			i = i + 1
			roundThreeId(i) = oRS2("Team2ID")
			i = i + 1
		Else
			roundThreeId(i) = 0
			i = i + 1
			roundThreeId(i) = 0
			i = i + 1
		End If
		oRS2.nextrecordset
	Next
Else
	strSQL = "select Team1ID, Team2ID, WinnerID from tbl_rounds where "
	strSQL = strSQL & "DivisionID='" & divid & "' and TournamentID='" & intTournamentID & "' and ROUND='3' Order by SeedOrder"
	oRS2.open strSQL, oConn

	if not(oRS2.eof and oRS2.bof) then
		If not oRS2.EOF Then
			roundThreeId(0) = oRS2("Team1ID")
			roundThreeId(1) = oRS2("Team2ID")
		Else
			roundThreeId(0) = 0
			roundThreeId(1) = 0
		End If
	end if
	oRS2.nextrecordset
End If

If LoopNumber2 <> False Then
	For y = 0 to LoopNumber2
		If roundThreeId(y) = "" Or roundThreeId(y) = 0 Then
			roundThreeId(y) = 0
			roundThreeNameString(y) = "TBD"
		Else
			strSQL = "SELECT TeamName, TeamTag FROM tbl_teams inner join lnk_t_M on lnk_t_m.teamid = tbl_teams.teamid WHERE TMLinkID='" & roundthreeId(y) & "'"
			oRS2.open strSQL, oConn
			if not(oRS2.eof and oRS2.bof) then
				if len(server.htmlencode(oRS2("teamname"))) < 16 then
					roundThreeNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamname")) & "</a>" & vbcrlf
				elseIf Len(oRS2("TeamTag")) > 0 Then
					roundThreeNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamtag")) & "</a>" & vbcrlf
				Else
					roundThreeNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(Left(oRS2("teamname"), 16)) & "</a>" & vbcrlf
				end if
			end if
			oRS2.nextrecordset
		End If
	Next
End if
'Round 4 Stuff

If TeamsPerDiv = 8 Then
	strSQL = "select WinnerID from tbl_rounds WHERE "
	strSQL = strSQL & "DivisionID='" & divid & "' and TournamentID='" & intTournamentID & "' and ROUND='3'"
	oRS2.open strSQL, oConn

	If Not(oRS2.eof and oRS2.bof) Then
		If Not oRS2.EOF Then
			roundFourId = oRS2("WinnerID")
		Else
			roundFourId = 0
		End If
	end if
	oRS2.nextrecordset

	If roundFourId = "" Or roundFourId = 0 Then
		roundFourId = 0
		roundFourNameString = "TBD"
	Else
		strSQL = "SELECT TeamName, TeamTag FROM tbl_teams inner join lnk_t_M on lnk_t_m.teamid = tbl_teams.teamid WHERE TMLinkID='" & RoundFourID & "'"
		oRS2.open strSQL, oConn
		if not(oRS2.eof and oRS2.bof) then
			if len(server.htmlencode(oRS2("teamname"))) < 16 then
				roundFourNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamname")) & "</a>" & vbcrlf
			elseIf Len(oRS2("TeamTag")) > 0 Then
				roundFourNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamtag")) & "</a>" & vbcrlf
			Else
				roundFourNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(Left(oRS2("teamname"), 16)) & "</a>" & vbcrlf
			end if
		end if
		oRS2.nextrecordset
	End If
ElseIf TeamsPerDiv > 8 Then
	If TeamsPerDiv = 16 Then
		LoopNumber1 = false
		LoopNumber2 = 1
	Elseif TeamsPerDiv = 32 Then
		LoopNumber1 = 1
		LoopNumber2 = 3
	End if
	
	If LoopNumber1 <> false Then
		i = 0
		For y = 0 to LoopNumber1
			strSQL = "select Team1ID, Team2ID, SeedOrder from tbl_rounds WHERE " &_
					"DivisionID='" & divid & "' and TournamentID='" & intTournamentID & "' and ROUND='4' and SeedOrder='" & y & "'"
			oRS2.open strSQL, oConn

			If not(oRS2.EOF and oRS2.bof) Then
				roundFourId(i) = oRS2("Team1ID")
				i = i + 1
				roundFourId(i) = oRS2("Team2ID")
				i = i + 1
			Else
				roundFourId(i) = 0
				i = i + 1
				roundFourId(i) = 0
				i = i + 1
			End If
			oRS2.nextrecordset
		Next
	Else
		strSQL = "select Team1ID, Team2ID, WinnerID from tbl_rounds where "
		strSQL = strSQL & "DivisionID='" & divid & "' and TournamentID='" & intTournamentID & "' and ROUND='4' Order by SeedOrder"
		oRS2.open strSQL, oConn

		if not(oRS2.eof and oRS2.bof) then
			If not oRS2.EOF Then
				roundFourId(0) = oRS2("Team1ID")
				roundFourId(1) = oRS2("Team2ID")
			Else
				roundFourId(0) = 0
				roundFourId(1) = 0
			End If
		end if
		oRS2.nextrecordset
	End If

	For y = 0 to LoopNumber2
		If roundFourId(y) = "" Or roundFourId(y) = 0 Then
			roundFourId(y) = 0
			roundFourNameString(y) = "TBD"
		Else
			strSQL = "SELECT TeamName, TeamTag FROM tbl_teams inner join lnk_t_M on lnk_t_m.teamid = tbl_teams.teamid WHERE TMLinkID='" & roundFourId(y) & "'"
			oRS2.open strSQL, oConn
			if not(oRS2.eof and oRS2.bof) then
				if len(server.htmlencode(oRS2("teamname"))) < 16 then
					roundFourNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamname")) & "</a>" & vbcrlf
				elseIf Len(oRS2("TeamTag")) > 0 Then
					roundFourNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamtag")) & "</a>" & vbcrlf
				Else
					roundFourNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(Left(oRS2("teamname"), 16)) & "</a>" & vbcrlf
				end if
			end if
			oRS2.nextrecordset
		End If
	Next
End if

'Round 5 Stuff

If TeamsPerDiv = 16 Then
	strSQL = "select WinnerID from tbl_rounds WHERE "
	strSQL = strSQL & "DivisionID='" & divid & "' and TournamentID='" & intTournamentID & "' and ROUND='4'"
	oRS2.open strSQL, oConn

	If Not(oRS2.eof and oRS2.bof) Then
		If Not oRS2.EOF Then
			roundFiveId = oRS2("WinnerID")
		Else
			roundFiveId = 0
		End If
	end if
	oRS2.nextrecordset

	If roundFiveId = "" Or roundFiveId = 0 Then
		roundFiveId = 0
		roundFiveNameString = "TBD"
	Else
		strSQL = "SELECT TeamName, TeamTag FROM tbl_teams inner join lnk_t_M on lnk_t_m.teamid = tbl_teams.teamid WHERE TMLinkID='" & roundFiveId & "'"
		oRS2.open strSQL, oConn
		if not(oRS2.eof and oRS2.bof) then
			if len(server.htmlencode(oRS2("teamname"))) < 16 then
				roundFiveNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamname")) & "</a>" & vbcrlf
			elseIf Len(oRS2("TeamTag")) > 0 Then
				roundFiveNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamtag")) & "</a>" & vbcrlf
			Else
				roundFiveNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(Left(oRS2("teamname"), 16)) & "</a>" & vbcrlf
			end if
		end if
		oRS2.nextrecordset
	End If
ElseIf TeamsPerDiv > 16 Then
	LoopNumber1 = false
	LoopNumber2 = 1

	strSQL = "select Team1ID, Team2ID, WinnerID from tbl_rounds where "
	strSQL = strSQL & "DivisionID='" & divid & "' and TournamentID='" & intTournamentID & "' and ROUND='5' Order by SeedOrder"
	oRS2.open strSQL, oConn

	if not(oRS2.eof and oRS2.bof) then
		If not oRS2.EOF Then
			roundFiveId(0) = oRS2("Team1ID")
			roundFiveId(1) = oRS2("Team2ID")
		Else
			roundFiveId(0) = 0
			roundFiveId(1) = 0
		End If
	end if
	oRS2.nextrecordset

	For y = 0 to LoopNumber2
		If roundFiveId(y) = "" Or roundFiveId(y) = 0 Then
			roundFiveId(y) = 0
			roundFiveNameString(y) = "TBD"
		Else
			strSQL = "SELECT TeamName, TeamTag FROM tbl_teams inner join lnk_t_M on lnk_t_m.teamid = tbl_teams.teamid WHERE TMLinkID='" & roundFiveId(y) & "'"
			oRS2.open strSQL, oConn
			if not(oRS2.eof and oRS2.bof) then
				if len(server.htmlencode(oRS2("teamname"))) < 16 then
					roundFiveNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamname")) & "</a>" & vbcrlf
				elseIf Len(oRS2("TeamTag")) > 0 Then
					roundFiveNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamtag")) & "</a>" & vbcrlf
				Else
					roundFiveNameString(y) = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(Left(oRS2("teamname"), 16)) & "</a>" & vbcrlf
				end if
			end if
			oRS2.nextrecordset
		End If
	Next
End if

'Round 6 Stuff

If TeamsPerDiv = 8 Then

Elseif TeamsPerDiv = 16 Then
Elseif TeamsPerDiv = 32 Then
	strSQL = "select WinnerID from tbl_rounds WHERE "
	strSQL = strSQL & "DivisionID='" & divid & "' and TournamentID='" & intTournamentID & "' and ROUND='5'"
	oRS2.open strSQL, oConn

	If Not(oRS2.eof and oRS2.bof) Then
		If Not oRS2.EOF Then
			roundSixId = oRS2("WinnerID")
		Else
			roundSixId = 0
		End If
	end if
	oRS2.nextrecordset

	If roundSixId = "" Or roundSixId = 0 Then
		roundSixId = 0
		roundSixNameString = "TBD"
	Else
		strSQL = "SELECT TeamName, TeamTag FROM tbl_teams inner join lnk_t_M on lnk_t_m.teamid = tbl_teams.teamid WHERE TMLinkID='" & roundSixId & "'"
		oRS2.open strSQL, oConn
		if not(oRS2.eof and oRS2.bof) then
			if len(server.htmlencode(oRS2("teamname"))) < 16 then
				roundSixNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamname")) & "</a>" & vbcrlf
			elseIf Len(oRS2("TeamTag")) > 0 Then
				roundSixNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(oRS2("teamtag")) & "</a>" & vbcrlf
			Else
				roundSixNameString = "<a href=""/viewTeam.asp?team=" & server.urlencode(oRS2("teamname")) & """>" & server.htmlencode(Left(oRS2("teamname"), 16)) & "</a>" & vbcrlf
			end if
		end if
		oRS2.nextrecordset
	End If
End if
%>
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 ALIGN=CENTER WIDTH=760>
<TR><TD>
<% If TeamsPerDiv = 4 Then %>
<!-- #include virtual="/tournament/brackets/bracketSkeleton-4.asp" -->
<% ElseIf TeamsPerDiv = 8 Then %>
<!-- #include virtual="/tournament/brackets/bracketSkeleton-8.asp" -->
<% Elseif TeamsPerDiv = 16 Then %>
<!-- #include virtual="/tournament/brackets/bracketSkeleton-16.asp" -->
<% Elseif TeamsPerDiv = 32 Then %>
<!-- #include virtual="/tournament/brackets/bracketSkeleton-32a.asp" -->
<% End If %>
</TD></TR></TABLE>
