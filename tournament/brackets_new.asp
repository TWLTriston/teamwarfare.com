<%

CONST ROWS_BETWEEN_TEAMS = 3
Dim iCurrentRound, iTeamPosition
Dim intTeamsPerDivision, intDivisions
Dim intRows, intRounds, iRound, iPowerOf2, iNextPowerOf2, iLastPowerOf2
Dim intRoundForLabel, iDivision, blnOnProduction
Dim intDivisionID

intDivisionID = Request.QueryString("div")
If Len(intDivisionID) = 0 Then
	intDivisionID = 1
ElseIf Not(IsNumeric(intDivisionID)) Then
	intDivisionID = 1
End If

blnOnProduction = True

strSQL = "SELECT TournamentID,  TeamsPerDiv, Divisions, Signup FROM tbl_tournaments WHERE TournamentName = '" & Replace(strTournamentName, "'", "''") & "'"
oRS.Open strSQL, oConn, 1, 1
If Not(oRS.EOF AND oRS.BOF) Then
	intTournamentID = oRS.Fields("TournamentID").Value
	intTeamsPerDivision = oRS.Fields("TeamsPerDiv").Value
	intDivisions = oRS.Fields("Divisions").Value
	blnSignUp = oRs.Fields("SignUp").Value
End If
oRS.NextRecordset
%>
	<style>
	<!--
	.t { background-color: <%=bgcone%>; color: #ffffff; font-size: 11px; font-family: Verdana; padding: 4px; border: 1px solid #444444; }
	.win { background-color: <%=bgcone%>; font-weight: bold; color: #ffffff; font-size: 11px; font-family: Verdana; padding: 4px; border: 1px solid #ffffff; }
	.d { PADDING: 2px; FONT-SIZE: 9px; FONT-WEIGHT: 600; }
	.b { FONT-SIZE: 9px; text-align: right; BORDER-RIGHT: 1px SOLID #444444; }
	.r { FONT-SIZE: 4px; BORDER-RIGHT: 1px SOLID #444444; }
	.rh { color: #ffcf3f; font-size: 11px; font-weight: bold; font-family: Verdana; text-align: center; padding: 4px; }
	.l { FONT-SIZE: 4px; BORDER-LEFT: 1px SOLID #444444; }
	.e { FONT-SIZE: 12px; }
	.c { COLOR: #FF0000; }
	.w { FONT-SIZE: 16px; font-weight: bold; text-align: center;}
	A, A:hover, A:link, A:active, A:visited { COLOR: #FFD142; FONT-FAMILY: Verdana; }
	//-->
	</style>
<table border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#444444">
<tr>
	<td>
		<table border="0" cellspacing="1" cellpadding="4" width="100%">
		<tr><th colspan="4" bgcolor="#000000">Divisions</th></tr>
		<tr>
		<% 
		Dim strDivisionName, intDivisionCount
		strSQL = "select DivisionName, DivisionID from tbl_tdivisions where TournamentID='" & inttournamentID & "' order by DivisionName ASC"
		oRs.Open strSQL, oConn
		if not(oRs.eof and oRs.bof) then
			intDivisionCount = 0
			do while not(oRs.eof)
				if intDivisionCount mod 4 = 0 AND intDivisionCount > 0 then
					response.write "</tr><tr>"
				end if
				intDivisionCount = intDivisionCount + 1
				If cint(oRs.fields("DivisionID").value) = cint(intDivisionID) Then 
					strdivisionname = oRs.fields("DivisionName").value%>
					<td align="center" bgcolor="<%=bgcone%>"><%= server.htmlencode(strdivisionname) %></td>
				<% Else %>
					<td align="center" bgcolor="<%=bgctwo%>"><a href="default2.asp?tournament=<%=Server.URLEncode(strTournamentName & "")%>&page=brackets&div=<%=oRs.fields("DivisionID").value%>"><%=server.htmlencode(oRs.fields("DivisionName").value)%></a></td>
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
Dim arrRounds(16, 256)
Dim arrTeamNames(16, 256)
Dim arrTeamTags(16, 256)
Dim arrWinner(16, 256)

Dim arrBracketBlurb(16, 15)
Dim i, j, iSeed
For i = 0 to 8
	For j = 0 to 256
		arrRounds(i,j) = i & ", " & j
		arrTeamNames(i,j) = Null
		arrTeamTags(i,j) = Null
	Next
Next

intRoundForLabel = Log(intTeamsPerDivision) / Log(2) + 1
'' Figure out how many rows and rounds we are expecting to display
intRows = (intTeamsPerDivision) + ROWS_BETWEEN_TEAMS * intTeamsPerDivision - 1
intRounds = Log(intTeamsPerDivision) / Log(2) + 1' + Log(intDivisions / 2) / Log(2) + 2

Dim strDivArray(50)
Dim intDivCounter, strClass
intDivCounter = 1

strSQL = "SELECT DivisionName FROM tbl_tdivisions WHERE TournamentID = '" & intTournamentID & "' ORDER BY TournamentID ASC"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	Do While Not (oRs.EOF)
		strDivArray(intDivCounter) = oRs.Fields("DivisionName").Value
		intDivCounter = intDivCounter + 1
		oRs.MoveNext	
	Loop
End If
oRs.NextRecordSet

If blnOnProduction Then
	For iRound = 1 to intRounds
		strSQL = "EXECUTE GetTournamentDivisionArray " & intTournamentID & ", " & intDivisionID & ", " & iRound
		'Response.Write strSQL
		'Response.End
		oRS.Open strSQl, oConn
		If oRS.State = 1 Then
			If Not(oRS.EOF AND oRS.BOF) Then
				Do While Not(oRS.EOF)
					on error resume next
					arrWinner(iRound, oRS.Fields("ArrayNumber").Value) = (oRS.Fields("TMLinkID").Value = oRS.Fields("WinnerID").Value)
					arrTeamNames(iRound, oRS.Fields("ArrayNumber").Value) = oRS.Fields("TeamName").Value
					arrTeamTags(iRound, oRS.Fields("ArrayNumber").Value) = oRS.Fields("TeamTag").Value
					arrBracketBlurb(iRound, oRs.Fields("SeedOrder").Value) = oRs.Fields("BracketBlurb").Value
					if err <> 0 then
						response.write iRound & " --" & oRS.Fields("ArrayNumber").Value
						response.end
					end if
					oRS.MoveNext
					On Error Goto 0
				Loop
			End If
			oRS.NextRecordset
		End If
	Next
End If
'' Ahh, finally the good stuff
''' This is where all the power is (and unfortunately all the processor time
''' This thing causes major processor spikes to do this looping math
''' Remember, you are dealing with a matrix of rounds * teams * brackets (3 + (teams * brackets))
''' In short, a 4 team 1 division tournament consists of an 18X3 matrix (it's big)

'' Start the table
Response.Write "<table cellspacing=""0"" align=""center"" cellpadding=""0"" border=""0"">" & vbCrLf 

Response.Write "<tr>" & vbCrLf 

'' Give the table a header.. in this case, rounds (may change the placement of the round header later
'' but for now it works well

For iRound = 1 to intRounds
	Response.Write vbTab & "<td width=""120"">&nbsp;</td>" & vbCrLf
	' Else
' 		Response.Write vbTab & "<th nowrap=""nowrap"">Round " & iCurrentRound & " </th>" & vbCrLf 	
' 	End If
 NextResponse.Write "</tr>" & vbCrLf 
'' Start going row by row 
For iTeamPosition = 0 to intRows + 1 Step 2
	Response.Write "<tr>" & vbCrLf
	'' Then go column by column
	For iRound = 1 to intRounds
		iCurrentRound = iRound ' - Abs(intRounds - iRound)
		' This system uses power's of 2 to figure out how to display a bracket
		' There may be another faster way, but this is the pattern i discovered in my testing
		'' do tell if there is another way
		iPowerOf2 = 2 ^ iCurrentRound
		iNextPowerOf2 = 2 ^ (iCurrentRound + 1)
		'' First check to see if this is a "seeded" table cell. If so, give it some color, 
		'' at a later date we can reference an array in this slot to plop names into these colored boxes
		If (iTeamPosition MOD (iPowerOf2)) = 0 AND (iTeamPosition / iPowerOf2) MOD 2 = 1 Then
			iSeed = ((iTeamPosition / iPowerOf2) + 1) / 2
			If iCurrentRound = intRounds - 1 Then
				iSeed = 0
			End If
			If IsNull(arrTeamNames(iCurrentRound, iSeed)) Then
				' No team name, therefore it's either open / tbd / bye
				If iCurrentRound = 1 Then
					If blnSignUp Then
						Response.Write vbTab & "<td nowrap=""nowrap"" class=""t""><b>Open</b></td>" & vbCrLf 
					Else
							Response.Write vbTab & "<td nowrap=""nowrap"" class=""t"">Bye</td>" & vbCrLf 
					End If
				Else
					Response.Write vbTab & "<td nowrap=""nowrap"" class=""t"">TBD</td>" & vbCrLf 
				End If
			Else
				' Team Name Exists
				If arrWinner(iCurrentRound, iSeed) Then 
					strClass = "win"
				Else
					strClass = "t"
				End If
				
				If Len(arrTeamNames(iCurrentRound, iSeed)) < 16 Then
					Response.Write vbTab & "<td nowrap=""nowrap"" class=""" & strClass  & """><a href=""http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode(arrTeamNames(iCurrentRound, iSeed)) & """>" & Server.HTMLEncode(arrTeamNames(iCurrentRound, iSeed)) & "</a></td>" & vbCrLf 
				Else
					Response.Write vbTab & "<td nowrap=""nowrap"" class=""" & strClass  & """><a href=""http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode(arrTeamNames(iCurrentRound, iSeed)) & """>" & Server.HTMLEncode(arrTeamTags(iCurrentRound, iSeed)) & "</a></td>" & vbCrLf 
				End If
			End If
		ElseIf ((iTeamPosition + 2) MOD (iPowerOf2)) = 0 AND ((iTeamPosition + 2) / iPowerOf2) MOD 2 = 1 AND Log(iTeamPosition + 2) / Log(2) = Int(Log(iTeamPosition + 2) / Log(2)) Then
			If iCurrentRound < intRounds Then
				Response.Write "<td nowrap=""nowrap"" class=""rh"">Round " & iCurrentRound & "</td>"
			Else 
				Response.Write "<td nowrap=""nowrap"" class=""rh"">Division Winner</td>"
			End If
		Else
			'' If this isnt a seeded table cell, see if the cell isnear the next cell that will be seeded, 
			'' and give it a border so the lines can be tracked
			If iCurrentRound < intRounds Then ' Exception for additional round
				If (iTeamPosition / iNextPowerOf2 ) MOD 2 = 1 And iCurrentRound <> intRounds Then
					If (iTeamPosition MOD (iNextPowerOf2)) = 0 Then
						Response.Write vbTab & "<td nowrap=""nowrap"" class=""b"">" & arrBracketBlurb(iCurrentRound, ((iTeamPosition / ((2 ^ iCurrentRound)) + 2) / 4) - 1) & "&nbsp;</td>" & vbCrLf 
					Else 
						Response.Write vbTab & "<td nowrap=""nowrap"" class=""r"">&nbsp;</td>" & vbCrLf 
					End If
				Else
					'' Other wise, this is an empty cell, just put the equivilance to nothing in here.
					Response.Write vbTab & "<td nowrap=""nowrap"" class=""e"">&nbsp;</td>"& vbCrLf 
				End If
			ElseIf iCurrentRound < intRounds Then
				'' Other wise, this is an empty cell, just put the equivilance to nothing in here.
				Response.Write vbTab & "<td nowrap=""nowrap"" class=""e"">&nbsp;</td>"& vbCrLf 
			ElseIf iCurrentRound = intRounds Then
				iLastPowerOf2 = 2 ^ (iCurrentRound - 1)
				iSeed = 1
				Response.Write vbTab & "<td nowrap=""nowrap"" class=""e"">&nbsp;</td>"& vbCrLf 
			End If
		End If
	'' Go for a round trip
	Next
	Response.Write "</tr>"& vbCrLf 
'' Next row...
Next

'' All done, wasn't that fun, I thought so !
Response.Write "</table>" & vbCrLf 
%>
