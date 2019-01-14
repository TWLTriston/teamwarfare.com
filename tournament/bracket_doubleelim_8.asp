<%
intDivisionID = Request.QueryString("div")
If Len(intDivisionID) = 0 Then
	intDivisionID = 1
ElseIf Not(IsNumeric(intDivisionID)) Then
	intDivisionID = 1
End If

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

		<table border="0" cellspacing="0" cellpadding="0" class="cssBordered">
		<tr><th colspan="4" bgcolor="#000000">Divisions</th></tr>
		<tr>
		<% 
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
					<td align="center" bgcolor="<%=bgctwo%>"><a href="default.asp?tournament=<%=Server.URLEncode(strTournamentName & "")%>&page=brackets&div=<%=oRs.fields("DivisionID").value%>"><%=server.htmlencode(oRs.fields("DivisionName").value)%></a></td>
				<% 
				End If 
				oRs.movenext
			Loop 
		end If
		oRs.NextRecordSet
		%>
		</tr>
		
		<tr>
			<% If cint(intDivisionID) = 0 Then %>
			<td colspan="4" align="center" bgcolor="#000000"><b>Tournament Finals</b></td>
			<% Else %>
			<td colspan="4" align="center" bgcolor="#000000"><a href="default.asp?tournament=<%=Server.URLEncode(strTournamentName & "")%>&page=brackets&div=0">Tournament Finals</a></td>
			<% End If %>
		</tr>
		</table>
	<br /><Br />

		<style>
	<!--
	.t { background-color: <%=bgcone%>; color: #ffffff; font-size: 11px; font-family: Verdana; padding: 4px; border: 1px solid #444444; }
	.tbd { background-color: #000000; color: #888888; font-size: 11px; font-family: Verdana; padding: 4px; border: 1px solid #444444; }
	.win { background-color: #111111; font-weight: bold; color: #ffffff; font-size: 11px; font-family: Verdana; padding: 3px; border: 2px ridge #ffffff; }
	.lose { background-color: #000000; color: #888888; font-size: 11px; font-family: Verdana; padding: 4px; border: 1px solid #444444; }
	.d { PADDING: 2px; FONT-SIZE: 9px; FONT-WEIGHT: 600; }
	.b { color: #ffffff; font-family: verdana; FONT-SIZE: 9px; text-align: right; BORDER-RIGHT: 1px SOLID #444444; }
	.bd { color: #ffffff; font-family: verdana; FONT-SIZE: 9px; text-align: right; BORDER-RIGHT: 1px dotted #444444; }
	.r { FONT-SIZE: 4px; BORDER-RIGHT: 1px SOLID #444444; }
	.rd { FONT-SIZE: 4px; BORDER-RIGHT: 1px dotted #444444; }
	.rh { color: #FFD142; font-size: 11px; font-weight: bold; font-family: Verdana; text-align: center; padding: 4px; }
	.rhr { color: #FFD142; font-size: 11px; font-weight: bold; font-family: Verdana; text-align: center; padding: 4px; border-right: 1px solid #444444; }
	.l { FONT-SIZE: 4px; BORDER-LEFT: 1px SOLID #444444; }
	.e { FONT-SIZE: 4px; }
	.c { COLOR: #FF0000; }
	.fw { background-color: #222222; color: #ffffff; font-size: 11px; font-family: Verdana; padding: 4px; border: 1px solid #444444; }
	.fwd { background-color: #222222; color: #ffffff; font-size: 11px; font-family: Verdana; padding: 4px; border: 1px dotted #444444; }
	.chmp { background-color: #222222; color: #ffffff; font-size: 11px; font-family: Verdana; padding: 4px; border: 1px dotted #444444; }
	.fl { background-color: <%=bgcone%>; color: #ffffff; font-size: 11px; font-family: Verdana; padding: 4px; border: 1px solid #444444; }
	.wb { background-color: <%=bgctwo%>; color: #ffffff; font-size: 11px; font-family: Verdana; padding: 4px; border: 1px solid #444444; }
	.w { FONT-SIZE: 16px; font-weight: bold; text-align: center;}
	//-->
	</style>	
<%
Dim arrDBLTeamNames (7, 16) ' Round, Seed Num
Dim arrDBLTeamTags (7, 16)
Dim arrDBLBracketBlurbs (7, 8)
strSQL = "EXECUTE GetTournamentDBLDivisionArray @TournamentID = '" & intTournamentID & "', @DivisionID = '" & intDivisionID & "'"
'Response.Write strSQL

oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	Do While Not(oRs.EOF)
		arrDBLTeamNames(oRs.Fields("Round").Value, oRs.Fields("SeedOrder").Value * 2 + oRs.Fields("Modifier").Value) = oRs.Fields("tmpTeamName").Value
		arrDBLTeamTags(oRs.Fields("Round").Value, oRs.Fields("SeedOrder").Value * 2 + oRs.Fields("Modifier").Value) = oRs.Fields("tmpTeamTag").Value
		If (oRs.Fields("Round").Value = 4 OR oRs.Fields("Round").Value = 5) AND oRs.Fields("SeedOrder").Value = 0 Then
			If oRs.Fields("tmpTMLinkID").Value = oRs.Fields("WinnerID").Value Then
				arrDBLTeamNames(6, 1) = oRs.Fields("tmpTeamName").Value
				arrDBLTeamTags(6, 1) = oRs.Fields("tmpTeamTag").Value
			End If
			If oRs.Fields("Round").Value = 4 AND oRs.Fields("WinnerID").Value <> 0 Then
				arrDBLTeamNames(5, 1) = "-"
				arrDBLTeamNames(5, 2) = "-"
				arrDBLTeamTags(5, 1) = "-"
				arrDBLTeamTags(5, 2) = "-"
			End If
		End If
		If (oRs.Fields("Round").Value = 5) AND oRs.Fields("WinnerID").Value = 0 Then
				arrDBLTeamNames(6, 1) = ""
				arrDBLTeamTags(6, 1) = ""
		End If
		oRs.MoveNext
	Loop
End If

Set oRs = oRs.NextRecordSet
If Not(oRs.EOF AND oRs.BOF) Then
	Do While Not(oRs.EOF)
		arrDBLBracketBlurbs (oRs.Fields("Round").Value, oRs.Fields("SeedOrder").Value + 1) = oRs.Fields("BracketBlurb").Value
		oRs.MoveNext
	Loop
End If

Sub WriteTeamName(iRound, iSeed)
		If IsNull(arrDBLTeamNames(iRound, iSeed)) OR Len(arrDBLTeamNames(iRound, iSeed)) = 0 Then
			Response.Write "TBD"
		ElseIf arrDBLTeamNames(iRound, iSeed) = "-" Then
			Response.Write "-"			
		Else
			If Len(arrDBLTeamNames(iRound, iSeed)) < 15 OR IsNull(arrDBLTeamTags(iRound, iSeed)) OR Len(arrDBLTeamTags(iRound, iSeed)) = 0 Then
				Response.Write "<a href=""/viewteam.asp?team=" & Server.URLEncode(arrDBLTeamNames(iRound, iSeed)) & """>" & Server.HTMLEncode(Left(arrDBLTeamNames(iRound, iSeed), 15)) & "</a>"
			Else
				Response.Write "<a href=""/viewteam.asp?team=" & Server.URLEncode(arrDBLTeamNames(iRound, iSeed)) & """>" & Server.HTMLEncode(Left(arrDBLTeamTags(iRound, iSeed), 15)) & "</a>"
			End If
		End If
End Sub

If (cInt(intDivisionID) <> 0) Then
%>

<table border="0" cellspacing="0" cellpadding="0" width="97%">
<tr>
	<td colspan="7" class="rh" style="border: 1px solid #444444;">Winner's Bracket</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="rh">Round 1</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 1, 1%></td>
	<td class="rh">Round 2</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(1, 1)%></td>
	<td class="wb"><%WriteTeamName 2, 1%></td>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 1, 2%></td>
	<td class="r">&nbsp;</td>
	<td class="rh">Round 3</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(2, 1)%></td>
	<td class="wb"><%WriteTeamName 3, 1%></td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 1, 3%></td>
	<td class="r">&nbsp;</td>
	<td class="r">&nbsp;</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(1, 2)%></td>
	<td class="wb"><%WriteTeamName 2, 2%></td>
	<td class="r">&nbsp;</td>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 1, 4%></td>
	<td class="e">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="rh">Round 4</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(3, 1)%></td>
	<td class="wb"><%WriteTeamName 4, 1%></td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 1, 5%></td>
	<td class="e">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="r">&nbsp;</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(1, 3)%></td>
	<td class="wb"><%WriteTeamName 2, 3%></td>
	<td class="r">&nbsp;</td>
	<td class="r">&nbsp;</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 1, 6%></td>
	<td class="r">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="rh">Finals</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(2, 2)%></td>
	<td class="wb"><%WriteTeamName 3, 2%></td>
	<td class="b"><%=arrDBLBracketBlurbs(4, 1)%></td>
	<td class="wb"><%WriteTeamName 5, 1%></td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 1, 7%></td>
	<td class="r">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="rd">&nbsp;</td>
	<td class="rh">Champion</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(1, 4)%></td>
	<td class="wb"><%WriteTeamName 2, 4%></td>
	<td class="e">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(5, 2)%></td>
	<td class="chmp"><%WriteTeamName 6, 1%></td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 1, 8%></td>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="rd">&nbsp;</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="fwd"><%WriteTeamName 5, 2%></td>
</tr>

<tr>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="r">&nbsp;</td>
</tr>
<tr>
	<td colspan="4" class="rh" style="border: 1px solid #444444;">Loser's Bracket</td>
	<td class="r">&nbsp;</td>
</tr>
<tr>
	<td class="rh">Loser Round 1</td>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="rh">Loser Round 4</td>
	<td class="r">&nbsp;</td>
</tr>
<tr>
	<td class="fl"><%WriteTeamName 1, 9%></td>
	<td class="rh">Loser Round 2</td>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 4, 3%></td>
	<td class="r">&nbsp;</td>
</tr>
</tr>
<tr>
	<td class="b"><%=arrDBLBracketBlurbs(1, 5)%></td>
	<td class="fw"><%WriteTeamName 2, 6%></td>
	<td class="rh">Loser Round 3</td>
	<td class="r">&nbsp;</td>
	<td class="r">&nbsp;</td>
</tr>
<tr>
	<td class="fl"><%WriteTeamName 1, 10%></td>
	<td class="b"><%=arrDBLBracketBlurbs(2, 3)%></td>
	<td class="fl"><%WriteTeamName 3, 3%></td>
	<td class="b"><%=arrDBLBracketBlurbs(4, 2)%></td>
	<td class="fl"><%WriteTeamName 4, 2%></td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fl"><%WriteTeamName 2, 5%></td>
	<td class="r">&nbsp;</td>
	<td class="r">&nbsp;</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(3, 2)%></td>
	<td class="fl"><%WriteTeamName 4, 4%></td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 2, 7%></td>
	<td class="r">&nbsp;</td>
</td>
<tr>
	<td class="fl"><%WriteTeamName 1, 11%></td>
	<td class="b"><%=arrDBLBracketBlurbs(2, 4)%></td>
	<td class="fl"><%WriteTeamName 3, 4%></td>
</tr>
<tr>
	<td class="b"><%=arrDBLBracketBlurbs(1, 6)%></td>
	<td class="fl"><%WriteTeamName 2, 8%></td>
</tr>
<tr>
	<td class="fl"><%WriteTeamName 1, 12%></td>
</tr>

</table>

<%

'Finals Brackets
Else

%>

<table border="0" cellspacing="0" cellpadding="0" width="97%">
<tr>
	<td colspan="7" class="rh" style="border: 1px solid #444444;">Winner's Bracket</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="rh">Round 1</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 4, 1%></td>
	<td class="rh">Round 2</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(1, 1)%></td>
	<td class="wb"><%WriteTeamName 5, 4%></td>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 4, 2%></td>
	<td class="r">&nbsp;</td>
	<td class="rh">Finals</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(2, 1)%></td>
	<td class="wb"><%WriteTeamName 7, 1%></td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 4, 3%></td>
	<td class="r">&nbsp;</td>
	<td class="r">&nbsp;</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(1, 2)%></td>
	<td class="wb"><%WriteTeamName 5, 3%></td>
	<td class="r">&nbsp;</td>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 4, 4%></td>
	<td class="e">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="rh">Champion</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(3, 1)%></td>
	<td class="wb"><%WriteTeamName 1, 1%></td>
</tr>

<tr>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="e">&nbsp;</td>
</tr>
<tr>
	<td colspan="3" class="rh" style="border: 1px solid #444444;">Loser's Bracket</td>
	<td class="r">&nbsp;</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="rh">Loser Round 1</td>
	<td class="e">&nbsp;</td>
	<td class="r">&nbsp;</td>
	<td class="e">&nbsp;</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 5, 6%></td>
	<td class="rh">Loser Round 2</td>
	<td class="r">&nbsp;</td>
	<td class="e">&nbsp;</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="b"><%=arrDBLBracketBlurbs(1, 5)%></td>
	<td class="fl"><%WriteTeamName 6, 1%></td>
	<td class="r">&nbsp</td>
	<td class="e">&nbsp;</td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 5, 5%></td>
	<td class="b"><%=arrDBLBracketBlurbs(6, 2)%></td>
	<td class="fl"><%WriteTeamName 7, 2%></td>
</tr>
<tr>
	<td class="e">&nbsp;</td>
	<td class="e">&nbsp;</td>
	<td class="fw"><%WriteTeamName 6, 2%></td>
	<td class="e">&nbsp;</td>
</tr>


</table>

<%
End If
%>