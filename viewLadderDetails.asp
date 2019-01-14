<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("ladder"), """", "&quot;")  & " Ladder Information"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strLadderName, intLadderID

strLadderName = Request.QueryString("ladder")

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart(Server.HTMLEncode(strLadderName) & " Ladder Information")
%>
<%
Dim strGameName, strGameAbbr, intForumID
Dim strLadderAbbr, strLocked
Dim i, strConfig

strSQL = "SELECT LadderID, LadderName, LadderAbbreviation, LadderLocked, LadderChallenge, LadderRules, "
strSQL = strSQL & " RosterLimit, MinPlayer, TimeZone, TimeOptions, Maps, MapConfiguration, "
strSQL = strSQL & " Scoring, GameID, RequireDistinctMaps, MatchDays, ChallengeDays, "
strSQL = strSQL & " LadderActive, LadderForumID FROM tbl_ladders WHERE LadderName = '" & CheckString(strLadderName) & "'"
oRs.Open strSQL, oConn
If oRs.EOF Then
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "ladderlist.asp?error=1"
Else
	intLadderID =  oRs.Fields("LadderID").Value
	strLadderName = oRs.Fields("LadderName").Value
	strLadderAbbr = oRs.Fields("LadderAbbreviation").Value
	If oRs.FieldS("LadderLocked") Then
		strLocked = "<b><font color=""red"">LOCKED</font></b>"
	Else
		strLocked = "<b><font color=""green"">Open for challenges</font></b>"
	End If
	%>
	<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
	<tr>
		<td>
		<table border="0" cellspacing="1" cellpadding="4">
		<tr>
			<th colspan="2" bgcolor="#000000">General Information</th>
		</tr>
		<%
		strSQL = "SELECT GameName, GameAbbreviation, ForumID FROM tbl_games WHERE GameID = '" & oRs.Fields("GameID").Value & "'"
		oRs2.Open strSQL, oConn
		If Not (oRs2.EOF AND oRs2.BOF) Then
			strGameName = oRs2.Fields("GameName").Value
			strGameAbbr = oRs2.Fields("GameAbbreviation").Value
			intForumID = oRs2.Fields("ForumID").Value
						
		End If
		oRs2.NextRecordSet
		If Not(IsNull(oRs.Fields("LadderForumID").Value)) AND oRs.Fields("LadderForumID").Value > 0 Then
			intForumID = oRs.Fields("LadderForumID").Value
		End If
		%>
		<tr>
			<td bgcolor="<%=bgctwo%>"><b>Name</b></td>
			<td bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(strLadderName & "")%> Ladder - <%=strLadderAbbr%> - <a href="viewladder.asp?ladder=<%=Server.URLEncode(strLadderName & "")%>">click to view</a> </td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>"><b>Game</b></td>
			<td bgcolor="<%=bgcone%>"><%=Server.HTMLEncode(strGameName & "")%> - <%=strGameAbbr%></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgctwo%>"><b>Status</b></td>
			<td bgcolor="<%=bgctwo%>"><%=strLocked%></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>"><b>Admins</b></td>
			<td bgcolor="<%=bgcone%>"><%
			strSQL = "SELECT PlayerHandle, PlayerEmail FROM tbl_players p "
			strSQL = strSQL & " INNER JOIN lnk_l_a l ON p.PlayerID = l.PlayerID "
			strSQL = strSQL & " WHERE l.LadderID = '" & oRs.Fields("LadderID").Value & "'"
			strSQL = strSQL & " ORDER BY l.PrimaryAdmin DESC, p.PlayerHandle ASC "
			oRs2.Open strSQL, oConn
			If Not(oRs2.EOF AND oRs2.BOF) THen
				Do While Not (oRS2.EOF)
					Response.Write "<a href=""mailto:" & oRs2.Fields("PlayerEmail").Value & """>" & Server.HTMLEncode(oRs2.Fields("PlayerHandle").Value & "") & "</a>"
					ors2.movenext
					If Not(ors2.eof) Then
						Response.Write ", "
					End If
				Loop
			End If
			oRs2.NextRecordSet
			%></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgctwo%>"><b>Rules</b></td>
			<td bgcolor="<%=bgctwo%>"><a href="rules.asp?set=<%=Server.URLEncode(oRs.Fields("LadderRules").Value & "")%>">Click here to view rules</a></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>"><b>Minimum Roster</b></td>
			<td bgcolor="<%=bgcone%>"><%=oRs.Fields("MinPlayer").Value%> players</td>
		</tr>
		<% If oRs.Fields("RosterLimit").Value > 0 Then %>
		<tr>
			<td bgcolor="<%=bgctwo%>"><b>Roster Limit</b></td>
			<td bgcolor="<%=bgctwo%>"><%=oRs.Fields("RosterLimit").Value%> players</td>
		</tr>
		<% Else %>
		<tr>
			<td bgcolor="<%=bgcone%>"><b>Roster Limit</b></td>
			<td bgcolor="<%=bgcone%>"><b>None</b></td>
		</tr>
		<% End If %>
		<tr>
			<td bgcolor="<%=bgctwo%>"><b>Challenge Rungs</b></td>
			<td bgcolor="<%=bgctwo%>"><%=oRs.Fields("LadderChallenge").Value%> rungs or 50% of your current rank, whichever is greater</td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>"><b>Maps Per Match</b></td>
			<% If oRs.Fields("Maps").Value = 1 Then %>
			<td bgcolor="<%=bgcone%>"><%=oRs.Fields("Maps").Value%> map</td>
			<% Else %>
			<td bgcolor="<%=bgcone%>"><%=oRs.Fields("Maps").Value%> maps</td>
			<% End If %>
		</tr>
		<tr>
			<td bgcolor="<%=bgctwo%>" valign="top"><b>Map Selection</b></td>
			<td bgcolor="<%=bgctwo%>">
			<%
			For i = 1 to oRs.Fields("Maps").Value
				Select Case (mid(oRs.Fields("MapConfiguration").Value, i, 1))
					case "R"
						strConfig = "Map #" & i & " - Randomly Selected<br />"
					case "D"
						strConfig = "Map #" & i & " - Defender Choice<br />"
					case "A"
						strConfig = "Map #" & i & " - Attacker Choice<br />"
					case "C"
						strConfig = "Map #" & i & " - Attacker Choice - at time of challenge<br />"
				End Select
				Response.Write strConfig
			Next
			%></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>"><b>Challenges can be initiated on</b></td>
			<td bgcolor="<%=bgcone%>">
			<%
			Dim intChallengeDays
			intChallengeDays = oRs.Fields("ChallengeDays").Value
			i = 0
			strConfig = ""
			if (intChallengeDays and 2 ^ vbSunday) Then 
				i = 1
				strConfig = "Sunday"
			End If
			if (intChallengeDays and 2 ^ vbMonday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Monday"
			End If
			if (intChallengeDays and 2 ^ vbTuesday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Tuesday"
			End If
			if (intChallengeDays and 2 ^ vbWednesday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Wednesday"
			End If
			if (intChallengeDays and 2 ^ vbThursday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Thursday"
			End If
			if (intChallengeDays and 2 ^ vbFriday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Friday"
			End If
			if (intChallengeDays and 2 ^ vbSaturday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Saturday"
			End If
			Response.Write strConfig
			%></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgctwo%>"><b>Matches can be played on</b></td>
			<td bgcolor="<%=bgctwo%>">
			<%
			Dim intMatchDays
			intMatchDays = oRs.Fields("MatchDays").Value
			i = 0
			strConfig = ""
			if (intMatchDays and 2 ^ vbSunday) Then 
				i = 1
				strConfig = "Sunday"
			End If
			if (intMatchDays and 2 ^ vbMonday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Monday"
			End If
			if (intMatchDays and 2 ^ vbTuesday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Tuesday"
			End If
			if (intMatchDays and 2 ^ vbWednesday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Wednesday"
			End If
			if (intMatchDays and 2 ^ vbThursday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Thursday"
			End If
			if (intMatchDays and 2 ^ vbFriday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Friday"
			End If
			if (intMatchDays and 2 ^ vbSaturday) Then 
				If i > 0 Then
					strConfig = strConfig & ", "
				End If
				i = 1
				strConfig = strConfig & "Saturday"
			End If
			Response.Write strConfig
			%></td>
		</tr>
		<tr>
			<td bgcolor="<%=bgcone%>"><b>Matches can be played at</b></td>
			<td bgcolor="<%=bgcone%>"><%=Replace(oRs.Fields("TimeOptions").Value, "|", ", ") & " - " & oRs.Fields("TimeZone").Value%></td>
		</tr>
		</table>
		</td>
	</tr>
	</table>
	<%
End If
oRs.NextRecordSet
%>
<br /><br />
	<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
	<tr>
		<td>
		<table border="0" cellspacing="1" cellpadding="4">
		<tr>
			<th colspan="3" bgcolor="#000000">Map List</th>
		</tr>

<%
strSQL = "SELECT DISTINCT MapName FROM tbl_maps m "
strSQL = strSQL & " INNER JOIN lnk_l_m l ON l.MapID = m.MapID WHERE LadderID= '" & intLadderID & "' ORDER BY MapName ASC "
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	i = 0
	Do While Not (oRS.EOF)
		if i mod 3 = 0 then
			If i > 0 Then
				Response.Write "</tr>"
			End if
			Response.Write "<tr>"
		End If
		i = i + 1
		Response.Write "<td bgcolor=""" & bgcone & """>" & oRs.FIeldS("MapName").Value & "</td>" & vbCrLf
		oRs.MoveNext
	Loop
	If i mod 3 <> 0 then
		Response.Write "<td bgcolor=""#000000"" colspan=""" & i mod 3 + 1 & """>&nbsp;</td>" & vbCrLf
	End If
	Response.Write "</tr>"
End If
oRs.NextRecordSet
%>
	</table>
	</td>
	</tr>
	</table>
<%

Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>