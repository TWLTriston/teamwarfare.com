<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Create Tournament"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim X, tournamentid

If Not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim strTournamentName, intTournamentID, intForumID, intGameID, intActive, intLocked, intSignup
Dim strRulesName, intRosterLock, strHeaderURL, strContentMain, strContentPrizes, strContentSponsors, intHasPrizes, intHasSponsors

strTournamentName = Request.QuerySTring("Tournament")
strSQL = "SELECT TournamentID, Signup, Locked, Active, RosterLock, GameID, ForumID, RulesName, HeaderURL, HasSponsors, HasPrizes, ContentMain, ContentPrizes, ContentSponsors "
strSQL = strSQL & " FROM tbl_tournaments WHERE TournamentName = '" & CheckString(strTournamentName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intTournamentID = oRs.Fields("TournamentID").Value '
	intForumID = oRs.Fields("ForumID").Value '
	intGameID = oRs.Fields("GameID").Value '
	intActive = oRs.Fields("Active").Value '
	intLocked = oRs.Fields("Locked").Value '
	intSignup = oRs.Fields("Signup").Value ' 
	strRulesName = oRs.Fields("RulesName").Value '
	intRosterLock = oRs.Fields("RosterLock").Value '
	strHeaderURL = oRs.Fields("HeaderURL").Value
	intHasPrizes = oRs.Fields("HasPrizes").Value
	intHasSponsors = oRs.Fields("HasSponsors").Value	
	strContentMain = oRs.Fields("ContentMain").Value
	strContentPrizes = oRs.Fields("ContentPrizes").Value
	strContentSponsors = oRs.Fields("ContentSponsors").Value
End If
oRs.NextRecordSet
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("Edit Tournament: " & strTournamentName)
%>
<table align=center cellpadding="0" cellspacing="0" bgcolor="#444444">
<form name="frmEditTournament" id="frmEditTournamentName" action="saveTournament.asp" method="post">
<input type="hidden" name="SaveType" id="SaveType" value="EditTournament" />
<input type="hidden" name="TournamentID" id="TournamentID" value="<%=intTournamentID%>" />
<tr>
	<td>
	<table align=center cellpadding="4" cellspacing="1">
	<tr>
		<th bgcolor="#00000" colspan="2">Edit Tournament</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Tournament Name</td>
		<td bgcolor="<%=bgctwo%>"><input type="text" name="TournamentName" id="TournamentName" value="<%=Server.HTMLENcode(strTournamentName)%>" size="50" /></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Rules Name</td>
		<td bgcolor="<%=bgctwo%>"><input type="text" name="RulesName" id="RulesName" value="<%=Server.HTMLENcode(strRulesName & "")%>" size="50" /></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Header Image URL</td>
		<td bgcolor="<%=bgctwo%>"><input type="text" name="HeaderURL" id="HeaderURL" value="<%=Server.HTMLENcode(strHeaderURL & "")%>" size="50" /></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Game</td>
		<td bgcolor="<%=bgctwo%>">
			<select name="GameID" id="GameID" class="text">
			<%
				strSQL = "SELECT GameID, GameName FROM tbl_Games WHERE GameID > 0 ORDER BY GameName ASC "
				oRS.Open strSQL, oConn
				If Not(oRS.EOF AND oRS.BOF) Then
					Do While Not(oRS.EOF)
						Response.Write "<OPTION VALUE=""" & oRS.Fields("GameID").Value & """ "
						If cStr(oRS.Fields("GameID").Value) = cStr(intGameID) Then
							Response.Write "selected=""selected"""
						End If
						Response.Write ">" & Server.HTMLEncode(oRS.Fields("GameName").Value & "") & "</OPTION>" & vbCrLf
						oRs.MoveNext
					Loop					
				End If
				oRs.NextRecordset
			%>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Forum</td>
		<td bgcolor="<%=bgctwo%>">
			<select name="ForumID" id="ForumID" class="text">
				<%
					strSQL = "SELECT tbl_forums.ForumID, tbl_forums.ForumName "
					strSQL = strSQL & " FROM tbl_category, tbl_forums "
					strSQL = strSQL & " WHERE tbl_forums.CategoryID = tbl_category.CategoryID "
					strSQL = strSQL & " AND tbl_category.CategoryOrder >= 0 "
					strSQL = strSQL & " ORDER BY tbl_forums.ForumName ASC"
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						Do While Not(oRS.EOF)
							Response.Write "<OPTION VALUE=""" & oRS.Fields("ForumID").Value & """ "
							If cStr(oRS.Fields("ForumID").Value  & "") = cStr(intForumID & "") Then
								Response.Write " SELECTED "
							End If
							Response.Write ">" & Server.HTMLEncode(oRS.Fields("ForumName").Value & "") & "</OPTION>" & vbCrLf
							oRs.MoveNext
						Loop					
					End If
					oRs.NextRecordset
				%>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Active / Visible to visitors</td>
		<td bgcolor="<%=bgctwo%>">
			<select name="Active" id="Active" class="text">
				<option value="0" <% If intActive = "0" Then Response.Write " selected " %>>No</option>
				<option value="1" <% If intActive = "1" Then Response.Write " selected " %>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Admin Panel Locked</td>
		<td bgcolor="<%=bgctwo%>">
			<select name="Locked" id="Locked" class="text">
				<option value="0" <% If intLocked = "0" Then Response.Write " selected " %>>No</option>
				<option value="1" <% If intLocked = "1" Then Response.Write " selected " %>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Open for Signups</td>
		<td bgcolor="<%=bgctwo%>">
			<select name="SignUp" id="SignUp" class="text">
				<option value="0" <% If intSignup = "0" Then Response.Write " selected " %>>No</option>
				<option value="1" <% If intSignup = "1" Then Response.Write " selected " %>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Roster Locked</td>
		<td bgcolor="<%=bgctwo%>">
			<select name="RosterLock" id="RosterLock" class="text">
				<option value="0" <% If intRosterLock = "0" Then Response.Write " selected " %>>No</option>
				<option value="1" <% If intRosterLock = "1" Then Response.Write " selected " %>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Has Sponsors</td>
		<td bgcolor="<%=bgctwo%>">
			<select name="HasSponsors" id="HasSponsors" class="text">
				<option value="0" <% If intHasSponsors = "0" Then Response.Write " selected " %>>No</option>
				<option value="1" <% If intHasSponsors = "1" Then Response.Write " selected " %>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Has Prizes</td>
		<td bgcolor="<%=bgctwo%>">
			<select name="HasPrizes" id="HasPrizes" class="text">
				<option value="0" <% If intHasPrizes = "0" Then Response.Write " selected " %>>No</option>
				<option value="1" <% If intHasPrizes = "1" Then Response.Write " selected " %>>Yes</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2" bgcolor="#000000"><input type="Submit" name="submit" value="Edit Tournament"></td>
	</tr>
	</table>
	</td>
</tr>
</form>
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