<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Create League"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

If Not(IsSysAdmin()) Then 
	Response.Clear
	Response.Redirect "errorpage.asp?error=3"
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
End IF

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

dim intGameID
Dim strLeagueName, blnLeagueLocked, blnLeagueActive, intLeagueID, strMode
Dim intWinPoints, intLossPoints, intDrawPoints, intNoShowPoints, blnRosterLocked 
Dim intMapWinPoints, intMapLossPoints, intMapDrawPoints, intMapNoShowPoints
Dim blnLeagueInviteOnly, intScoring, blnSignup, intIdentifierID
strMode = "Add"
If Request.QueryString("IsEdit") = "true" Then
	strLeagueName = Request.QueryString("League")
	strSQL = "SELECT LeagueName, LeagueID, LeagueGameID, LeagueInviteOnly, WinPoints, LossPoints, "
	strSQL = strSQL & " DrawPoints, NoShowPoints, LeagueActive, LeagueLocked, RosterLock, SignUp, "
	strSQL = strSQL & " MapWinPoints, MapLossPoints, MapDrawPoints, MapNoShowPoints, Scoring, IdentifierID "
	strSQL = strSQL & " FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "'"
	oRs.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		strLeagueName = oRS.Fields("LeagueName").Value
		intLeagueID = oRS.FIelds("LeagueID").Value
		blnLeagueInviteOnly = cBool(oRS.Fields("LeagueInviteOnly").Value)
		intGameID = orS.Fields("LeagueGameID").Value
		intWinPoints = oRS.Fields("WinPoints").Value
		intLossPoints = oRS.Fields("LossPoints").Value
		intDrawPoints = oRS.Fields("DrawPoints").Value
		intNoShowPoints = oRS.Fields("NoShowPoints").Value
		intMapWinPoints = oRS.Fields("MapWinPoints").Value
		intMapLossPoints = oRS.Fields("MapLossPoints").Value
		intMapDrawPoints = oRS.Fields("MapDrawPoints").Value
		intMapNoShowPoints = oRS.Fields("MapNoShowPoints").Value
		intScoring = oRS.Fields("Scoring").Value
		blnLeagueActive = cBool(oRS.FIelds("LeagueActive").Value)
		blnLeagueLocked = cBool(oRS.Fields("LeagueLocked").Value)
		blnRosterLocked = cBool(oRS.Fields("RosterLock").Value)
		blnSignup = cBool(oRS.Fields("SignUp").Value)
		intIdentifierID = oRs.Fields("IdentifierID").Value
		
		strMode = "Edit"
	Else
		blnLeagueLocked = 0
		blnLeagueActive = 0
		blnLeagueInviteOnly = 0
	End If
	oRS.NextRecordSet
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart(strMode & " League")
%>
<form name="frmCreateLeague" id="frmCreateLeague" action="saveLeague.asp" method="post">
	<table cellpadding="0" cellspacing="0" border="0" bgcolor="#444444">
	<tr><td>
	<table cellpadding="4" cellspacing="1" border="0">
		<tr>
			<td bgcolor="<%=bgcone%>">League Name:</td>
			<td bgcolor="<%=bgcone%>"><input type="text" style="width: 200px" value="<%=Server.HTMLEncode(strLeagueName & "")%>" name="txtLeagueName" id="txtLeagueName"/></td></tr>
		<tr>
			<td bgcolor="<%=bgctwo%>">Game:</td>
			<td bgcolor="<%=bgctwo%>"><select name="selGameID" id="selGameID">
					<%
					strSQL = "SELECT GameID, GameName FROM tbl_Games WHERE GameID > 0 ORDER BY GameName ASC "
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						Do While Not(oRS.EOF)
							Response.Write "<OPTION VALUE=""" & oRS.Fields("GameID").Value & """ "
							If cStr(oRS.Fields("GameID").Value  & "") = cStr(intGameID & "") Then
								Response.Write " SELECTED "
							End If
							Response.Write ">" & Server.HTMLEncode(oRS.Fields("GameName").Value & "") & "</OPTION>" & vbCrLf
							oRs.MoveNext
						Loop					
					End If
					oRs.NextRecordset
					%></select></td></tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Invite Only?</td>
			<td bgcolor="<%=bgcone%>"><select name="selInviteOnly" id="="selInviteOnly">
					<option value="0" <% If Not(blnLeagueInviteOnly) Then Response.Write " SELECTED " End If %>>No</option>
					<option value="1" <% If blnLeagueInviteOnly Then Response.Write " SELECTED " End If %>>Yes</option>
				</select></td></tr>
		<tr>
			<td bgcolor="<%=bgctwo%>">Points for a win:</td>
			<td bgcolor="<%=bgctwo%>"><input type="text" name="txtWinPoints" id="txtWinPoints" value="<%=intWinPoints%>" style="width:50px"/></td></tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Points for a loss:</td>
			<td bgcolor="<%=bgcone%>"><input type="text" name="txtLossPoints" id="txtLossPoints" value="<%=intLossPoints%>" style="width:50px" /></td></tr>
		<tr>
			<td bgcolor="<%=bgctwo%>">Points for a draw:</td>
			<td bgcolor="<%=bgctwo%>"><input type="text" name="txtDrawPoints" id="txtDrawPoints" value="<%=intDrawPoints%>" style="width:50px" /></td></tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Points for a No Show:</td>
			<td bgcolor="<%=bgcone%>"><input type="text" name="txtNoShowPoints" id="txtNoShowPoints" value="<%=intNoShowPoints%>" style="width:50px" /></td></tr>
		<tr>
			<td bgcolor="<%=bgctwo%>">Points for a map win:</td>
			<td bgcolor="<%=bgctwo%>"><input type="text" name="txtMapWinPoints" id="txtMapWinPoints" value="<%=intMapWinPoints%>" style="width:50px"/></td></tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Points for a map loss:</td>
			<td bgcolor="<%=bgcone%>"><input type="text" name="txtMapLossPoints" id="txtMapLossPoints" value="<%=intMapLossPoints%>" style="width:50px" /></td></tr>
		<tr>
			<td bgcolor="<%=bgctwo%>">Points for a map draw:</td>
			<td bgcolor="<%=bgctwo%>"><input type="text" name="txtMapDrawPoints" id="txtMapDrawPoints" value="<%=intMapDrawPoints%>" style="width:50px" /></td></tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Points for a map no show:</td>
			<td bgcolor="<%=bgcone%>"><input type="text" name="txtMapNoShowPoints" id="txtMapNoShowPoints" value="<%=intMapNoShowPoints%>" style="width:50px" /></td></tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Scoring:</td>
			<td bgcolor="<%=bgcone%>"><select name="selScoring" id="selScoring">
				<option value="0">A match = 1 win on the record</option>
				<option value="1" <% If intScoring = 1 Then Response.Write " selected " end if %>>Each map constitutes a win on the record</option>
				</select>
			</td></tr>
		<tr>
			<td bgcolor="<%=bgctwo%>">Active</td>
			<td bgcolor="<%=bgctwo%>"><select name="selActive" id="selActive">
					<option value="0" <% If Not(blnLeagueActive) Then Response.Write " SELECTED " End If %>>No</option>
					<option value="1" <% If blnLeagueActive Then Response.Write " SELECTED " End If %>>Yes</option>
				</select></td></tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Locked</td>
			<td bgcolor="<%=bgcone%>"><select name="selLocked" id="selLocked">
					<option value="0" <% If Not(blnLeagueLOcked) Then Response.Write " SELECTED " End If %>>No</option>
					<option value="1" <% If blnLeagueLOcked Then Response.Write " SELECTED " End If %>>Yes</option>
				</select></td></tr>
		<tr>
			<td bgcolor="<%=bgctwo%>">Lock Rosters?</td>
			<td bgcolor="<%=bgctwo%>"><select name="selRosterLocked" id="selRosterLocked">
					<option value="0" <% If Not(blnRosterLocked) Then Response.Write " SELECTED " End If %>>No</option>
					<option value="1" <% If blnRosterLocked Then Response.Write " SELECTED " End If %>>Yes</option>
				</select></td></tr>
		<tr>
			<td bgcolor="<%=bgcone%>">Avail for Signups?</td>
			<td bgcolor="<%=bgctwo%>"><select name="selSignup" id="selSignup">
					<option value="0" <% If Not(blnSignup) Then Response.Write " SELECTED " End If %>>No</option>
					<option value="1" <% If blnSignup Then Response.Write " SELECTED " End If %>>Yes</option>
				</select></td></tr>
		<tr bgcolor=<%=bgcone%>><td>Anti-Smurf Criteria:</td><td>
			<select name="selIdentifierID" id="selIdentifierID">
				<option value="">No identifier tracked</option>
				<%
				strSQL = "SELECT IdentifierID, IdentifierName FROM tbl_identifiers ORDER BY IdentifierName ASC "
				oRs.Open strSQL, oConn
				If Not(oRS.EOF AND oRs.BOF) THen
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & oRs.Fields("IdentifierID") & """"
						If CStr(oRs.FieldS("IdentifierID").Value & "") = CStr(intIdentifierID & "") Then
							Response.Write "Selected=""selected"""
						End if
						Response.Write ">" & Server.HTMLEncode(oRs.Fields("IdentifierName").Value) & "</option>" & vBCrLf
						oRs.MoveNext
					Loop
				End If
				oRs.NextRecordSet
				%>
				</select>			
		</td></tr>
 				
		<% If strMode <> "Edit" Then %>
		<tr>
			<td bgcolor="<%=bgcone%>">Number of Conferences: </td>
			<td bgcolor="<%=bgcone%>"><input type="text" name="txtNumConferences" id="txtNumConferences" style="width:50" /></td></tr>
		<% End If %>
		<tr>
			<td colspan="2" bgcolor="#000000" align="center"><input type="submit" value="Save League" /></td></tr>
	</table></td></tr>
	</table>

<input type="hidden" name="saveType" value="LeagueAdd" />
<input type="hidden" name="LeagueID" value="<%=intLeagueID%>" />
<input type="hidden" name="SaveMode" value="<%=strMode%>" />

</form>
					
				
	
			
	
	
	
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>