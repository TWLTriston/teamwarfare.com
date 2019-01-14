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

If request.form("submit") = "Create Tournament" Then
	
	strSQL = "INSERT INTO tbl_tournaments (TournamentName,divisions,teamsPerDiv,GameID,TournamentStyle, FinalsStyle) VALUES ('"
	strSQL = strSQL & CheckString(request.form("name")) & "', '"
	strSQL = strSQL & request.form("divisions") & "', '"
	strSQL = strSQL & request.form("teamsPerDiv") & "', '" & Request.Form("GameID") & "', '" & Request.Form("selTournamentStyle") & "', '" & Request.Form("selFinalsStyle") & "')"
	oConn.Execute (strSQL )
	strSQL = "select TournamentID from tbl_tournaments where TournamentName='" & CheckString(Request.Form("Name")) & "'"
	ors.open strSQL , oconn
	if not(ors.eof and ors.bof) then
		TournamentID = ors.fields(0).value
	end if
	ors.nextrecordset
	for x = 1 to Request.Form("Divisions")
		strSQL  = "insert into tbl_Tdivisions (DivisionID, TournamentID, DivisionName) values "
		strSQL  = strSQL  & "('" & x & "', '" & TournamentID & "', '" & replace(Request.Form("Name"), "'", "''") & " Div " & x & "')"
'		Response.Write sql
		oConn.execute (strSQL )
	next
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear 
	Response.Redirect("/tournament/createtourny.asp")
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("Add a new Tournament")
%>
<table align=center cellpadding="4" cellspacing="0" bgcolor="#000000">
<form method="post" action="createtourny.asp">
	<tr>
		<td><b>Tournament Name:</b></td>
		<td><input type="Text" name="name"></td>
	</tr>
	<tr>
		<td><b>Number of Divisions:</b></td>
		<td><select name="divisions">
		<option value="1">1 Division
		<option value="2">2 Divisions
		<option value="3">3 Divisions
		<option value="4">4 Divisions
		<option value="8">8 Divisions
		</select></td>
	</tr>
	<tr>
		<td><b>Teams per Division:</b></td>
		<td><select name="teamsperdiv">
		<option value="4">4 Teams
		<option value="8">8 Teams
		<option value="16">16 Teams
		<option value="32">32 Teams
		</select></td>
	</tr>
		<tr><td><b>Game:</b></td><td><SELECT NAME=GameID Class=text>
				<%
					strSQL = "SELECT GameID, GameName FROM tbl_Games WHERE GameID > 0 ORDER BY GameName ASC "
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						Do While Not(oRS.EOF)
							Response.Write "<OPTION VALUE=""" & oRS.Fields("GameID").Value & """ "
							Response.Write ">" & Server.HTMLEncode(oRS.Fields("GameName").Value & "") & "</OPTION>" & vbCrLf
							oRs.MoveNext
						Loop					
					End If
					oRs.NextRecordset
					%>
					</SELECT></td></tr>
	<tr>
		<td><b>Tournament Style:</b></td>
		<td><SELECT NAME=selTournamentStyle Class=text>
			<option value="S">Single Elimination</option>
			<option value="D">Double Elimination</option>
			</select>
		</td>
	</tr>
	<tr>
		<td><b>Finals Style:</b></td>
		<td><SELECT NAME=selFinalsStyle Class=text>
			<option value="S">Single Elimination</option>
			<option value="D">Double Elimination</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align="center" colspan="2"><input type="Submit" name="submit" value="Create Tournament"></td>
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