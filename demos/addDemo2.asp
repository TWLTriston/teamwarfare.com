<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Save Demo File"

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

Dim Winner_ID, Loser_Id, MatchDate, History_ID, Ladder_id, Winner_Name
Dim Loser_name, Ladder_name, msg, MatchName
Dim my_team, my_team_id, my_team_tllink, my_id
Dim map1,map2,map3
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
If Not(Session("LoggedIn")) then 
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Redirect("/errorpage.asp?error=2")
End If

' get winner/loser ID's, make sure player was on one of the teams
strSQL = "SELECT * FROM vHistory WHERE MatchID=" & Request.QueryString("MatchID")
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	winner_id = oRS("MatchWinnerID")
	loser_id = oRS("MatchLoserID")
	MatchDate = oRS("MatchDate")
	history_id = oRS("HistoryID")
	ladder_id = ors("MatchLadderID")
	winner_name = oRS("WinnerName")
	ladder_name = oRS("LadderName")
	loser_name = oRS("LoserName")
	map1 = oRS("MatchMap1")
	map2 = oRS("MatchMap2")
	map3 = oRS("MatchMap3")
	
End If
oRS.NextRecordset 'get team name
strSQL = "SELECT TeamName, TeamID, TLLinkID, PlayerID FROM vPlayerTeams WHERE PlayerHandle='" & CheckString(session("uName")) & "' AND LadderName='" & CheckString(ladder_name) & "'"
oRs.Open strSQL, oConn
If not(ors.Eof And Ors.Bof) Then
	my_team = oRS("TeamName")
	my_team_id = oRS("TeamID")
	my_team_tllink = oRS("TLLinkID")
	my_id = oRS("PlayerID")
EnD If
oRS.NextRecordset 

If my_team_tllink <> winner_id and my_team_tllink <> loser_id Then 
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Redirect("/addDemo.asp")
End If
	
	

' match name to be displayed
MatchName = winner_name & " vs. " & loser_name

bgc=bgcone
Call ContentStart("Add Demo")
%>
	<table width=760 border="0" cellspacing="0" cellpadding="2">
	<tr><td>
	<table width=500 align=center border=0 cellpadding=2 cellspacing=0>
		<tr bgcolor=<%=bgc%>>
			<td>Provide demo details...</td>
			<td align=right>Step: 1 <b>2</b> 3</td>
		</tr>
	</table>
	<br><br>
	
	<% If Msg <> "" Then %>
	<table width=300 align=center>
		<tr>
			<td>You left the following fields blank, please fill them in and submit again.<br>
			<%=Msg%></td>
		</tr>
	</table>
	<%End If%>
	
	<FORM ENCTYPE="MULTIPART/FORM-DATA" METHOD="POST" ACTION="savedemo.asp" id=form1 name=form1>
	<input type=hidden name=MatchID value="<%=Request.QueryString("MatchID")%>">
	<input type=hidden name=IPAddress value="<%=Request.ServerVariables("REMOTE_ADDR") %>">
	<input type=hidden name=PlayerID value="<%=my_id%>">
	<input type=hidden name=HistoryID value="<%=history_id%>">
	<input type=hidden name=TLLinkID value="<%=my_team_tllink%>">
	
	<table WIDTH=50% align=center border=0 cellpadding=0 cellspacing=0 BGCOLOR="#444444">
	<TR><TD>
	<table width=100% align=center border=0 cellpadding=2 cellspacing=1>
		<tr bgcolor="#000000"><tH colspan=2>Demo Of: <%=MatchName%></th></tr>
		<tr bgcolor=<%=bgcone%>>
			<td width=50%><b>POV:</b></td>
			<td width=50%><%=Session("uName")%></td>
		</tr>
		<tr bgcolor=<%=bgctwo%>>
			<td width=50%><b>Team POV:</b></td>
			<td width=50%><%=my_team%></td>
		</tr>		
		<tr bgcolor=<%=bgcone%>>
			<td width=50%><b>Match Date</b></td>
			<td width=50%><%=MatchDate%></b></td>
		</tr>
		<tr bgcolor=<%=bgctwo%>>
			<td width=50%><b>Ladder</b></td>
			<td width=50%><%=ladder_name%></td>			
		</tr>
		<tr bgcolor=<%=bgcone%>>
			<td width=50%><b>Map</b></td>
			<td width=50%><SELECT name=map>
				<OPTION VALUE="">--Choose Map--</OPTION>
				<OPTION VALUE="<%=Map1%>"><%=Map1%></OPTION>
				<OPTION VALUE="<%=Map2%>"><%=Map2%></OPTION>
				<OPTION VALUE="<%=Map3%>"><%=Map3%></OPTION>
			</SELECT>
			</td>
		</tr>
		<tr bgcolor=<%=bgctwo%>>
			<td width=50%><b>Position</b></td>
			<td width=50%><SELECT name=PositionPlayed>
			<OPTION VALUE="">--Choose Position--</OPTION>
			<OPTION VALUE="Offense" <% If Request("PositionPlayed") = "Offense" Then Response.Write " SELECTED "%>>Offense</OPTION>
			<OPTION VALUE="Defense" <% If Request("PositionPlayed") = "Defense" Then Response.Write " SELECTED "%>>Defense</OPTION>
			</SELECT>
		</tr>		
		<tr bgcolor=<%=bgcone%>>
			<td width=50% valign=top><b>Comments</b></td>
			<td width=50%><textarea name=comments rows=5 cols=24><%=Request("comments")%></textarea></td>
		</tr>	
		<tr bgcolor=<%=bgctwo%>>
			<td width=50% valign=top><b>Public?</b></td>
			<td width=50%><input type=checkbox name=public checked><br>
				Public demos are viewable by every TWL member. Demos that are not public are only viewable by the submitter's team.  
			</td>
		</tr>	
		<tr bgcolor=<%=bgcone%>>
			<td valign=top width=50%><b>Admins Only?</b></td>
			<td width=50%><input type=checkbox name=admin><br>
				<font size=1>
					If this box is checked, the demo is only for admin review, no TWL player will have access to it.
				</font>
			</td>
		</tr>	
		<tr bgcolor=<%=bgctwo%>>
			<td valign=top width=50%><b>File</b></td>
			<td width=50%><INPUT TYPE="FILE" NAME="FILE1"><br>
				<font size=1>
					Please upload demos in .zip or .rec format.
				</font>
			</td>
		</tr>		
		<tr bgcolor=<%=bgc%>>
			<td colspan=2 align=center><input type=submit name="Submit" value=" Add Demo "></td>
		</tr>		
	</table>
	</TD></TR></TABLE>
	</form>
	
	</td></tr>
	</table>            
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>