<% 'Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "Roster Report"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Set oRS2 = Nothing
	Response.Clear 
	Response.Redirect "/errorpage.asp?error=3"
End If
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Roster Report") %>
<form name=ladder id=tournament method=get action=leaguerosterreport.asp>
	<table align=center>
<TR><TD>Ladder:</TD><TD>
<select name=tname id=tname>
<%
strsql = "select LEagueName, LEagueID from tbl_Leagues where LeagueActive = 1 order by LEagueName asc"
ors.open strsql, oconn
if not(ors.eof and ors.bof) then
	do while not(ors.eof)
		Response.Write "<option value=""" & ors.fields(1).value & """"
		if cint(Request.QueryString("lname")) = cint(ors.fields(1).value) then
			Response.Write " selected "
		end if
		Response.Write ">" & ors.fields(0).value & "</option>" & vbcrlf
		ors.movenext
	loop
end if
ors.nextrecordset
%>
</select>
</td></TR>
<TR><TD>Minimum # of players:</TD><TD><select name=numplayer id=numplayer>
<%
for i = 1 to 20
	Response.Write "<option value=" & i
		if cint(Request.QueryString("numplayer")) = cint(i) then
			Response.Write " selected "
		end if
	Response.write ">" & i & "</option>" & vbcrlf
next
%>
</select>
</TD></TR>
<tr><TD align=center colspan=2><input type=submit id=submit name=submit value="Run Query"></TD></TR>
</form>
</Table>
<BR><BR>
<table width="97%" border="0" cellspacing="0" cellpadding="0">
<%
bgc = bgcone
totalteam = 0
undermin = 0
overmin = 0
overdate = 0
underdate = 0
violation = 0
if Request.QueryString("tname") <> "" then
	response.write "<TR><TD><B>Team Name (Rank)</B></TD><TD><B># of players</B></TD><TD><B>Founder Join/Last Login</B></TD></TR>"
	strSQL = "SELECT TeamName, Counter = count(distinct playerid), NULL, "
	strSQL = strSQL & " FounderJoin = (SELECT JoinDate FROM lnk_league_team_player tmp2 WHERE tmp2.PlayerID = tbl_teams.TeamFounderID AND lnkLeagueTeamID = lnk_league_team.lnkLeagueTeamID), TeamFounderID "
	strSQL = strSQL & " FROM lnk_league_team_player "
	strSQL = strSQL & " 	INNER JOIN lnk_league_team ON lnk_league_team.lnkLeagueTeamID = lnk_league_team_player.lnkLeagueTEamID "
	strSQL = strSQL & " 	INNER JOIN tbl_teams ON lnk_league_team.teamid = tbl_teams.teamid "
	strSQL = strSQL & " WHERE lnk_league_team.LeagueID= '" & Request.QueryString("tname") & "' "
	strSQL = strSQL & " GROUP BY lnk_league_team.lnkLeagueTeamID, tbl_teams.teamname, TeamFounderID order by count(*) DESC"
	'Response.Write strsql
	ors.open strsql, oconn
	if not(ors.eof and ors.bof) then
		do while not(ors.eof)
			totalteam = totalteam + 1			
			DateJoined = oRS.fields("FounderJoin").value
			lastlogin = null 'ors.fields("LastLogin").value
			Response.Write "<TR height=20 bgcolor=" & bgc & "><TD>" & totalteam & ". <a href=/viewteam.asp?team=" & server.urlencode(ors.fields(0).value) & ">" & server.htmlencode(ors.fields(0).value) & "</a></TD>"
			counter = ors.fields(1).value
			if cint(counter) < cint(Request.QueryString("numplayer")) then
				text = "<font color=red><B>" & counter & "</B></font>"
				undermin = undermin + 1
			else
				text = counter
				overmin = overmin + 1
			end if
			Response.Write "<TD><B>" & text & "</B></TD>"
			if abs(datediff("d", date(), datejoined)) >= 14 then
				text = "<font color=red><B>" & datejoined & "</B></font>"
				overdate = overdate + 1
				if cint(counter) < cint(Request.QueryString("numplayer")) then
					violation = violation + 1
				end if
			else
				text = datejoined
				underdate = underdate + 1
			end if
			if abs(datediff("d", date(), lastlogin)) >= 14 then
				lastlogin = "<font color=red><B>" & lastlogin & "</B></font>"
			end if
			
			Response.Write "<TD><B>" & text & " / " & lastlogin & "</B></TD></TR>"
			ors.movenext
			if bgc = bgcone then
				bgc = bgctwo
			else
				bgc = bgcone
			end if
		loop
	end if
	ors.nextrecordset	    
    	response.write "<tr height=3><td colspan=3><img src=""/images/spacger.gif"" width=1 height=3></td></TR>"
    	response.write "<TR><TD colspan=3><B>Total teams: " & totalteam & "</B></TD></TR>"
    	response.write "<TR><TD colspan=3><B>Total under size: " & undermin & "</B></TD></TR>"
    	response.write "<TR><TD colspan=3><B>Total over size: " & overmin & "</B></TD></TR>"
    	response.write "<TR><TD colspan=3><B>Total within date: " & underdate& "</B></TD></TR>"
    	response.write "<TR><TD colspan=3><B>Total outside date: " & overdate & "</B></TD></TR>"
    	response.write "<TR><TD colspan=3><font color=red><B>Total in volation: " & violation & "</B></font></TD></TR>"
end if
%>
</table>    
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->

<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>