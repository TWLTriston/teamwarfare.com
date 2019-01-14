<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Change Record"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
Dim strTeamName , strLeagueName
strTeamName = Request.QueryString("team")
strLeagueName = Request.QueryString("league")

Dim intWins, intLosses, intNoShows, intLnkID, intDraws, intRoundsWon, intRoundsLost
Dim intLeaguePoints, intMatchesPlayed
strSQL = "SELECT lnkLeagueTeamID, Wins, Losses, NoShows, Draws, WinPct, RoundsWon, RoundsLost, LeaguePoints, MatchesPlayed FROM lnk_league_team lnk INNER JOIN tbl_teams t ON t.TeamID = lnk.TeamID INNER JOIN tbl_leagues l ON l.LeagueID = lnk.LeagueID "
strSQL = strSQL & " WHERE TeamName = '" & CheckString(strTeamName) & "' AND LeagueName='" & CheckString(strLeagueName) & "'"
'Response.Write strSQL
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) THen
	intLnkID = oRs.FieldS("lnkLeagueTeamID").Value
	intWins = oRs.FieldS("Wins").Value
	intLosses = oRs.FieldS("Losses").Value
	intNoShows = oRs.FieldS("NoShows").Value
	intDraws = oRs.FieldS("Draws").Value
	intRoundsWon = oRs.FieldS("RoundsWon").Value
	intRoundsLost = oRs.FieldS("RoundsLost").Value
	intLeaguePoints = oRs.FieldS("LeaguePoints").Value
	intMatchesPlayed = oRs.FieldS("MatchesPlayed").Value
Else 
	Response.Clear
	Response.Write "Bad Data"
	Response.end
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<html>
<head>
	<title>TWL: Chage team record</title>
	<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
</head>

<body bgcolor="#000000" leftmargin="0" topmargin="00" marginwidth="000" marginheight="0000">
<TABLE height=100% width=100% border=0 cellspacing=0 cellpadding=0 valign=center align=center>
<tr valign=center>
	<td align="center">
	<form name="frmChangeRecord" id="frmChangeRecord" method="post" action="saveitem.asp">
	<input type="hidden" name="lnkLeagueTeamID" id="lnkLeagueTeamID" value="<%=intLnkID%>" />
	<input type="hidden" name="SaveType" id="SaveType" value="ChangeLeagueRecord" />
	<input type="hidden" name="Team" id="Team" value="<%=Server.HTMLEncode(strTeamName)%>" />
	<input type="hidden" name="League" id="League" value="<%=Server.HTMLEncode(strLeagueName)%>" />
	
	<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
	<tr>
		<td>
			<table border="0" cellspacing="1" cellpadding="4">
			<tr>
				<th bgcolor="#000000" colspan="4">Change Team Record</th>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right">Team:</td>
				<td bgcolor="<%=bgcone%>" colspan="3"><%=Server.HTMLEncode(strTeamName)%></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgctwo%>" align="right">League:</td>
				<td bgcolor="<%=bgctwo%>" colspan="3"><%=Server.HTMLEncode(strLeagueName)%></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right">Wins:</td>
				<td bgcolor="<%=bgcone%>"><input type="text" name="wins" id="wins" value="<%=intWins%>" size="5" maxlength="5" /></td>
				<td bgcolor="<%=bgcone%>" align="right">Losses:</td>
				<td bgcolor="<%=bgcone%>"><input type="text" name="Losses" id="Losses" value="<%=intLosses%>" size="5" maxlength="5" /></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgctwo%>" align="right">Draws:</td>
				<td bgcolor="<%=bgctwo%>"><input type="text" name="Draws" id="Draws" value="<%=intDraws%>" size="5" maxlength="5" /></td>
				<td bgcolor="<%=bgctwo%>" align="right">No Shows:</td>
				<td bgcolor="<%=bgctwo%>"><input type="text" name="NoShows" id="NoShows" value="<%=intNoShows%>" size="5" maxlength="5" /></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right">Rounds Won:</td>
				<td bgcolor="<%=bgcone%>"><input type="text" name="RoundsWon" id="RoundsWon" value="<%=intRoundsWon%>" size="5" maxlength="5" /></td>
				<td bgcolor="<%=bgcone%>" align="right">Rounds Lost:</td>
				<td bgcolor="<%=bgcone%>"><input type="text" name="RoundsLost" id="RoundsLost" value="<%=intRoundsLost%>" size="5" maxlength="5" /></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right">League Points:</td>
				<td bgcolor="<%=bgcone%>"><input type="text" name="LeaguePoints" id="LeaguePoints" value="<%=intLeaguePoints%>" size="5" maxlength="5" /></td>
				<td bgcolor="<%=bgcone%>" align="right">Mathes Played:</td>
				<td bgcolor="<%=bgcone%>"><input type="text" name="MatchesPlayed" id="MatchesPlayed" value="<%=intMatchesPlayed%>" size="5" maxlength="5" /></td>
			</tr>
			<tr>
				<td colspan="4" align="center" bgcolor="#000000"><input type="submit" value="Change Record" /></td>
			</tr>
			</table>
		</td>
	</tr>
	</table>
	</form>
	</td>
</tr>
</table>
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>