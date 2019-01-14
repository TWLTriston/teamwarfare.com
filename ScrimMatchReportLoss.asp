<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Report Loss"

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

Dim intMatchID, intLinkID
Dim strDefenderName, strAttackerName, intLadderID, strTeamName, intAttackerEloID, intDefenderEloID
Dim strLadderName, strMapArray(6)
Dim i, intMaps

intMaps = 5
intMatchID = Request("MatchID")
intLinkID = Request("LinkID")

strSQL = "SELECT 'DefenderName' = d.TeamName, 'AttackerName' = a.TeamName, MatchDate, AttackerEloTeamID, DefenderEloTeamID, Map1, Map2, Map3, Map4, Map5, l.EloLadderName, l.EloLadderID "
strSQL = strSQL & " FROM tbl_elo_matches em "
strSQL = strSQL & " INNER JOIN tbl_elo_ladders l ON l.EloLadderID = em.EloLadderID "
strSQL = strSQL & " INNER JOIN lnk_elo_team aet ON aet.lnkEloTeamID = em.AttackerEloTeamID "
strSQL = strSQL & " INNER JOIN lnk_elo_team det ON det.lnkEloTeamID = em.DefenderEloTeamID "
strSQL = strSQL & " INNER JOIN tbl_teams a ON a.TeamID = aet.TeamID "
strSQL = strSQL & " INNER JOIN tbl_teams d ON d.TeamID = det.TeamID "
strSQL = strSQL & " WHERE em.EloMatchID = '" & intMatchID & "'"
oRs.Open strSQL, oconn
If Not (oRs.EOF AND oRs.BOF) Then
	strDefenderName = oRs.Fields("DefenderName").Value
	strAttackerName = oRs.Fields("AttackerName").Value
	intAttackerEloID = oRs.Fields("AttackerEloTeamID").Value
	intDefenderEloID = oRs.Fields("DefenderEloTeamID").Value
	strLadderName = oRs.Fields("EloLadderName").Value
	intLadderID = oRs.Fields("EloLadderID").Value
	strMapArray(1) = oRs.Fields("Map1").Value
	strMapArray(2) = oRs.Fields("Map2").Value
	strMapArray(3) = oRs.Fields("Map3").Value
	strMapArray(4) = oRs.Fields("Map4").Value
	strMapArray(5) = oRs.Fields("Map5").Value
'	dtmMatchDate = oRs.Fields("MatchDate").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.clear
	response.redirect "/errorpage.asp?error=7"
End If
oRs.NextRecordSet

If CInt(intLinkID) = intAttackerEloID THen
	strTeamName = strAttackerName
Else
	strTeamName = strDefenderName
End If

If Not(bSysAdmin or IsEloLadderAdmin(strLadderName) or IsTeamFounder(strTeamName) OR IsEloTeamCaptain(strTeamName, strLadderName) ) Then 
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "/errorpage.asp?error=3"
end if

If Len(strMapArray(1)) = 0 Then
	intMaps = 0
ElseIf Len(strMapArray(2)) = 0 Then
	intMaps = 1
ElseIf Len(strMapArray(3)) = 0 Then
	intMaps = 2
ElseIf Len(strMapArray(4)) = 0 Then
	intMaps = 3
ElseIf Len(strMapArray(5)) = 0 Then
	intMaps = 4
Else
	intMaps = 5
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->
<SCRIPT LANGUAGE="JavaScript">
<!--
	function acceptSubmit(objForm) {
		var err = "n";
		var errStr = "Error: \n";
		<% For i = 1 to intMaps 
			If Len(strMapArray(i)) > 0 Then %>
			if (isNaN(objForm.Map<%=i%>DefScore.value)) {
				err = "y";
				errStr = errStr + "You must enter a number for <%=strMapArray(i)%> defender score.\n";
			}
			if (isNaN(objForm.Map<%=i%>AttScore.value)) {
				err = "y";
				errStr = errStr + "You must enter a number for <%=strMapArray(i)%> attacker score.\n";
			}
			<% End If %>
		<% Next %>
		if (err == "n") {
			// alert("Submitting...");
			objForm.submit();
		} else {
			alert(errStr);
		}
	}
//-->
</SCRIPT>
<% Call ContentStart("Report Loss on " & Server.HTMLEncode(strLadderName) & " Ladder") %>

<table border=0 cellpadding=0 cellspacing=0 align=center BGCOLOR="#444444">
<form name=frmReportLoss action=ScrimMatchReportLossValidate.asp method=post>
<tr>
	<td>
		<table border=0 cellpadding=4 cellspacing=1 align=center WIDTH="100%">
		<tr bgcolor=#000000>
			<th colspan="5"><%=Server.htmlencode(strDefenderName)%> vs <%=Server.htmlencode(strAttackerName)%></tH>
		</tr>
		<tr bgcolor=<%=bgctwo%>>
			<td align=center colspan="5">Match Date:&nbsp;<input type=text align=right name="MatchDate" value=<%=FormatDateTime(Now(),2)%> class=bright style="width: 70px"></td>
</tr>
<% For i = 1 To intMaps 
	If Len(strMapArray(i)) > 0 Then 
		%>
		<tr BGCOLOR="#000000"><td colspan="5"><img src="/images/spacer.gif" width="1" height="10"></td></tr>
	
		<tr bgcolor=#000000>
			<TH>&nbsp;Map <%=i%> Results</TH>
			<TH><%=Server.htmlencode(strMapArray(i))%></TH>
			<TH align=center>OT</TH>
		</tr>
		<tr bgcolor=<%=bgctwo%>>
			<td align=right><%=Server.htmlencode(strDefenderName)%> Score:</td>
			<td align=left>&nbsp;<input type=text name="Map<%=i%>DefScore" maxlength="5" class=bright style="width: 70px"></td>
			<td align=center><input type=radio class=borderless name="Map<%=i%>OTwin" value=Defender></td>
		</tr>
	
		<tr bgcolor=<%=bgcone%>>
			<td align=right><%=Server.htmlencode(strAttackerName)%> Score:</td>
			<td align=left>&nbsp;<input type=text name="Map<%=i%>AttScore" maxlength="5" class=bright style="width: 70px"></td>
			<td align=center><input type=radio class=borderless name="Map<%=i%>OTWin" value=Attacker></td>
		</tr>
		<tr bgcolor=<%=bgctwo%>>
			<td align=center>&nbsp;</td>
			<td align=right>No OT:&nbsp;</td>
			<td align=center><input type=radio class=borderless name="Map<%=i%>OTWin" checked VALUE=""></td>
		</tr>
	<% End If %>
<% Next %>

<tr BGCOLOR="#000000"><td colspan="5"><img src="/images/spacer.gif" width="1" height="10"></td></tr>
<tr bgcolor=<%=bgctwo%>>
	<td colspan=5 align=center>
	<input type=hidden name=DefenderName value="<%=Server.htmlencode(strDefenderName)%>">
	<input type=hidden name=AttackerName value="<%=Server.htmlencode(strAttackerName)%>">
	<input type=hidden name=DefenderID value=<%=intDefenderEloID%>>
	<input type=hidden name=AttackerID value=<%=intAttackerEloID%>>
	<input type=hidden name=Map1 value="<%=Server.htmlencode(strMapArray(1) & "")%>">
	<input type=hidden name=Map2 value="<%=Server.htmlencode(strMapArray(2) & "")%>">
	<input type=hidden name=Map3 value="<%=Server.htmlencode(strMapArray(3) & "")%>">
	<INPUT TYPE=hidden NAME="Maps" VALUE="<%=intMaps%>">
	<input type=hidden name=matchid value="<%=intMatchID%>">
	<input type=hidden name=linkid value="<%=intLinkID%>">
	<input type=hidden name=yourteam value="<%=Server.htmlencode(strTeamName)%>">
	<input type=hidden name=ladderid value="<%=intLadderID%>">
	<input type=BUTTON name=submit1 value="Report Loss" class=bright ONCLICK="javaScript:acceptSubmit(this.form);"></td>
</tr>
</form>
</table>
</TD></TR>
</TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

