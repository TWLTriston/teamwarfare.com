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

Dim MatchID, TeamID
Dim mLadderName, mMap1, mMap2, mMap3
Dim mLadderID, mDefenderID, mAttackerID, mDefenderTeamID, mAttackerTeamID
Dim mDefenderName, mAttackerName, mBookedDate, yourteam
Dim i, strMapArray(6), intMaps
MatchID = Request("MatchID")
TeamID = Request("teamID")

strSQL = "Select * from vMatches where MatchID='" & matchid & "'"
ors.Open strSQL, oconn
if not (ors.EOF and ors.BOF) then
	mDefenderID = ors.Fields("MatchDefenderID").Value
	mAttackerID = ors.Fields("MatchAttackerID").Value 
	mBookedDate = ors.Fields("MatchDate").Value
	strMapArray(1) = ors.Fields("MatchMap1ID").Value
	strMapArray(2) = ors.Fields("MatchMap2ID").Value
	strMapArray(3) = ors.Fields("MatchMap3ID").Value
	strMapArray(4) = ors.Fields("MatchMap4ID").Value
	strMapArray(5) = ors.Fields("MatchMap5ID").Value
	mLadderID = ors.Fields("MatchLadderID").Value
	mLadderName = ors.Fields("LadderName").Value
	mDefenderName = ors.Fields("DefenderName").Value 
	mAttackerName = ors.Fields("AttackerName").value
	mDefenderTeamID = oRS.Fields("DefenderTeamID").Value
	mAttackerTeamID = oRs.Fields("AttackerTeamID").Value
	intMaps = oRS.Fields("Maps").Value 
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.clear
	response.redirect "/errorpage.asp?error=7"
End If
ors.Close

If Not(bSysAdmin or IsLadderAdmin(mLadderName) or IsTeamFounder(mAttackerName) Or IsTeamFounder(mDefenderName) Or IsteamCaptain(mAttackerName, mLadderName) Or IsTeamCaptain(mDefenderName, mLadderName)) Then 
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "/errorpage.asp?error=3"
end if

If cStr(mDefenderTeamID) = cStr(TeamID) Then
	YourTeam = mDefenderName
Else
	YourTeam = mAttackerName
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<SCRIPT LANGUAGE="JavaScript">
<!--
	function acceptSubmit(objForm) {
		var err = "n";
		var errStr = "Error: \n";
		<% For i = 1 to intMaps %>
			if (isNaN(objForm.Map<%=i%>DefScore.value)) {
				err = "y";
				errStr = errStr + "You must enter a number for <%=strMapArray(i)%> defender score.\n";
			}
			if (isNaN(objForm.Map<%=i%>AttScore.value)) {
				err = "y";
				errStr = errStr + "You must enter a number for <%=strMapArray(i)%> attacker score.\n";
			}
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
<% Call ContentStart("Report Loss on " & Server.HTMLEncode(mLadderName) & " Ladder") %>

<table border=0 cellpadding=0 cellspacing=0 align=center BGCOLOR="#444444">
<TR><TD>
<table border=0 cellpadding=2 cellspacing=1 align=center WIDTH="100%">
<form name=frmReportLoss action=MatchReportLossValidate.asp method=post>

<tr bgcolor=#000000>
<TH colspan="5"><%=Server.htmlencode(mDefenderName)%> vs <%=Server.htmlencode(mAttackerName)%></tH>
</tr><tr bgcolor=<%=bgctwo%>>
<td align=center colspan="5">Match Date:&nbsp;<input type=text align=right name=matchdate value=<%=formatdatetime(now,2)%> class=bright style="width: 70px"></td>
</tr>
<% For i = 1 To intMaps %>
	<tr BGCOLOR="#000000"><td colspan="5"><img src="/images/spacer.gif" width="1" height="10"></td></tr>

	<tr bgcolor=#000000>
		<TH>&nbsp;Map <%=i%> Results</TH>
		<TH><%=Server.htmlencode(strMapArray(i))%></TH>
		<TH align=center>OT</TH>
		<TH>Forfeit (Loss)</TH>
	</tr>
	<tr bgcolor=<%=bgctwo%>>
		<td align=right><%=Server.htmlencode(mDefenderName)%> Score:</td>
		<td align=left>&nbsp;<input type=text name="Map<%=i%>DefScore" maxlength="5" class=bright style="width: 70px"></td>
		<td align=center><input type=radio class=borderless name="Map<%=i%>OTwin" value=Defender></td>
		<td align=center><input type=radio class=borderless name="Map<%=i%>Forfeit" value=Defender></td>
	</tr>

	<tr bgcolor=<%=bgcone%>>
		<td align=right><%=Server.htmlencode(mAttackerName)%> Score:</td>
		<td align=left>&nbsp;<input type=text name="Map<%=i%>AttScore" maxlength="5" class=bright style="width: 70px"></td>
		<td align=center><input type=radio class=borderless name="Map<%=i%>OTWin" value=Attacker></td>
		<td align=center><input type=radio class=borderless name="Map<%=i%>Forfeit" value=Attacker></td>
	</tr>
	<tr bgcolor=<%=bgctwo%>>
		<td align=center>&nbsp;</td>
		<td align=right>No OT/Forfiet:&nbsp;</td>
		<td align=center><input type=radio class=borderless name="Map<%=i%>OTWin" checked VALUE=""></td>
		<td align=center><input type=radio class=borderless name="Map<%=i%>Forfeit" checked VALUE=""></td>
	</tr>
<% Next %>

<tr BGCOLOR="#000000"><td colspan="5"><img src="/images/spacer.gif" width="1" height="10"></td></tr>
<tr bgcolor=<%=bgctwo%>>
	<td colspan=5 align=center>
	<input type=hidden name=DefenderName value="<%=Server.htmlencode(mDefenderName)%>">
	<input type=hidden name=AttackerName value="<%=Server.htmlencode(mAttackerName)%>">
	<input type=hidden name=DefenderID value=<%=mDefenderID%>>
	<input type=hidden name=AttackerID value=<%=mAttackerID%>>
	<input type=hidden name=Map1 value="<%=Server.htmlencode(strMapArray(1) & "")%>">
	<input type=hidden name=Map2 value="<%=Server.htmlencode(strMapArray(2) & "")%>">
	<input type=hidden name=Map3 value="<%=Server.htmlencode(strMapArray(3) & "")%>">
	<input type=hidden name=Map4 value="<%=Server.htmlencode(strMapArray(4) & "")%>">
	<input type=hidden name=Map5 value="<%=Server.htmlencode(strMapArray(5) & "")%>">
	<INPUT TYPE=hidden NAME="Maps" VALUE="<%=intMaps%>">
	<input type=hidden name=matchid value="<%=matchid%>">
	<input type=hidden name=teamid value="<%=teamid%>">
	<input type=hidden name=yourteam value="<%=Server.htmlencode(yourteam)%>">
	<input type=hidden name=ladderid value='<%=mLadderID%>'>
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

