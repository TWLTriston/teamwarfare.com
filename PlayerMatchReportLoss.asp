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

Dim MatchID, LinkID
Dim mLadderName, mMap1
Dim mLadderID, mDefenderID, mAttackerID, mDefenderTeamID, mAttackerTeamID
Dim mDefenderName, mAttackerName, mBookedDate, yourteam, ReporterName
MatchID = Request("MatchID")
LinkID = Request("LinkID")
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<%
	strSQL = "Select * from vPlayerMatches where PlayermatchID='" & matchid & "'"
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		mDefenderID=ors.Fields("MatchDefenderID").Value
		mAttackerID=ors.Fields("MatchAttackerID").Value 
		mBookedDate=ors.Fields("MatchDate").Value
		mMap1=ors.Fields("MatchMap1ID").Value
		mLadderID=ors.Fields("MatchLadderID").Value
		mLadderName=ors.Fields("PlayerLadderName").Value
		mDefenderName=ors.Fields("DefenderName").Value 
		mAttackerName=ors.Fields("AttackerName").value
 	end if
	ors.Close
	reporterName = request("Player") 
	
	if not (bSysAdmin OR session("uName") = ReporterName OR IsPlayerLadderAdmin(mLadderName)) Then
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		response.clear
		response.redirect "errorpage.asp?error=3"
	end if

	
	If mDefenderID = LinkID Then
		YourTeam=mDefenderName
 	Else
 		YourTeam=mAttackerName
 	End If
 	
 	Call ContentStart("Report Loss on " & Server.HTMLEncode (mLadderName) & " Ladder")
%><table border=0 cellpadding=0 cellspacing=0 align=center BGCOLOR="#444444">
<TR><TD>
<table border=0 cellpadding=2 cellspacing=1 align=center WIDTH="100%">
<form name=frmReportLoss action=/playerMatchReportLossValidate.asp method=post>

<tr bgcolor=#000000>
<TH colspan="5"><%=Server.htmlencode(mDefenderName)%> vs <%=Server.htmlencode(mAttackerName)%></tH>
</tr><tr bgcolor=<%=bgctwo%>>
<td align=center colspan="7">Match Date:&nbsp;<input type=text align=right name=matchdate value=<%=formatdatetime(now,2)%> class=bright style="width: 70px"></td>
</tr>

<tr bgcolor=<%=bgcone%>>
<td align=left>&nbsp;Map 1 Results</td>
<td align=center><%=Server.HTMLEncode (mMap1)%></td>
<td align=center>Forfiet</td>
</tr>
<tr bgcolor=<%=bgctwo%>>
	<td align=right><%=Server.HTMLEncode (mDefenderName)%> Score:</td>
	<td align=left>&nbsp;<input type=text name=Map1DefScore class=bright style="width: 70px"></td>
	<td align=center><input type=radio class=borderless name=Map1Forfeit value=Defender></td>
</tr>

<tr bgcolor=<%=bgcone%>>
	<td align=right><%=Server.HTMLEncode (mAttackerName)%> Score:</td>
	<td align=left>&nbsp;<input type=text name=Map1AttScore class=bright style="width: 70px"></td>
	<td align=center><input type=radio class=borderless name=Map1Forfeit value=Attacker></td>
</tr>
<tr bgcolor=<%=bgctwo%> height=20>
	<td align=center>&nbsp;</td>
	<td align=right>No Forfiet:&nbsp;</td>
	<td align=center><input type=radio class=borderless name=Map1Forfeit checked></td>
</tr>
<tr bgcolor=<%=bgctwo%> height=25>
	<td colspan=3 align=center>
	<input type=hidden name=DefenderName value="<%=Server.HTMLEncode (mDefenderName)%>">
	<input type=hidden name=AttackerName value="<%=Server.HTMLEncode (mAttackerName)%>">
	<input type=hidden name=DefenderID value=<%=mDefenderID%>>
	<input type=hidden name=AttackerID value=<%=mAttackerID%>>
	<input type=hidden name=Map1 value="<%=Server.HTMLEncode (mMap1)%>">
	<input type=hidden name=matchid value='<%=matchid%>'>
	<input type=hidden name=linkId value='<%=LinkID%>'>
	<input type=hidden name=Reporter value="<%=Server.HTMLEncode (reporterName)%>">
	<input type=hidden name=ladderid value='<%=mLadderID%>'>
	<input type=submit name=submit1 value="Report Loss" class=bright></td>
</tr>
</form>
</table>
</TD></TR></TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>
