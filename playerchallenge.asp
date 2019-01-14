<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Challenge"

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

Dim strOpponent, strLadder, strPlayerName
strOpponent = Request.QueryString("opponent")
strLadder = Request.QueryString ("ladder")
strPlayerName = Request.QueryString("playername")

Dim strStatus, playerid, linkid 
Dim strResult, strEnemyName

If Not(bSysAdmin Or IsPlayerLadderAdmin(strLadder) Or Session("uName") = strPlayerName) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "/errorpage.asp?error=3"
end if

strsql= "SELECT lnk.Status, p.PlayerHandle, p.PlayerID, l.PlayerLadderName, lnk.PPLLinkID "
strsql= strsql & "FROM lnk_p_pl lnk, tbl_players p, tbl_playerladders l "
strsql= strsql & "WHERE p.PlayerHandle='" & CheckString(strOpponent) & "' "
strSQL = strSQL & " AND lnk.isactive=1 and p.PlayerID = lnk.PlayerID "
strSQL = strSQL & " AND l.PlayerLadderName='" & CheckString(strLadder) & "' and l.PlayerLadderID = lnk.PlayerLadderID"

ors.open strsql, oconn
if not (ors.eof and ors.bof) then
	strStatus=ors.fields("Status").value
	playerid = ors.Fields("PlayerID").Value
	linkid = ors.Fields ("PPLLinkID").Value  
Else
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=4"
End If
ors.NextRecordset 

if strStatus <> "Available" and left(strStatus, 8) <> "Defeated" then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=4"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Summary of " & Server.HTMLEncode(strOpponent) & " on " & Server.HTMLEncode(strLadder) & " Ladder")
%>
		<table align=center border=0 CELLSPACING=0 WIDTH=75% cellpadding=0 BGCOLOR="#444444">
		<TR>
		<TD>
			<table border=0 WIDTH=100% ALIGN=CENTER cellspacing=1 cellpadding=2>
			<TR BGCOLOR="#000000"><TH COLSPAN=5>Recent History</TH></TR>
			<TR BGCOLOR="#000000">
				<TH>Ladder</TH>
				<TH>Opponent</TH>
				<TH>Result</TH>
				<TH>Date</TH>
			</tr>
<%
	bgc=bgctwo
	strSQL="select TOP 2 * from vPlayerHistory where (matchwinnerid=" & linkID & " or matchloserid=" & linkID & ") and matchforfeit=0 order by matchdate desc"
	ors.Open strSQL, oconn
	if not (ors.eof and ors.BOF) then
		do while not ors.EOF
			If ors.Fields("MatchWinnerID") = linkID Then
				strEnemyName = ors.Fields("LoserName").Value 
				strResult = "Win"
			Else
				strEnemyName = ors.Fields("WinnerName").Value 
				strResult = "Loss"
			End If
			%>
			<tr bgcolor=<%=bgc%>><td>&nbsp;<a href=viewplayerladder.asp?ladder=<%=server.urlencode(ors.Fields("LadderName").Value )%>><%=Server.HTMLEncode(ors.Fields("LadderName").Value)%></a></td>
			<td><a href=viewplayer.asp?player=<%=server.urlencode(strEnemyName)%>><%=Server.HTMLEncode(strEnemyName)%></a></td>
			<td><%=strResult%></td>
			<td align=right><%=ors.Fields("MatchDate").Value%>&nbsp;</td></tr>
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.MoveNext
		loop
	end if
	oRs.NextRecordset 
%>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=4><a href=saveitem.asp?SaveType=playerchallenge&ladder=<%=server.urlencode(Request.QueryString("ladder"))%>&player=<%=server.urlencode(strPlayerName)%>&opponent=<%=server.urlencode(Request.QueryString("opponent"))%> >Submit Challenge</a> - <a href=PlayerLadderAdmin.asp?ladder=<%=server.urlencode(Request.QueryString("ladder"))%>&player=<%=server.urlencode(strPlayerName)%> >Abort Challenge</a></TH>
	</tr>
</TABLE></TD></TR>
</table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

