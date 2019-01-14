<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Validate Loss Report"

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
Dim mDefenderName, mAttackerName, mBookedDate, yourteam, mDid, mAid, mMatchDate
Dim mMap1DefScore, mMap2DefScore, mMap3DefScore
Dim mMap1AttScore, mMap2AttScore, mMap3AttScore
Dim mMap1OTWin, mMap2OTWin, mMap3OTWin
Dim mMap1Forfeit, mMap2Forfeit, mMap3Forfeit
Dim dScore, aScore, m1c, m2c, m3c
Dim mWinner1, mWinner2, mWinner3
Dim pScore, mMatchWinnerIsDefender, mWinner, mLoser, mWinnerID, mLoserID
Dim Map1IsOT, Map2IsOT, Map3IsOT
Dim Map1Forfeit, Map2Forfeit, Map3Forfeit

matchid=Request.Form("matchid")
teamid = Request.Form("teamid")
mLadderID = Request.Form("ladderid")
mDefenderID = request.form("DefenderID")
mAttackerID = request.form("AttackerID")%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Validate Match Report") %>

<%

'mDefenderName=Request.Form("DefenderName")
'mAttackerName=Request.Form("AttackerName")	
strSQL = "SELECT p.PlayerHandle FROM tbl_players p, lnk_p_pl lnk where lnk.PlayerID = p.PLayerID and lnk.PPLLinkID = " & mDefenderID
	ors.Open strSQL, oconn
	if not (ors.EOF and ors.BOF) then
		mDefenderName=ors.Fields(0).Value 
	end if
	ors.Close
	strSQL = "SELECT p.PlayerHandle FROM tbl_players p, lnk_p_pl lnk where lnk.PlayerID = p.PLayerID and lnk.PPLLinkID = " & mAttackerID
	ors.Open  strSQL, oconn 
	if not (ors.EOF and ors.BOF) then
		mAttackerName=ors.Fields(0).value
 	end if
	ors.Close
	
strsql="Select Playerladdername from tbl_Playerladders where PlayerLadderID=" & mLadderID
ors.open strsql, oconn

if not (bSysAdmin OR (session("uName") = mDefenderName) OR (session("uName") = mAttackerName) OR IsPlayerLadderAdmin(ors("PlayerLadderName"))) Then
	oRS.Close
	oConn.Close 
	Set oRS = Nothing
	Set oConn = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
end if
ors.close
mDiD = Request.form("mDefenderID")
mAID = request.form("mWinnerID")
mMatchDate=Request.Form("matchdate")
mMap1=Request.Form("Map1")
mMap1DefScore=Request.Form("Map1DefScore")
mMap1AttScore=Request.Form("Map1AttScore")
mMap1Forfeit=Request.Form("Map1Forfeit")

dScore=0
aScore=0
m1c=0

'--------------------------
if len(mmap1defscore) <> len(mmap1attscore) then
	if len(mmap1defscore) > len(mmap1attscore) then
		mmap1attscore="0" & mmap1attscore
	else
		mmap1defscore="0" & mmap1defscore
	end if
end if
if len(mmap2defscore) <> len(mmap2attscore) then
	if len(mmap2defscore) > len(mmap2attscore) then
		mmap2attscore="0" & mmap2attscore
	else
		mmap2defscore="0" & mmap2defscore
	end if
end if
if len(mmap3defscore) <> len(mmap3attscore) then
	if len(mmap3defscore) > len(mmap3attscore) then
		mmap3attscore="0" & mmap3attscore
	else
		mmap3defscore="0" & mmap3defscore
	end if
end if
'Response.Write "Map 1:" & mmap1attscore & " to " & mmap1defscore & "<br>"
'--------------------------

if mMap1DefScore > mMap1AttScore then
	mWinner1 = mDefenderName
	dScore = dScore +1
	if m1c=0 then
		m1c = 1
	end if
elseif mMap1DefScore < mMap1AttScore then
	mWinner1 = mAttackerName
	aScore=aScore +1
	if m1c=0 then
		m1c = 1
	end if
else
	mWinner1= "None"
end if

if mMap1Forfeit ="Defender" then
	Map1Forfeit = mAttackerName & " wins by Forfeit"
	mWinner1= mAttackerName
	mMap1Forfeit=1
	if m1c=0 then
		aScore=aScore +1
		m1c=1
	end if
end if
if mMap1Forfeit ="Attacker" then
	Map1Forfeit = mDefenderName & " wins by Forfeit"
	mWinner1 = mDefenderName
	mMap1Forfeit=1
	if m1c=0 then
		dScore = dScore +1
		m1c=1
	end if
end if

if dScore > 1 then 
	dScore = 1
end if
if aScore > 1 then
	aScore = 1
end if
if dScore > aScore then
	mWinner = mDefenderName
	mWinnerID = mDefenderID
	mLoserID = mAttackerID
	mLoser=mAttackerName
	mMatchWinnerIsDefender = True
	pScore = dscore & "-" & ascore
else
	mWinnerID = mAttackerID
	mLoserID = mDefenderID
	mWinner = mAttackerName
	mLoser = mDefenderName
	mMatchWinnerIsDefender = False	
	pScore = ascore & "-" & dscore
end if
%>
<TABLE BORDER=0 CELLSpACING=0 CELLPADDING=0 BGCOLOR="#444444">
<TR><TD>
<table align=center border=0 cellspacing=1 cellpadding=2 width="100%">
<%
if mWinner = Request("Reporter") then
	%>
	<tr bgcolor=<%=bgctwo%>><td align=center><%=Server.HTMLEncode (mDefenderName)%> vs. <%=Server.HTMLEncode (mAttackerName)%></td></tr>
	<tr bgcolor=<%=bgcone%>><td align=center>You are not permitted to report as the winner of the match</td></tr>
	<tr bgcolor=<%=bgctwo%>><td align=center><a href="MatchReportLoss.asp?matchid=<%=matchid%>&teamid=<%=teamid%>">Click here to return</a></td></tr></table>
<%
else
	if mMap1Forfeit <> "1" then
		mMap1Forfeit=0
	end if
	%>
	<tr bgcolor=#000000><TH colspan=2><%=Server.HTMLEncode (mDefenderName)%> vs. <%=Server.HTMLEncode (mAttackerName)%> </TH></tr>
	<tr bgcolor=<%=bgcone%> height=30><td align=center colspan=2><b>&nbsp;<%=Server.HTMLEncode (mWinner)%> wins the match <%=pScore%></b></td></tr>
	<tr bgcolor=<%=bgctwo%>><td colspan=2>&nbsp;<%=Server.HTMLEncode (mMap1)%> winner: <%=Server.HTMLEncode (mWinner1)%></td></tr>
	<tr bgcolor=<%=bgcone%>><td>&nbsp;<%=Server.HTMLEncode (mDefenderName)%> Score:</td><td><%=mMap1DefScore%></td></tr>
	<tr bgcolor=<%=bgctwo%>><td>&nbsp;<%=Server.HTMLEncode (mAttackerName)%> Score: </td><td><%=mMap1AttScore%></td></tr>
	<%
	if mMap1Forfeit="1" then %>
	<tr bgcolor=<%=bgcone%>><td colspan=2>&nbsp;<%=Map1Forfeit%></td></tr>
	<% 
	end if 
	%>
	<tr bgcolor=<%=bgcone%> valign=center><td colspan=2 align=center><input type=button value="Back" onclick="window.location.href='MatchReportLoss.asp?matchid=<%=Request.Form("matchid")%>&teamid=<%=request.form("teamid")%>';" class=bright>
	&nbsp;&nbsp;
	<script language="JavaScript">
	urlreport="saveitem.asp?SaveType=PlayerReportMatch&matchid=<%=matchid%>&matchwinner=<%=server.urlencode(mwinner)%>&matchloser=<%=server.urlencode(mLoser)%>&map1=<%=server.urlencode(mMap1)%>&map2=<%=server.urlencode(mMap2)%>&map3=<%=server.urlencode(mMap3)%>&map1defenderscore=<%=mMap1DefScore%>&map2defenderscore=<%=mMap2DefScore%>&map3defenderscore=<%=mMap3DefScore%>&map1attackerscore=<%=mMap1AttScore%>&map2attackerscore=<%=mMap2AttScore%>&map3attackerscore=<%=mMap3AttScore%>&map1ot=<%=mMap1OTWin%>&map2ot=<%=mMap2OTWIn%>&map3ot=<%=mMap3OTWin%>&map1forfeit=<%=mMap1Forfeit%>&map2Forfeit=<%=mMap2Forfeit%>&map3forfeit=<%=mMap3Forfeit%>&matchdate=<%=mMatchDate%>&matchwinnerdefending=<%=mMatchWinnerIsDefender%>&ladderid=<%=mLadderID%>&matchwinnerid=<%=mwinnerID%>&matchloserid=<%=mloserid%>";
	</script>
	<input type=button name=saveresults value="Save Results" class=bright onclick="window.location.href=urlreport;">
	</td></tr>
<%
end if
%>
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
