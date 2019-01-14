<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Voting Ballot"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

DIm bLadderType

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim BallotID, BallotName, BallotType, qrs, rrs, isActive, bLadder
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
BallotID = Request.QueryString("BallotId")
If Len(BallotID) = 0 OR Not(IsNUmeric(BallotID)) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If

strsql="select top 1 * from tbl_ballot where BallotID = '" & BallotID & "'"
ors.Open strsql, oconn
bLadder = 0

if not (ors.EOF and ors.BOF) then
	ballotid = ors.Fields(0).Value
	ballotname = ors.Fields(1).Value
	ballottype = ors.Fields("Type").Value
	isActive = ors.Fields("isactive").Value 
	bLadder = ors.Fields("ladderid").Value 
	bLadderType = oRs.Fields("LadderType").Value
end if
'response.write bLadder & "<br><br>"
ors.NextRecordset 

If isActive = 0 Then
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/ballot/default.asp?error=2"
end if

select case BallotType
	case 1
		if not(IsAnyTeamFounder()) then
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "/ballot/default.asp?error=3"
		end if
		'bLadder = Request.Form("bLadder")
		'strsql = "select count(*) from lnk_T_L inner join tbl_teams on tbl_teams.teamid=lnk_T_L.teamid where ladderid=" & bLadder & " and teamadmin='" & Session("uName") & "'"
		if bLadderType = "T" Then 
			strsql = "select count(TLLinkID) from lnk_T_L inner join tbl_teams on tbl_teams.teamid=lnk_T_L.teamid inner join tbl_players on tbl_players.playerid=tbl_teams.teamfounderid where ladderid=" & bLadder & " and playerhandle='" & Session("uName") & "' and isactive=1"
			'response.write strsql
			ors.Open strsql, oconn
			if ors.Fields(0).Value=0 then
				oConn.Close 
				Set oConn = Nothing
				Set oRs = Nothing
				Response.Clear
				Response.Redirect "/ballot/default.asp?error=3"
			end if
			ors.Close
		ElseIf bLadderType = "L" Then
			strsql = "select count(lnkLeagueTeamID) from lnk_league_team inner join tbl_teams on tbl_teams.teamid=lnk_league_team.teamid inner join tbl_players on tbl_players.playerid=tbl_teams.teamfounderid where LeagueID=" & bLadder & " and playerhandle='" & Session("uName") & "' and active=1"
			'response.write strsql
			ors.Open strsql, oconn
			if ors.Fields(0).Value=0 then
				oConn.Close 
				Set oConn = Nothing
				Set oRs = Nothing
				Response.Clear
				Response.Redirect "/ballot/default.asp?error=4"
			end if
			ors.Close 			
		End If
	case else
		if not(Session("LoggedIn")) then
			oConn.Close 
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "/errorpage.asp?error=2"
		end if
end select

strsql = "select voterid from tbl_votes v, tbl_questions q where q.ballotid = " & ballotid & " AND q.qid = v.qid AND v.voterid = (select TOP 1 PlayerID from tbl_players where playerhandle = '" & CheckString(session("Uname")) & "')"
ors.Open strsql, oconn
if not(ors.EOF and ors.BOF) then
	oRs.Close
	oConn.Close 
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/ballot/default.asp?error=1"
end if
ors.Close 

set qrs = server.CreateObject("ADODB.recordset")
set rrs = server.CreateObject("ADODB.recordset")
Call ContentStart(BallotName & " Ballot")

Dim q
%>
<form method=post action=saveballotvote.asp id=form1 name=form1>
    <table width="90%" border="0">
<tr><td>This ia a ballot used for different options that will be voted on for Teamwarfare. Choose the option you wish to vote on
or choose to abstain to not make a vote on that question.</TD></TR>
<TR><TD><HR class="forum"></TD></TR>
<TR><TD>
<%
q=0
	strsql="select * from tbl_questions where ballotid=" & ballotid & " order by questionnum"
	qrs.Open strsql, oconn
	if not (qrs.EOF and qrs.BOF) then
		do while not qrs.eof 
			q=q+1
			Response.Write "<TR><TD>"
			Response.Write "Question " & qrs.Fields(2).Value & ": <u>" & qrs.Fields(3).Value & "</u></TD></TR>"& vbcrlf
			strsql= "select * from tbl_responses where QID=" & qrs.Fields(0).Value & " order by RVal"
			rrs.Open strSQL, oconn
			if not (rrs.EOF and rrs.BOF) then
				do while not rrs.EOF
					Response.Write "<TR><TD><p class=small><input name=q_" & q & "_vote type=radio class=borderless value=""" & rrs.Fields(2).Value & """>"& rrs.Fields(2).Value & "</b>: " & rrs.Fields(3).Value & "</P></TR></TD>" & Vbcrlf
					rrs.MoveNext
				loop
			end if
			Response.Write "<TR><TD><p class=small><input type=radio class=borderless name=q_" & q & "_vote value=""""> Abstain from this question.<input type=hidden name=QID_" & q & " value=" & qrs.Fields(0).Value & "></TD></TR>"
			Response.Write "<TR><TD><hr class=forum></TD></TR>"
			rrs.NextRecordset 
			qrs.MoveNext
		loop
	end if
	qrs.Close 
%>
<TR><TD><input type=hidden value=<%=q%> name=QCount>
<TR><TD><input type=hidden value="<%=ballotid%>" name=BallotID>
<center><input type=submit value="Vote Now"></center>
</td></tr></table>
</form>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
Set rrS = Nothing
Set qrs = Nothing
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>