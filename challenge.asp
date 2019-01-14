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

Dim strOpponent, strLadder, strTeam, strStatus
Dim strEnemyName, strResult

strOpponent = Request.QueryString("Opponent")
strLadder = Request.QueryString("Ladder")
strTeam = Request.QueryString("Team")

If Not(bSysAdmin Or IsTeamFounder(strTeam) Or IsTeamCaptain(strTeam, strLadder) Or IsLadderAdmin(strLadder)) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If	

strsql= "SELECT lnk_T_L.Status, tbl_Teams.TeamName, tbl_Ladders.LadderName "
strsql= strsql & "FROM (lnk_T_L INNER JOIN tbl_Teams ON lnk_T_L.TeamID = tbl_Teams.TeamID) INNER JOIN tbl_Ladders ON lnk_T_L.LadderID = tbl_Ladders.LadderID "
strsql= strsql & "WHERE (((tbl_Teams.TeamName)='" & CheckString(strOpponent) & "') AND lnk_t_l.isactive=1 and ((tbl_Ladders.LadderName)='" & CheckString(strLadder) & "'))"
ors.open strsql, oconn
if not (ors.eof and ors.bof) then
	strStatus=ors.fields(0).value
Else
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
ors.close

If strStatus <> "Available" and left(strStatus, 8) <> "Defeated" then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=4"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call Content33BoxStart("Summary of team: " & Server.HTMLEncode(strOpponent))%>
	<table width="780" border="0" cellspacing="0" cellpadding="0" background="">
	<tr valign=top>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	<td width="250">

	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" WIDTH="245" ALIGN="CENTER">
	<TR><TD>
	<table width=100% align=center border=0 cellpadding=2 cellspacing=1>
	<TR BGCOLOR="#000000">
		<TH>Current Roster</TH>
	</TR>
<%
	strSQL="SELECT tbl_Players.PlayerHandle, tbl_Teams.TeamName, tbl_Ladders.LadderName "
	strSQL= strSQL & "FROM tbl_Ladders INNER JOIN (tbl_Players INNER JOIN "
	strSQL= strSQL & "(tbl_Teams INNER JOIN (lnk_T_L INNER JOIN lnk_T_P_L ON lnk_T_L.TLLinkID = lnk_T_P_L.TLLinkID) ON "
	strSQL= strSQL & "tbl_Teams.TeamID = lnk_T_L.TeamID) ON tbl_Players.PlayerID = lnk_T_P_L.PlayerID) ON tbl_Ladders.LadderID = lnk_T_L.LadderID "
	strSQL= strSQL & "WHERE (((tbl_Teams.TeamName)='" & CheckString(strOpponent) & "') AND ((tbl_Ladders.LadderName)='" & CheckString(strLadder) & "')) order by tbl_Players.PlayerHandle;"
	oRs.Open strSQL, oConn
	bgc=bgcone
	if not (ors.EOF and ors.BOF) then 
		do while not ors.EOF
			Response.Write "<tr bgcolor=" & bgc & "><td align=center><a href=viewplayer.asp?Player=" & server.urlencode(ors.Fields(0).Value) & ">" & Server.HTMLEncode (ors.Fields("PlayerHandle").Value) & "</td></tr>"
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.MoveNext
		loop
	end if
	oRS.NextRecordset 
%>
	</table>
	</TD></TR>
	</TABLE>
	
	</td>
	<td><img src="/images/spacer.gif" width="10" height="1"></td>
	<td width="510">
<% If False Then 
	' Removed Sunday, June 4th, 2006 by Triston -- Causing timeoutes
	%>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" WIDTH="505" ALIGN="CENTER">
	<TR><TD>
	<table width=100% align=center border=0 cellpadding=2 cellspacing=1>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=4>Recent History</TH>
	</TR>
	<TR BGCOLOR="#000000">
		<TH>Ladder</TH>
		<TH>Opponent</TH>
		<TH>Result</TH>
		<TH>Date</TH>
	</TR>	
	<%
	strSQL = "SELECT TOP 2 * FROM vHistory WHERE MatchForfeit = 0 AND (WinnerName = '" & CheckString(strOpponent) & "' OR LoserName='" & CheckString(strOpponent) & "') ORDER BY MatchDate Desc"
	oRS.Open strsql, oconn
	if not (oRS.EOF and oRS.BOF) then
		do while not oRS.EOF
			If oRS.Fields("WinnerName") = strOpponent Then
				strEnemyName = oRS.Fields("LoserName").Value 
				strResult = "Win"
			Else
				strEnemyName = oRS.Fields("WinnerName").Value 
				strResult = "Loss"
			End If
			%>
			<tr bgcolor=<%=bgc%>><td>&nbsp;<a href=viewladder.asp?ladder=<%=server.urlencode(oRS.Fields("LadderName").Value )%>><%=Server.HTMLEncode(oRS.Fields("LadderName").Value)%></a></td>
			<td><a href=viewteam.asp?team=<%=server.urlencode(strEnemyName)%>><%=Server.HTMLEncode(strEnemyName)%></a></td>
			<td><%=strResult%></td>
			<td align=right><%=oRS.Fields("MatchDate").Value%>&nbsp;</td></tr>
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			oRS.MoveNext
		loop
	end if
	oRS.Close 
	%>
	<tr bgcolor="#000000">
	</TABLE>
	</TD></TR>
	</TABLE>
	<BR><BR>
<% End If %>
	
	<% 
	dim strMapConfiguration, i
	strSQL = "SELECT MapConfiguration FROM tbl_ladders WHERE LadderName = '" & CheckString(strLadder) & "'"
	ors.Open strSQL, oConn
	If Not( ors.EOF and ors.BOF) Then
		strMapConfiguration = ors.Fields("MapConfiguration").Value 
	End If
	ors.NextRecordset
	%> 
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN="CENTER">
	<TR><TD>
	<table width=100% align=center border=0 cellpadding=2 cellspacing=1>
	<FORM NAME="frmChallenge" ID="frmChallenge" METHOD="GET" ACTION="SaveItem.asp">
	<INPUT TYPE="HIDDEN" name="SaveType" value="challenge">
	<input type="hidden" name="ladder" value="<%=strLadder%>">
	<input type="hidden" name="team" value="<%=Server.HTMLEncode(strTeam & "")%>">
	<input type="hidden" name="opponent" value="<%=Server.HTMLEncode(strOpponent & "")%>">
	<TR bgcolor="#000000">
		<TH COLSPAN=2>Confirm Challenge</TH>
	</TR>
	<TR bgcolor="<%=bgcone%>">
		<TD align="right"><b>Ladder:</b></TD>
		<TD><%=strLadder%></TD>
	</TR>
	<TR bgcolor="<%=bgcone%>">
		<TD align="right"><b>Your Team:</b></TD>
		<TD><%=strTeam%></TD>
	</TR>
	<TR bgcolor="<%=bgcone%>">
		<TD align="right"><b>Opponent Team:</b></TD>
		<TD><%=strOpponent%></TD>
	</TR>
	<%
	For i = 1 to Len(strMapConfiguration) 
		If mid(strMapConfiguration, i, 1) = "C" Then
			%>
			<TR bgcolor="<%=bgctwo%>">
				<td align=right><b>Choose map <%=i%>:</b></TD>
				<td><SELECT NAME="Map<%=i%>" class="bright">
					<%
					strSQL = "GetLadderMapList '" & CheckString(strLadder) & "', '" & i & "'"
					ors.Open strSQL, oconn
					If (ors.State = 1) Then
						If not(ors.EOF and ors.BOF) Then
							Do While Not(ors.EOF)
								Response.Write "<OPTION VALUE=""" & ors.Fields("MapName").Value & """>" & ors.Fields("MapName").Value  & "</OPTION>" & vbCrLf
								ors.MoveNext
							Loop
						End If
						ors.Close
					End If
					%></select></td>
			</tr>	
			<%
		End If
	Next
	%>
	<TR bgcolor="#000000">
		<TD COLSPAN=2 ALIGN=CENTER><INPUT TYPE="SUBMIT" VALUE="Submit Challenge"></TD>
	</TR>
	<TR bgcolor="#000000">
		<TD COLSPAN=2 ALIGN=CENTER><a href="TeamLadderAdmin.asp?ladder=<%=server.urlencode(strLadder)%>&team=<%=server.urlencode(strTeam)%>"  onmouseover="(window.status='Return to admin page and abort challenge'); return true" onmouseout="(window.status='TWL'); return true">Click here to abort challenge</a></TD>
	</TR>
	</TABLE>
	</TD></TR></TABLE>
	</td>
	<td width=5><img src="/images/spacer.gif" width="5" height="1"></td>
	</tr>
	</table>
<% Call Content33BoxEnd() %>

<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>