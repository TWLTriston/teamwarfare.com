<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Administrative Menu"

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

If Not(bSysAdmin Or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Redirect "/errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Pending Forfeits") %>
				<table class="cssBordered" width="60%" align="center">
	<TR BGCOLOR="#000000">
		<TH>Team Based Ladders</TH>
	</TR>
	<%
		
	If bSysAdmin Then
		strSQL = " SELECT LadderName, LadderID, ForFeits = ( select count(*) from tbl_matches where matchawaitingforfeit=1 and left(forfeitreason, 5) <> 'Admin' AND matchladderid=l.LadderID) "
		strSQL = strSQL & " FROM tbl_ladders l WHERE LadderActive = 1 ORDER BY LadderName"	
	Else
		strSQL = " SELECT LadderName, l.LadderID, ForFeits = ( select count(*) from tbl_matches where matchawaitingforfeit=1 and left(forfeitreason, 5) <> 'Admin' AND matchladderid=l.LadderID) "
		strSQL = strSQL & " FROM tbl_ladders l, lnk_l_a lnk WHERE lnk.LadderID = l.LadderID AND lnk.PlayerID='" & Session("PlayerID") & "' AND LadderActive = 1 ORDER BY LadderName"	
	End If
	ors.open strSQL, oconn
	if not (ors.eof and ors.bof) then
		Do While Not(oRS.EOF) 
			if ors.fields("ForFeits").value > 0 then
				bgc = bgcone
				Response.Write "<tr bgcolor=" & bgc & "><td align=center>"
				Response.Write "<a href=adminops.asp?rAdmin=Forfeit&ladderid=" & ors.Fields("LadderID").Value & "&laddername=" & server.urlencode(ors.Fields("LadderName").Value) & ">" & ors.fields("ForFeits").value & " " & ors.fields("LadderName").value & " forfeits pending</a></td></tr>"
				If bgc = bgcOne Then
					bgc = bgcTwo
				Else
					bgc = bgcone
				End If
			end if
			ors.movenext
		loop
	end if
	ors.close
	%>
	</TABLE>
<BR><BR>
				<table class="cssBordered" width="60%" align="center">
	<TR BGCOLOR="#000000">
		<TH>Player Based Ladders</TH>
	</TR>
	<%
	If bSysAdmin Then
		strSQL = "SELECT l.PlayerLadderID, PlayerLaddername, ForFeits = ( select count(PlayerMatchID) from vPlayerMatches where matchawaitingforfeit=1 and left(forfeitreason, 5) <> 'Admin' AND matchladderid=l.PlayerLadderID) "
		strSQL = strSQL & " FROM tbl_playerLadders l WHERE Active > 0 ORDER BY PlayerLaddername"
	Else
		strSQL = "SELECT l.PlayerLadderID, PlayerLaddername, ForFeits = ( select count(PlayerMatchID) from vPlayerMatches where matchawaitingforfeit=1 and left(forfeitreason, 5) <> 'Admin' AND matchladderid=l.PlayerLadderID) "
		strSQL = strSQL & " FROM tbl_playerLadders l, lnk_pl_a lnk WHERE lnk.PlayerLadderID = l.PlayerLadderID AND lnk.PlayerID='" & Session("PlayerID") & "' AND Active > 0 ORDER BY PlayerLaddername"
	End If
	ors.open strSQL, oconn
	if not (ors.eof and ors.bof) then
		bgc=bgcone
		Do While Not ors.eof
			if ors.fields("ForFeits").value > 0 then
				Response.Write "<tr bgcolor=" & bgc & ">"
				Response.Write "<td align=center>"
				Response.Write "<a href=adminops.asp?rAdmin=PForfeit&ladderid=" & ors.Fields("PlayerLadderID").Value 
				Response.Write "&laddername=" & server.urlencode(ors.Fields("PlayerLadderName").Value)
				Response.Write ">" & ors.fields("ForFeits").value & " "
				Response.write ors.fields("PlayerLadderName").value & " forfeits pending</a></td></tr>"
				if bgc=bgcone then
					bgc=bgctwo
				else
					bgc=bgcone
				End If
			end if
			ors.movenext
		loop
	end if
	ors.close
	%>
	</TABLE>
	</TD>
</TR>
</TABLE>
<% 
Call ContentEnd()
Call ContentStart("Admin Options") 
%>
	<form name=frmAdminMenu action=adminops.asp method=post>
				<table class="cssBordered" width="100%">
	<tr bgcolor="#000000"><tH align=center colspan=2>Admin Menu</td></tr>
	<tr bgcolor=<%=bgctwo%>><td align=right WIDTH=15><input type=radio class=borderless name=rAdmin value=Match></td><td WIDTH="385">Match Admin</td></tr>
	<tr bgcolor=<%=bgcone%>><td align=right WIDTH=15><input type=radio class=borderless name=rAdmin value=Forfeit></td><td>Forfeit Admin</td></tr>
	<tr bgcolor=<%=bgctwo%>><td align=right WIDTH=15><input type=radio class=borderless name=rAdmin value=PMatch></td><td>Match (PlayerLadder) Admin</td></tr>
	<tr bgcolor=<%=bgcone%>><td align=right WIDTH=15><input type=radio class=borderless name=rAdmin value=PForfeit></td><td>Forfeit (PlayerLadder) Admin</td></tr>
	<tr bgcolor=<%=bgctwo%>><td align=right WIDTH=15><input type=radio class=borderless name=rAdmin value=Ladder></td><td>Ladder Admin</td></tr>
	<tr bgcolor=<%=bgcone%>><td align=right WIDTH=15><input type=radio class=borderless name=rAdmin value=Rank></td><td>Rankings Admin</td></tr>
	<tr bgcolor=<%=bgctwo%>><td align=right WIDTH=15><input type=radio class=borderless name=rAdmin value=History></td><td>History Admin</td></tr>
	<%
	if bSysAdmin then
	%>
	<tr bgcolor=<%=bgcone%>><td align=right WIDTH=15><input type=radio class=borderless name=rAdmin value=Player></td><td>Player Admin</td></tr>
	<tr bgcolor=<%=bgctwo%>><td align=right WIDTH=15><input type=radio class=borderless name=rAdmin value=Team></td><td>Team Admin</td></tr>
	<%
	end if
	%>
	<tr bgcolor=<%=bgctwo%>><td align=center colspan=2><input class=bright value='Enter Administration Section' type=submit name=submit1></td></tr>

	</table>
	</form>
	<br />
	<br />
	<a href="/request/AdmNames.asp">SITE SUPPORT ONLY!</a>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>