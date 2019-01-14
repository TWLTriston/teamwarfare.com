<% Option Explicit %>
<%
Response.Buffer = True
Response.Expires = -1440
Dim strPageTitle

strPageTitle = "TWL: Ladder Admin Assignments"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Redirect "/errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("Current Admins")

Dim strGameName, strLadderName, blnLadderActive, blnLadderLocked, intLadderID
strSQL = "SELECT tbl_games.GameName, tbl_ladders.LadderName, tbl_ladders.LadderID, tbl_ladders.LadderActive, tbl_ladders.LadderLocked FROM tbl_games LEFT OUTER JOIN tbl_ladders ON tbl_games.GameID = tbl_ladders.GameID WHERE tbl_games.GameID > 0 AND LadderShown = 1 ORDER BY GameName, LadderLocked ASC, LadderActive DESC, LadderName ASC"
oRS.Open strSQL, oConn, 3, 3
If Not(oRS.EOF AND oRS.BOF) Then
	strGameName = oRS.Fields("GameName").Value
	Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" bgcolor=""#444444"" width=""50%"" align=""center""><tr><td>"
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""2"" width=""100%"">"
	Response.Write "<tr><th bgcolor=""#000000"">Jump to Game</th></tr>"
	Response.Write "<tr><td align=""center"" bgcolor=""" & bgcone & """><a href=""#admin"">Add an admin</a></td></tr>"
	Response.Write "<tr><td bgcolor=""" & bgcone & """><a href=""#" & Replace(strGameName, " ", "") & """>" & strGameName & "</a></td></tr>"
	bgc = bgctwo
	Do While Not(oRS.EOF)
		If strGameName <> oRs.Fields("GameName").Value Then
			strGameName = oRS.Fields("GameName").Value
			Response.Write "<tr><td bgcolor=""" & bgcone & """><a href=""#" & Replace(strGameName, " ", "") & """>" & strGameName & "</a></td></tr>"
			If bgc = bgcone THen
				bgc = bgctwo
			Else
				bgc = bgcone
			End If
		End If
		oRS.MoveNext
	Loop
	oRS.MoveFirst
	Response.Write "</table></td></tr></table><br>"

	strGameName = oRS.Fields("GameName").Value
	Response.Write "<a name=""" & Replace(strGameName, " ", "") & """ /><table border=""0"" cellspacing=""0"" cellpadding=""0"" bgcolor=""#444444"" width=""90%"" align=""center""><tr><td>"
	Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""2"" width=""100%"">"
	Response.Write "<tr><th bgcolor=""#000000"" colspan=""3"">" & strGameName & " <a href=""#top"">Return to top</a></th></tr>"
	Response.Write "<tr><th width=""90%"" bgcolor=""#000000"">Ladder Name</th><th width=""75"" bgcolor=""#000000"">Locked</th><th width=""75"" bgcolor=""#000000"">Active</th></tr>"
	Do While Not(oRS.EOF)
		If strGameName <> oRs.Fields("GameName").Value Then
			strGameName = oRS.Fields("GameName").Value
			Response.Write "</table></td></tr></table><br><br><br><a name=""" & Replace(strGameName, " ", "") & """ />"
			Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" bgcolor=""#444444"" width=""90%"" align=""center""><tr><td>"
			Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""2"" width=""100%"">"
			Response.Write "<tr><th bgcolor=""#000000"" colspan=""3"">" & strGameName & " <a href=""#top"">Return to top</a></th></tr>"
			Response.Write "<tr><th width=""90%"" bgcolor=""#000000"">Ladder Name</th><th width=""75"" bgcolor=""#000000"">Locked</th><th width=""75"" bgcolor=""#000000"">Active</th></tr>"
		End If
		strLadderName = oRS.Fields("LadderName").Value
		
		If IsNull(oRS.Fields("LadderName").Value) Then
			Response.Write "<tr><td colspan=""3"" bgcolor=""#000000""><i>No ladders for this game</td></tr>"
		Else
			intLadderID = oRS.Fields("LadderID").Value
			blnLadderLocked = cBool(oRs.Fields("LadderLocked").Value)
			blnLadderActive = cBool(oRs.Fields("LadderActive").Value)
			Response.Write "<tr><td bgcolor=""" & bgcone & """><a name=""" & intLadderID & """ /><a href=""viewladder.asp?ladder=" & Server.URLEncode(strLadderName) & """>" & strLadderName & "</a></td>"
			If blnLadderLocked Then
				Response.Write "<td align=""center"" bgcolor=""" & bgcone & """><font color=""red"">Yes</font></td>"
			Else
				Response.Write "<td align=""center"" bgcolor=""" & bgcone & """>No</td>"
			End If
			If blnLadderActive Then
				Response.Write "<td align=""center"" bgcolor=""" & bgcone & """>Yes</td>"
			Else
				Response.Write "<td align=""center"" bgcolor=""" & bgcone & """><font color=""red"">No</font></td>"
			End If
			
			Response.Write "<tr><td colspan=""3"" align=""center"" bgcolor=""#000000"">"
			strSQL = "SELECT tbl_players.PlayerHandle, lnk_l_a.PrimaryAdmin, lnk_l_a.LALinkID, lnk_l_a.LadderID FROM lnk_l_a INNER JOIN tbl_players ON lnk_l_a.PlayerID = tbl_players.PlayerID WHERE lnk_l_a.LadderID = '" & intLadderID & "' ORDER BY PrimaryAdmin DESC, PlayerHandle ASC "
			oRS2.Open strSQL, oConn
			If Not(oRS2.EOF AND oRS2.BOF) Then
				Response.Write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""50%""><tr><td bgcolor=""#444444"">"
				Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""2"" width=""100%"">"
				Response.Write "<tr><th bgcolor=""#000000"">Login Name</th><th bgcolor=""#000000"" width=""75"">Primary Admin</th><th width=""50"" bgcolor=""#000000"">Delete</th><th width=""100"" bgcolor=""#000000"">Make Primary</th></tr>"
				bgc = bgcone
				Do While Not(oRS2.EOF)
					Response.Write "<tr><td bgcolor=""" & bgc & """><a href=""viewplayer.asp?player=" & Server.URLEncode(oRS2.Fields("PlayerHandle").Value) & """>" & oRS2.Fields("PlayerHandle").Value & "</a></td>"
					If oRS2.Fields("PrimaryAdmin").Value Then
						Response.Write "<td bgcolor=""" & bgc & """ align=""center""><font color=""red"">Primary</font></td>"
					Else
						Response.Write "<td bgcolor=""" & bgc & """>&nbsp;</td>"
					End If
					Response.Write "<td bgcolor=""" & bgc & """ align=""center""><a href=""javascript:DeleteAdmin('" & Replace(oRS2.Fields("PlayerHandle").Value, "'", "\'") & "', '" & Replace(strLadderName, "'", "\'") & "', '" & oRS2.Fields("LALinkID").Value & "', '" & oRS2.Fields("LadderID").Value & "')"">Delete</a></td>"
					Response.Write "<td bgcolor=""" & bgc & """ align=""center""><a href=""javascript:MakePrimary('" & Replace(oRS2.Fields("PlayerHandle").Value, "'", "\'") & "', '" & Replace(strLadderName, "'", "\'") & "', '" & oRS2.Fields("LALinkID").Value & "', '" & oRS2.Fields("LadderID").Value & "')"">Make Primary</a></td>"
					Response.Write "</tr>"
					If bgc = bgcone Then
						bgc = bgctwo
					Else
						bgc = bgcone
					End If
					oRS2.MoveNext
				Loop
				Response.Write "</table></td></tr></table><br>"
			Else
				Response.Write "<i>No admins are assigned for this ladder...</i>"			
			End If
			oRS2.NextRecordSet
			
			Response.Write "</td></tr>"
			
		End If
		oRS.MoveNext
	Loop
	Response.Write "</table></td></tr></table>"
End If
oRS.NextRecordSet

Call ContentEnd()

Call Content2BoxStart("Team Ladder Admins")
%>
<script language="javascript">
<!--
	function MakePrimary(strPlayerName, strLadderName, RelevantID, LadderID) {
		if (confirm('Yo dude, you really wanna make this fool (' + strPlayerName + ') the PRIMARY admin for ' + strLadderName + '? Is he ghey enough?')) {
			window.location.href = "saveitem.asp?SaveType=PrimaryAdmin&LALinkID=" + RelevantID + "&LadderID=" + LadderID;
		} else {
			alert("I didnt think so");
		}
	}
	
	function DeleteAdmin(strPlayerName, strLadderName, RelevantID, LadderID) {
		if (confirm('Faggotry begins... removing admin... ' + strPlayerName + ' as admin on ' + strLadderName + '. Confirm.')) {
			window.location.href = "saveitem.asp?SaveType=DeleteAdmin&LALinkID=" + RelevantID + "&LadderID=" + LadderID;
		} else {
			alert("Admin is safe for now.");
		}
	}
//-->
</script>
<a name="admin">
<table width=780 border="0" cellspacing="0" cellpadding="0" BACKGROUND="">
<tr>
<td><img src="/images/spacer.gif" width="5" height="1"></td>
<td width=380>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" WIDTH="375" ALIGN=CENTER>
	<TR><TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
	<% if request("player") <> "" Then %>
		<form action=saveItem.asp method=post id=form4 name=form4>
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2>Select Player to Admin</TH>
		</TR>
		<tr bgcolor=<%=bgcone%>><td>&nbsp;&nbsp;Player: </td><td><select name=player class=bright style="width:200">
		<%
		strsql="select playerhandle, tbl_players.playerid from tbl_players WHERE playerHandle like '%" & CheckString(SearchString(request("player"))) & "%' order by playerhandle"
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			do while not ors.EOF
				Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value) & "</option>"
				ors.MoveNext 
			loop
		end if
		ors.NextRecordSet 
		%>
		</td></tr>
		<tr bgcolor=<%=bgctwo%> height=30><td>&nbsp;&nbsp;Ladder</td><td><select name=ladder class=bright style="width:300">
		<%
		strsql="select LadderName, LadderID from tbl_ladders WHERE LadderShown = 1 order by LadderName"
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			do while not ors.EOF
				Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value) & "</option>"
				ors.MoveNext
			loop
		end if
		ors.NextRecordSet
		%>
	<% Else %>
		<form action=assignadmin.asp#admin method=get id=form2 name=form2>
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2>Search for Player</TH>
		</TR>
		<tr bgcolor=<%=bgcone%> height=30><td>&nbsp;&nbsp;Player Name: </td><td><input type=text name=player class=bright style="width:200"></td></tr>
	<% End IF %>
	<tr bgcolor=<%=bgcone%>><td colspan=2 align=center><input type=submit value='Make it So' class=bright id=submit5 name=submit5><input type=hidden name=SaveType value=SetLadderAdmin></td></tr>
	</form>
	</TABLE>
	</TD>
</TR>
</TABLE>
	
</td>
<td><img src="/images/spacer.gif" width="10" height="1"></td>
<td width=379>

	
</td>
<td><img src="/images/spacer.gif" width="5" height="1"></td>
</tr>
</table>
<%
Call Content2BoxEnd()
Call Content2BoxStart("Player Ladder Admins")
%>
<table width=780 border="0" cellspacing="0" cellpadding="0" BACKGROUND="">
<tr>
<td><img src="/images/spacer.gif" width="5" height="1"></td>
<td width=380>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" WIDTH="375" ALIGN=CENTER>
	<TR><TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
	<% if request("player") <> "" Then %>
		<form action=saveItem.asp method=post id=form4 name=form4>
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2>Select Player to Admin</TH>
		</TR>
		<tr bgcolor=<%=bgcone%>><td>&nbsp;&nbsp;Player: </td><td><select name=player class=bright style="width:200">
		<%
		strsql="select playerhandle, tbl_players.playerid from tbl_players WHERE playerHandle like '%" & CheckString(SearchString(request("player"))) & "%' order by playerhandle"
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			do while not ors.EOF
				Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value) & "</option>"
				ors.MoveNext 
			loop
		end if
		ors.NextRecordSet 
		%>
		</td></tr>
		<tr bgcolor=<%=bgctwo%> height=30><td>&nbsp;&nbsp;Ladder</td><td><select name=ladder class=bright style="width:300">
		<%
		strsql="select playerLadderName, playerLadderID from tbl_playerladders order by playerLadderName"
		ors.Open strsql, oconn
		if not (ors.EOF and ors.BOF) then
			do while not ors.EOF
				Response.Write "<option value=" & ors.Fields(1).Value & ">" & Server.HTMLEncode(ors.Fields(0).Value) & "</option>"
				ors.MoveNext
			loop
		end if
		ors.NextRecordSet
		%>
	<% Else %>
		<form action=assignadmin.asp#admin method=get id=form6 name=form6>
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2>Search for Player</TH>
		</TR>
		<tr bgcolor=<%=bgcone%> height=30><td>&nbsp;&nbsp;Player Name: </td><td><input type=text name=player class=bright style="width:200"></td></tr>
	<% End IF %>
	<tr bgcolor=<%=bgcone%>><td colspan=2 align=center><input type=submit value='Make it So' class=bright id=submit2 name=submit2><input type=hidden name=SaveType value=PlayerSetLadderAdmin></td></tr>
</form>
</table>
</TD></TR>
</TABLE>
		
</td>
<td><img src="/images/spacer.gif" width="10" height="1"></td>
<td width=379>

<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" WIDTH="375" ALIGN=CENTER>
<TR><TD>
<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
<form action=saveItem.asp method=post id=form5 name=form5>
<TR BGCOLOR="#000000">
	<TH>Remove a player ladder admin</TH>
</TR>
<tr bgcolor=<%=bgctwo%>><td align=center><select class=bright style="width:200" name=LnkID>
<%
strSQl="select lnk.PLAdminID, p.PlayerHandle, l.PlayerLadderName "
strSQL = strSQL & " FROM lnk_pl_a lnk, tbl_players p, tbl_playerLadders l "
strSQL = strSQL & " WHERE l.PlayerLadderID = lnk.PlayerLadderID AND p.PlayerID = lnk.PlayerID ORDER BY PlayerHandle"
ors.open strsql, oconn
if not (ors.eof and ors.bof) then
	do while not ors.eof
		response.write "<option value=" & ors.fields(0).value & ">" & Server.HTMLEncode(ors.fields(1).value) & " - " & Server.HTMLEncode(ors.fields(2).value) & "</option>"
		ors.movenext
	loop
end if
%>
</select>
</td></tr>
<tr bgcolor=<%=bgcone%>><td align=center><input type=submit value='Remove Ladder Admin Rights' class=bright id=submit3 name=submit3><input type=hidden name=SaveType value=PlayerUnSetLadderAdmin></td></tr>
</form>
</table>
</TD></TR>
</TABLE>
	
</td>
<td><img src="/images/spacer.gif" width="5" height="1"></td>
</tr>
</table>
<%
Call Content2BoxEnd()
%>

<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>