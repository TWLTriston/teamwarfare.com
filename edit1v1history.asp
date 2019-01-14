<% 'Option Explicit %>
<%
'// function to retrieve player names
Function getPlayerNameByID(id,oConn)

	Dim rs, sql, name
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
	sql = "EXECUTE getPlayerNameforHistory @ID = " & id
	
	'sql = "SELECT tbl_Players.PlayerHandle AS PlayerName, lnk_p_pl.PPLLinkID FROM dbo.tbl_Players INNER JOIN "
	'sql = sql & "lnk_p_pl ON tbl_Players.PlayerID = lnk_p_pl.PlayerID WHERE (lnk_p_pl.PPLLinkID = '" & id & "')"
	rs.Open sql,oConn
	
	If not (rs.eof and rs.bof) then
		rs.MoveFirst
		name = rs.Fields("PlayerName").Value
	Else
		name = "[UNKNOWN]"
	End If
	
	rs.Close
	
	Set rs = Nothing
	
	getPlayerNameByID = name
	
End Function

'// function to retrieve ladder name
Function getLadderNameByID(id,oConn)

	Dim rs, sql, name
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
	sql = "SELECT PlayerLadderName FROM tbl_playerladders WHERE PlayerLadderID='" & id & "' "
	rs.Open sql,oConn
	
	If not (rs.eof and rs.bof) then
		rs.MoveFirst
		name = rs.Fields("PlayerLadderName").Value
	Else
		name = "[UNKNOWN]"
	End If
	
	rs.Close
	
	Set rs = Nothing
	
	getLadderNameByID = name
	
End Function
%>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Edit 1v1 History"

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

If Not(bSysAdmin or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing	
	Set oRS2 = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim strLadder
strLadder = Request.QueryString("ladder")
If (strLadder <> "" AND Not(bSysAdmin OR IsLadderAdmin(strLadder))) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing	
	Set oRS2 = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If	
Dim LadderID, counter,i
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Edit History") %>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER ID="Table1">
	<TR><TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=4 ID="Table2">
	<TR BGCOLOR="#000000">
		<TH COLSPAN=3>Choose a Ladder</TH>
	</TR>
<%
	'// get list of player ladders 
	If bSysAdmin Then 
		strsql="Select PlayerLadderID, PlayerLadderName from tbl_playerladders WHERE Active = 1 order by PlayerLadderName"
	Else
		strSQL = "SELECT l.PlayerLadderID, l.PlayerLadderName from tbl_playerladders l, lnk_pl_a lnk WHERE l.Active = 1  AND lnk.playerLadderID = l.PlayerLadderID AND lnk.PlayerID = '" & Session("PlayerID") & "' order by l.PlayerLadderName"
	End If
	ors.open strsql, oconn
	bgc=bgctwo
	if not (ors.eof and ors.bof) then
		i = 0
		do while not (ors.eof)
			if i Mod 3 = 0 Then
				If i > 0 Then
					Response.Write "</TR>"
				End If
				Response.Write "<TR BGCOLOR=" & bgc & ">"
			End IF
			response.write "<td>&nbsp;<a href=edit1v1history.asp?ladder=" & server.urlencode(ors.fields(0).value) & ">" & Server.htmlencode(ors.fields(1).value) & "</a></td>"
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.movenext
			i = i + 1
		loop
		While i mod 3 <> 0
			Response.Write "<TD>&nbsp;</TD>"
			i = i + 1
		Wend
		Response.Write "</TR>"
	end if
	ors.close
%>
	</TABLE>
	</TD></TR>
	</TABLE>
<%
	Call ContentEnd()
	if strLadder <> "" then
%>
		<SCRIPT LANGUAGE="JavaScript">
		<!-- 
			function Delete(objForm) {
				if (confirm("Are you sure you want to delete this match from history?")) {
					if (confirm("This isnt fixable, you better be really sure...")) {
						objForm.action = "saveitem.asp";
						objForm.submit();
					}
				}
			
			}
		//-->
		</SCRIPT>
		<%
		Call ContentStart("")
		'strsql="SELECT * FROM tbl_PlayerHistory WHERE MatchLadderID='" & CheckString(strLadder) & "' ORDER BY MatchDate desc"
		strsql = "EXECUTE PlayerLadderHistory @MatchLadderID = '" & CheckString(strLadder) & "'"
		ors.open strsql, oconn
		if not (ors.eof and ors.bof) then
			response.write "<form action=edit1v1history.asp?ladder=" & server.urlencode(strLadder) & " method=post name=choosematch id=matchchooser>"
			response.write "<table width=97% border=0 cellspacing=0 cellpadding=0>"
			response.write "<tr height=20><td align=center><p class=small><b>Select match to edit:</b></p></td></tr>"
			response.write "<tr height=35><td align=center><select name=HistoryID>"
			counter=0
			do while not (ors.eof)
				response.write "<option value=" & ors.fields("HistoryID").value & ">" & oRS.Fields("WinnerHandle").Value  & " vs. " & oRS.Fields("LoserHandle").Value & " (" & ors.fields("MatchDate").value & ")"
				ors.movenext
			loop
			response.write "</select></td></tr>"
			response.write "<tr height=35><td align=center><input type=submit class=bright value='Edit this Match' style=""width:150"" name=submit1 value=submit1></td></tr>"
			If bSysAdmin Then
				response.write "<tr height=35><td align=center><input type=BUTTON class=bright value='Delete This Match' style=""width:150"" name=button value=button ONCLICK=""javascript:Delete(this.form);""></td></tr>"
			End If
			Response.Write "<INPUT TYPE=HIDDEN NAME=""SaveType"" VALUE=""Delete1v1History"">"
			Response.Write "<INPUT TYPE=HIDDEN NAME=""Ladder"" VALUE=""" & Server.HTMLEncode(strLadder) & """>"
			response.write "</table></form>"
		end if
		oRS.NextRecordset 
		
		ID="Form1" 
		
		historyid = request.form("HistoryID")
		if historyID <> "" then
			strsql = "select * from tbl_PlayerHistory where HistoryID=" & request.form("HistoryID")
			ors.open strsql, oconn
			if not (ors.eof and ors.bof) then
				laddername = getLadderNameByID(ors.fields("MatchLadderID").value,oConn)
				ladderid = ors.fields("MatchLadderID").value
				matchdate = ors.fields("MatchDate").value
				defwin = ors.fields("MatchWinnerDefending").value
				
				if isNull(defwin) Then
					defwin = 0
				End IF
				
				map1=ors.fields("MatchMap1").value
				map1defscore=ors.fields("MatchMap1DefenderScore").value
				map1attscore=ors.fields("MatchMap1AttackerScore").value
				map1ft=ors.fields("MatchMap1Forfeit").value
				matchft=ors.Fields("MatchForfeit").value
				
				winid = ors.fields("MatchWinnerID").value
				losid = ors.fields("MatchLoserID").value
				
				winname = getPlayerNameByID(ors.fields("MatchWinnerID").value,oConn)
				losname = getPlayerNameByID(ors.fields("MatchLoserID").value,oConn)
				
				attrank = ors.fields("MatchAttackerRank").value
				defrank = ors.fields("MatchDefenderRank").value
				
'						response.write "<table width=97% border=0 cellspacing=0 cellpadding=0>"
				if defwin then
					defname = winname
					attname = losname
					defid = winid
					attid = losid
					'defrank = winrank
					'attrank = losrank
'							response.write "<tr><td><p class=small>Winner: " & htmlencode(winname) & "</p></td></tr>"
'							response.write "<tr><td><p class=small><b>Since the defender won, there was no change in rank</b></p></td></tr>"
				else
					defname = losname
					attname = winname
					defid = losid
					attid = winid
					'defrank = losrank
					'attrank = winrank
'							response.write "<tr><td><p class=small>Winner: " & htmlencode(winname) & "</p></td></tr>"
'							response.write "<tr><td><p class=small><b>Since the attacker won, there was change in rank</b></p></td></tr>"
				end if
				
				Response.Write "<form name=MatchChange id=changer action=saveitem.asp method=post>"
						
				Response.Write "<input type=hidden name=LadderID value=" & ladderid & ">"
				Response.Write "<input type=hidden name=MatchDate value='" & matchdate & "'>"
				Response.Write "<input type=hidden name=HistoryID value=" & historyid & ">"
				Response.Write "<input type=hidden name=DefOldRank value=" & defrank & ">"
				Response.Write "<input type=hidden name=AttOldRank value=" & attrank & ">"
				Response.Write "<input type=hidden name=DefenderId value=" & DefID & ">"
				Response.Write "<input type=hidden name=AttackerId value=" & AttID & ">"
				Response.Write "<input type=hidden name=SaveType value=Change1v1Match>"
				Response.Write "<input type=hidden name=DefOldWin value=" & defwin & ">"
						
				Response.Write "<table width=97% cellspacing=0 border=0 cellpadding=0>"
				
				If defwin Then
					Response.Write "<tr height=25><td colspan=2 align=center><p class=headline><b> Defender " & Server.HTMLEncode(defname) & " defeated " & Server.HTMLEncode(attname) & "</b></p></td></tr>"
				Else
					Response.Write "<tr height=25><td colspan=2 align=center><p class=headline><b> Attacker " & Server.HTMLEncode(attname) & " defeated " & Server.HTMLEncode(defname) & "</b></p></td></tr>"
				End If
				
				Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Defender Rank (#" & Server.HTMLEncode(defrank & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=DefNewRank value='" & Server.HTMLEncode(defrank & "") & "' id=DefNewRank1 class=bright style=""width:50""></p></td></tr>"
				Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Attacker Rank (#" & Server.HTMLEncode(attrank & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=AttNewRank value='" & Server.HTMLEncode(attrank & "") & "' id=AttNewRank1 class=bright style=""width:50""></p></td></tr>"

				Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map: (" & Server.HTMLEncode(map1 & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map1Name value='" & Server.HTMLEncode(Map1 & "") & "' id=Map1Name class=bright style=""width:150""></p></td></tr>"
				Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Defender Score Map 1: (" & Server.HTMLEncode(map1defscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map1DefScore value='" & Server.HTMLEncode(Map1DefScore & "") & "' id=Map1DefScore class=bright style=""width:25""></p></td></tr>"
				Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Attacker Score Map 1: (" & Server.HTMLEncode(map1attscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map1AttScore value='" & Server.HTMLEncode(Map1attScore & "") & "' id=Map1AttScore class=bright style=""width:25""></p></td></tr>"
				
				if map1ft then
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map Forfeit?: (" & Server.HTMLEncode(map1ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=Map1FT value=0 id=Map1FTNo> Yes: <input type=radio class=borderless name=Map1FT value=1 id=Map1FTYes checked></p></td></tr>"
				else
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map Forfeit?: (" & Server.HTMLEncode(map1ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=Map1FT value=0 id=Map1FTNo checked> Yes: <input type=radio class=borderless name=Map1FT value=1 id=Map1FTYes></p></td></tr>"
				end if
				
				if matchft then
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Admin Forfeit?: (" & Server.HTMLEncode(map1ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=MatchFT value=0 id=MatchFTNo> Yes: <input type=radio class=borderless name=MatchFT value=1 id=MatchFTYes checked></p></td></tr>"
				else
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Admin Forfeit?: (" & Server.HTMLEncode(map1ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=MatchFT value=0 id=MatchFTNo checked> Yes: <input type=radio class=borderless name=MatchFT value=1 id=MatchFTYes></p></td></tr>"
				end if
				
				
				if defwin then
					Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Defender Win?: (" & Server.HTMLEncode(defwin & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=radio class=borderless name=DefWin checked value=True id=DefWinTrue></p></td></tr>"
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Attacker Win?: (" & Server.HTMLEncode(not(defwin) & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=radio class=borderless name=DefWin value=False id=DefWinFalse></p></td></tr>"
				else
					Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Defender Win?: (" & Server.HTMLEncode(defwin & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=radio class=borderless name=DefWin value=True id=DefWinTrue></p></td></tr>"
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Attacker Win?: (" & Server.HTMLEncode(not(defwin)) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=radio class=borderless name=DefWin checked value=False id=DefWinFalse></p></td></tr>"
				end if							
						
				Response.Write "<tr height=25><td align=center colspan=2><input type=submit class=bright value='Save Changes' name=submit2 id=submit2></td></tr>"
				Response.Write "</table></form>"
			end if
			ors.close
		end if
	end if
	Call ContentEnd()
	%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>


