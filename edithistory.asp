<% 'Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Edit History"

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
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER>
	<TR><TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=4>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=3>Choose a Ladder</TH>
	</TR>
<%
	If bSysAdmin Then 
		strsql="Select ladderID, ladderName from TBL_ladders WHERE LadderActive = 1 order by LadderName"
	Else
		strSQL = "SELECT l.ladderID, ladderName from tbl_ladders l, lnk_l_a lnk WHERE LadderActive = 1  AND lnk.LadderID = l.LadderiD AND lnk.PlayerID = '" & Session("PlayerID") & "' order by LadderName"
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
			response.write "<td>&nbsp;<a href=edithistory.asp?ladder=" & server.urlencode(ors.fields(1).value) & ">" & Server.htmlencode(ors.fields(1).value) & "</a></td>"
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
		strsql="SELECT * FROM vHistory WHERE LadderName='" & CheckString(strLadder) & "' ORDER BY MatchDate desc"
		ors.open strsql, oconn
		if not (ors.eof and ors.bof) then
			response.write "<form action=edithistory.asp?ladder=" & server.urlencode(strLadder) & " method=post name=choosematch id=matchchooser>"
			response.write "<table width=97% border=0 cellspacing=0 cellpadding=0>"
			response.write "<tr height=20><td align=center><p class=small><b>Select match to edit:</b></p></td></tr>"
			response.write "<tr height=35><td align=center><select name=HistoryID>"
			counter=0
			do while not (ors.eof)
				response.write "<option value=" & ors.fields("HistoryID").value & ">" & oRS.Fields("WinnerName").Value  & " vs. " & oRS.Fields("LoserName").Value & " (" & ors.fields("MatchDate").value & ")"
				ors.movenext
			loop
			response.write "</select></td></tr>"
			response.write "<tr height=35><td align=center><input type=submit class=bright value='Edit this Match' style=""width:150"" name=submit1 value=submit1></td></tr>"
			If bSysAdmin Then
				response.write "<tr height=35><td align=center><input type=BUTTON class=bright value='Delete This Match' style=""width:150"" name=button value=button ONCLICK=""javascript:Delete(this.form);""></td></tr>"
			End If
			Response.Write "<INPUT TYPE=HIDDEN NAME=""SaveType"" VALUE=""DeleteHistory"">"
			Response.Write "<INPUT TYPE=HIDDEN NAME=""Ladder"" VALUE=""" & Server.HTMLEncode(strLadder) & """>"
			response.write "</table></form>"
		end if
		oRS.NextRecordset 
		
		historyid = request.form("HistoryID")
		if historyID <> "" then
			strsql = "select * from vHistory where HistoryID=" & request.form("HistoryID")
			ors.open strsql, oconn
			if not (ors.eof and ors.bof) then
				laddername = ors.fields("LadderName").value
				ladderid = ors.fields("MatchLadderID").value
				matchdate = ors.fields("MatchDate").value
				defwin = ors.fields("WinnerDefending").value
				if isNull(defwin) Then
					defwin = 0
				End IF
				map1=ors.fields("MatchMap1").value
				map1defscore=ors.fields("MatchMap1DefenderScore").value
				map1attscore=ors.fields("MatchMap1AttackerScore").value
				map1ot=ors.fields("map1ot").value
				map1ft=ors.fields("map1forfeit").value
				map2=ors.fields("MatchMap2").value
				map2defscore=ors.fields("MatchMap2DefenderScore").value
				map2attscore=ors.fields("MatchMap2AttackerScore").value
				map2ot=ors.fields("map2ot").value
				map2ft=ors.fields("map2forfeit").value
				map3=ors.fields("MatchMap3").value
				map3defscore=ors.fields("MatchMap3DefenderScore").value
				map3attscore=ors.fields("MatchMap3AttackerScore").value
				map3ot=ors.fields("map3ot").value
				map3ft=ors.fields("map3forfeit").value
				map4=ors.fields("MatchMap4").value
				map4defscore=ors.fields("MatchMap4DefenderScore").value
				map4attscore=ors.fields("MatchMap4AttackerScore").value
				map4ot=ors.fields("map4ot").value
				map4ft=ors.fields("map4forfeit").value
				map5=ors.fields("MatchMap5").value
				map5defscore=ors.fields("MatchMap5DefenderScore").value
				map5attscore=ors.fields("MatchMap5AttackerScore").value
				map5ot=ors.fields("map5ot").value
				map5ft=ors.fields("map5forfeit").value
				matchft=ors.Fields("matchforfeit").value
				winid = ors.fields("MatchWinnerID").value
				losid = ors.fields("MatchLoserID").value
				winrank=ors.fields("WinnerRank").value
				losrank = ors.fields("LoserRank").value
				winname = ors.fields("WinnerName").value
				losname = ors.fields("LoserName").value
'						response.write "<table width=97% border=0 cellspacing=0 cellpadding=0>"
				if defwin then
					defname = winname
					attname = losname
					defid = winid
					attid = losid
					defrank = winrank
					attrank = losrank
'							response.write "<tr><td><p class=small>Winner: " & htmlencode(winname) & "</p></td></tr>"
'							response.write "<tr><td><p class=small><b>Since the defender won, there was no change in rank</b></p></td></tr>"
				else
					defname = losname
					attname = winname
					defid = losid
					attid = winid
					defrank = losrank
					attrank = winrank
'							response.write "<tr><td><p class=small>Winner: " & htmlencode(winname) & "</p></td></tr>"
'							response.write "<tr><td><p class=small><b>Since the attacker won, there was change in rank</b></p></td></tr>"
				end if

				if map1ot = 0 then
					map1otverb = "no"
				else
					map1otverb = "yes"
				end if
				if map1ft = 0 then
					map1ftverb = "no"
				else
					map1ftverb = "yes"
				end if
				
				if map2ot = 0 then
					map2otverb = "no"
				else
					map2otverb = "yes"
				end if
				if map2ft = 0 then
					map2ftverb = "no"
				else
					map2ftverb = "yes"
				end if
				
				if map3ot = 0 then
					map3otverb = "no"
				else
					map3otverb = "yes"
				end if
				if map3ft = 0 then
					map3ftverb = "no"
				else
					map3ftverb = "yes"
				end if
				

				if map4ot = 0 then
					map4otverb = "no"
				else
					map4otverb = "yes"
				end if
				if map4ft = 0 then
					map4ftverb = "no"
				else
					map4ftverb = "yes"
				end if

				if map5ot = 0 then
					map5otverb = "no"
				else
					map5otverb = "yes"
				end if
				if map5ft = 0 then
					map5ftverb = "no"
				else
					map5ftverb = "yes"
				end if
				maps = oRS.Fields("Maps").Value 
'						response.write "<tr><td><p class=small>Defender: " & htmlencode(defname) & defrank & "</p></td></tr>"
'						response.write "<tr><td><p class=small>Attacker: " & htmlencode(attname) & attrank & "</p></td></tr>"
'						response.write "<tr><td><p class=small>Map 1: " & htmlencode(map1) & "</p></td></tr>"
'						response.write "<tr><td><p class=small>" & htmlencode(defname) & " -  " & htmlencode(map1defscore) & "</p></td></tr>"
'						response.write "<tr><td><p class=small>" & htmlencode(attname) & " -  " & htmlencode(map1attscore) & "</p></td></tr>"
'						Response.Write "<tr><td><p class=small>OverTime: " & htmlencode(map1otverb) & "</p></td></tr>"
'						Response.Write "<tr><td><p class=small>Forfeit: " & htmlencode(map1ftverb) & "</p></td></tr>"
'						response.write "<tr><td><p class=small>Map 2: " & htmlencode(map2) & "</p></td></tr>"
'						response.write "<tr><td><p class=small>" & htmlencode(defname) & " -  " & htmlencode(map2defscore) & "</p></td></tr>"
'						response.write "<tr><td><p class=small>" & htmlencode(attname) & " -  " & htmlencode(map2attscore) & "</p></td></tr>"
'						Response.Write "<tr><td><p class=small>OverTime: " & htmlencode(map2otverb) & "</p></td></tr>"
'						Response.Write "<tr><td><p class=small>Forfeit: " & htmlencode(map2ftverb) & "</p></td></tr>"
'						response.write "<tr><td><p class=small>Map 3: " & htmlencode(map3) & "</p></td></tr>"
'						response.write "<tr><td><p class=small>" & htmlencode(defname) & " -  " & htmlencode(map3defscore) & "</p></td></tr>"
'						response.write "<tr><td><p class=small>" & htmlencode(attname) & " -  " & htmlencode(map3attscore) & "</p></td></tr>"
'						Response.Write "<tr><td><p class=small>OverTime: " & htmlencode(map3otverb) & "</p></td></tr>"
'						Response.Write "<tr><td><p class=small>Forfeit: " & htmlencode(map3ftverb) & "</p></td></tr>"
'						Response.Write "<tr><td><p class=small>Overall Forfeit?: " & htmlencode(matchft) & "</p></td></tr>"
'						response.write "</table>"
						
				Response.Write "<form name=MatchChange id=changer action=saveitem.asp method=post>"
						
				Response.Write "<input type=hidden name=LadderID value=" & ladderid & ">"
				Response.Write "<input type=hidden name=MatchDate value='" & matchdate & "'>"
				Response.Write "<input type=hidden name=HistoryID value=" & historyid & ">"
				Response.Write "<input type=hidden name=DefOldRank value=" & defrank & ">"
				Response.Write "<input type=hidden name=AttOldRank value=" & attrank & ">"
				Response.Write "<input type=hidden name=DefenderId value=" & DefID & ">"
				Response.Write "<input type=hidden name=AttackerId value=" & AttID & ">"
				Response.Write "<input type=hidden name=SaveType value=ChangeMatch>"
				Response.Write "<input type=hidden name=DefOldWin value=" & defwin & ">"
						
				Response.Write "<table width=97% cellspacing=0 border=0 cellpadding=0>"
				Response.Write "<tr height=25><td colspan=2 align=center><p class=headline><b>" & Server.HTMLEncode(defname) & " defended against " & Server.HTMLEncode(attname) & "</b></p></td></tr>"
				Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Defender Current Rank (#" & Server.HTMLEncode(defrank & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=DefNewRank value='" & Server.HTMLEncode(defrank & "") & "' id=DefNewRank1 class=bright style=""width:50""></p></td></tr>"
				Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Attacker Current Rank (#" & Server.HTMLEncode(attrank & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=AttNewRank value='" & Server.HTMLEncode(attrank & "") & "' id=AttNewRank1 class=bright style=""width:50""></p></td></tr>"

				Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 1: (" & Server.HTMLEncode(map1 & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map1Name value='" & Server.HTMLEncode(Map1 & "") & "' id=Map1Name class=bright style=""width:150""></p></td></tr>"
				Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Defender Score Map 1: (" & Server.HTMLEncode(map1defscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map1DefScore value='" & Server.HTMLEncode(Map1DefScore & "") & "' id=Map1DefScore class=bright style=""width:25""></p></td></tr>"
				Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Attacker Score Map 1: (" & Server.HTMLEncode(map1attscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map1AttScore value='" & Server.HTMLEncode(Map1attScore & "") & "' id=Map1AttScore class=bright style=""width:25""></p></td></tr>"
				if map1ot then
					Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 1 OT?: (" & Server.HTMLEncode(map1otverb & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=Map1OT value=0 id=Map1OTNo> Yes: <input type=radio class=borderless name=Map1OT value=1 id=Map1OTYes checked></p></td></tr>"
				else
					Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 1 OT?: (" & Server.HTMLEncode(map1otverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=Map1OT value=0 id=Map1OTNo checked> Yes: <input type=radio class=borderless name=Map1OT value=1 id=Map1OTYes></p></td></tr>"
				end if
				if map1ft then
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 1 FT?: (" & Server.HTMLEncode(map1ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=Map1FT value=0 id=Map1FTNo> Yes: <input type=radio class=borderless name=Map1FT value=1 id=Map1FTYes checked></p></td></tr>"
				else
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 1 FT?: (" & Server.HTMLEncode(map1ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=Map1FT value=0 id=Map1FTNo checked> Yes: <input type=radio class=borderless name=Map1FT value=1 id=Map1FTYes></p></td></tr>"
				end if
				
				If maps > 1 Then
					Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 2: (" & Server.HTMLEncode(map2 & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map2Name value='" & Server.HTMLEncode(Map2 & "") & "' id=Map2Name class=bright style=""width:150""></p></td></tr>"
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Defender Score Map 2: (" & Server.HTMLEncode(map2defscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map2DefScore value='" & Server.HTMLEncode(Map2DefScore & "") & "' id=Map2DefScore class=bright style=""width:25""></p></td></tr>"
					Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Attacker Score Map 2: (" & Server.HTMLEncode(map2attscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map2AttScore value='" & Server.HTMLEncode(Map2attScore & "") & "' id=Map2AttScore class=bright style=""width:25""></p></td></tr>"
					if map2ot then
						Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 2 OT?: (" & Server.HTMLEncode(map2otverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map2OT value=0 id=map2OTNo> Yes: <input type=radio class=borderless name=map2OT value=1 id=map2OTYes checked></p></td></tr>"
					else
						Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 2 OT?: (" & Server.HTMLEncode(map2otverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map2OT value=0 id=map2OTNo checked> Yes: <input type=radio class=borderless name=map2OT value=1 id=map2OTYes></p></td></tr>"
					end if
					if map2ft then
						Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 2 FT?: (" & Server.HTMLEncode(map2ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map2FT value=0 id=map2FTNo> Yes: <input type=radio class=borderless name=map2FT value=1 id=map2FTYes checked></p></td></tr>"
					else
						Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 2 FT?: (" & Server.HTMLEncode(map2ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map2FT value=0 id=map2FTNo checked> Yes: <input type=radio class=borderless name=map2FT value=1 id=map2FTYes></p></td></tr>"
					end if
				End If
				
				If maps > 2 Then
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 3: (" & Server.HTMLEncode(map3 & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map3Name value='" & Server.HTMLEncode(Map3 & "") & "' id=Map3Name class=bright & ""t style=""width:150""></p></td></tr>"
					Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Defender Score Map 3: (" & Server.HTMLEncode(map3defscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map3DefScore value='" & Server.HTMLEncode(Map3DefScore & "") & "' id=Map3DefScore class=bright style=""width:25""></p></td></tr>"
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Attacker Score Map 3: (" & Server.HTMLEncode(map3attscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map3AttScore value='" & Server.HTMLEncode(Map3attScore & "") & "' id=Map3AttScore class=bright style=""width:25""></p></td></tr>"
					if map3ot then
						Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 3 OT?: (" & Server.HTMLEncode(map3otverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map3OT value=0 id=map3OTNo> Yes: <input type=radio class=borderless name=map3OT value=1 id=map3OTYes checked></p></td></tr>"
					else
						Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 3 OT?: (" & Server.HTMLEncode(map3otverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map3OT value=0 id=map3OTNo checked> Yes: <input type=radio class=borderless name=map3OT value=1 id=map3OTYes></p></td></tr>"
					end if
					if map3ft then
						Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 3 FT?: (" & Server.HTMLEncode(map3ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map3FT value=0 id=map3FTNo> Yes: <input type=radio class=borderless name=map3FT value=1 id=map3FTYes checked></p></td></tr>"
					else
						Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 3 FT?: (" & Server.HTMLEncode(map3ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map3FT value=0 id=map3FTNo checked> Yes: <input type=radio class=borderless name=map3FT value=1 id=map3FTYes></p></td></tr>"
					end if
				End If
						
				If maps > 3 Then
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 4: (" & Server.HTMLEncode(map4 & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map4Name value='" & Server.HTMLEncode(Map4 & "") & "' id=Map4Name class=bright & ""t style=""width:150""></p></td></tr>"
					Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Defender Score Map 4: (" & Server.HTMLEncode(map4defscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map4DefScore value='" & Server.HTMLEncode(Map4DefScore & "") & "' id=Map4DefScore class=bright style=""width:25""></p></td></tr>"
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Attacker Score Map 4: (" & Server.HTMLEncode(map4attscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map4AttScore value='" & Server.HTMLEncode(Map4attScore & "") & "' id=Map4AttScore class=bright style=""width:25""></p></td></tr>"
					if map4ot then
						Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 4 OT?: (" & Server.HTMLEncode(map4otverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map4OT value=0 id=map4OTNo> Yes: <input type=radio class=borderless name=map4OT value=1 id=map4OTYes checked></p></td></tr>"
					else
						Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 4 OT?: (" & Server.HTMLEncode(map4otverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map4OT value=0 id=map4OTNo checked> Yes: <input type=radio class=borderless name=map4OT value=1 id=map4OTYes></p></td></tr>"
					end if
					if map4ft then
						Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 4 FT?: (" & Server.HTMLEncode(map4ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map4FT value=0 id=map4FTNo> Yes: <input type=radio class=borderless name=map4FT value=1 id=map4FTYes checked></p></td></tr>"
					else
						Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 4 FT?: (" & Server.HTMLEncode(map4ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map4FT value=0 id=map4FTNo checked> Yes: <input type=radio class=borderless name=map4FT value=1 id=map4FTYes></p></td></tr>"
					end if
				End If
				
				If maps > 4 Then
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 5: (" & Server.HTMLEncode(map5 & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map5Name value='" & Server.HTMLEncode(Map5 & "") & "' id=Map5Name class=bright & ""t style=""width:150""></p></td></tr>"
					Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Defender Score Map 5: (" & Server.HTMLEncode(map5defscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map5DefScore value='" & Server.HTMLEncode(Map5DefScore & "") & "' id=Map5DefScore class=bright style=""width:25""></p></td></tr>"
					Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Attacker Score Map 5: (" & Server.HTMLEncode(map5attscore & "") & "):&nbsp;</b></p></td><td width=50% align=left><p class=small><input type=text name=Map5AttScore value='" & Server.HTMLEncode(Map5attScore & "") & "' id=Map5AttScore class=bright style=""width:25""></p></td></tr>"
					if map5ot then
						Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 5 OT?: (" & Server.HTMLEncode(map5otverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map5OT value=0 id=map5OTNo> Yes: <input type=radio class=borderless name=map5OT value=1 id=map5OTYes checked></p></td></tr>"
					else
						Response.Write "<tr bgcolor=" & bgctwo & " height=22><td width=50% align=right><p class=small><b>Map 5 OT?: (" & Server.HTMLEncode(map5otverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map5OT value=0 id=map5OTNo checked> Yes: <input type=radio class=borderless name=map5OT value=1 id=map5OTYes></p></td></tr>"
					end if
					if map5ft then
						Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 5 FT?: (" & Server.HTMLEncode(map5ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map5FT value=0 id=map5FTNo> Yes: <input type=radio class=borderless name=map5FT value=1 id=map5FTYes checked></p></td></tr>"
					else
						Response.Write "<tr bgcolor=" & bgcone & " height=22><td width=50% align=right><p class=small><b>Map 5 FT?: (" & Server.HTMLEncode(map5ftverb) & "):&nbsp;</b></p></td><td width=50% align=left><p class=small>No: <input type=radio class=borderless name=map5FT value=0 id=map5FTNo checked> Yes: <input type=radio class=borderless name=map5FT value=1 id=map5FTYes></p></td></tr>"
					end if
				End If
				Response.Write "<INPUT TYPE=HIDDEN NAME=NumMaps VALUE=" & maps & ">"

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