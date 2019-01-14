<% 	
	bgcone = "#000031"
	bgctwo = "#212152"
	
	set oConn = Server.CreateObject("ADODB.Connection")
	oConn.connectionstring="file name=c:\twl.udl"
	oConn.Open 
	set oRs = Server.CreateObject("ADODB.Recordset")
	set oRs2 = Server.CreateObject("ADODB.Recordset")
	set oRs1 = Server.CreateObject("ADODB.Recordset")

%>
	<table width="97%" border="0" cellspacing="0" cellpadding="0">
	<tr bgcolor=<%=bgcone%> height=25>
	<td width=100><p class=small><b>Defender</b></p></td>
	<td width=100><p class=small><b>Attacker</b></p></td>
	<td width=100><p class=small><b>Result</b></p></td>
	<td><p class=small><b>Date</b></p></td>
	<td align=center><p class=small><b>Winner</b></p></td>
</tr>
<%	strSQL="select top 20 * from tbl_history where MatchLadderID='5' and MatchMap1 is not Null order by historyid desc"
	ors.Open strSQL, oconn
	if not (ors.eof and ors.BOF) then
		do while not ors.EOF
			defwin=ors.fields("matchwinnerdefending").value

			strsql="select teamname, lnk_t_l.rank from tbl_teams inner join lnk_T_L on lnk_T_L.teamid=tbl_teams.teamid where TLLInkID=" & ors.Fields("MatchLoserID").Value
			ors2.open strsql, oconn
			losername = "Unknown"
			if not(ors2.eof and ors2.bof) then
				losername=ors2.fields(0).value
				loserrank=ors2.fields(1).value
			end if
			ors2.close
			strsql="select teamname, lnk_t_l.rank from tbl_teams inner join lnk_T_L on lnk_T_L.teamid=tbl_teams.teamid where TLLInkID=" & ors.Fields("MatchWinnerID").Value
			ors2.open strsql, oconn
			winnername = "Unknown"
			if not(ors2.eof and ors2.bof) then
				winnername=ors2.fields(0).value
				winnerrank=ors2.fields(1).value
			end if
			ors2.close
			
			if ors.Fields("MatchWinnerID").Value = teamid then
				strsql="select teamname from tbl_teams inner join lnk_T_L on lnk_T_L.teamid=tbl_teams.teamid where TLLInkID=" & ors.Fields("MatchLoserID").Value
				result="Win"
			else
				strsql="select teamname from tbl_teams inner join lnk_T_L on lnk_T_L.teamid=tbl_teams.teamid where TLLInkID=" & ors.Fields("MatchwinnerID").Value
				result="Loss"
			end if

			ors2.Open strsql, oconn
			if not (ors2.EOF and ors2.BOF) then
				oname=ors2.Fields(0).Value 
			end if
			ors2.Close
			
			'assume that att won
			
			attname = winnername
			defname = losername
			attrank = winnerrank
			defrank = loserrank
			result = "<font color='brightred'><b>Attacker Won</b>"
			If defwin then
				defname = winnername
				defrank = winnerrank
				attrank = loserrank
				attname = losername
				result = "<font color='Green'><b>Defender Won</b>"
			end if				
			if defrank = "" then
				defrank = "0"
			end if
			if attrank = "" then
				attrank = "0"
			end if

			map1=ors.fields("MatchMap1").value
			map1defscore=ors.fields("MatchMap1DefenderScore").value
			map1attscore=ors.fields("MatchMap1AttackerScore").value
			map1ot=ors.fields("matchmap1ot").value
			map1ft=ors.fields("matchmap1forfeit").value
			map2=ors.fields("MatchMap2").value
			map2defscore=ors.fields("MatchMap2DefenderScore").value
			map2attscore=ors.fields("MatchMap2AttackerScore").value
			map2ot=ors.fields("matchmap2ot").value
			map2ft=ors.fields("matchmap2forfeit").value
			map3=ors.fields("MatchMap3").value
			map3defscore=ors.fields("MatchMap3DefenderScore").value
			map3attscore=ors.fields("MatchMap3AttackerScore").value
			map3ot=ors.fields("matchmap3ot").value
			map3ft=ors.fields("matchmap3forfeit").value
			%>
			<tr bgcolor=<%=bgctwo%>>
			<td><p class=small><a href="http://www.teamwarfare.com/viewteam.asp?team=<%=server.urlencode(defname)%>" onmouseover="(window.status='View team\'s profile'); return true" onmouseout="(window.status='<%=javatitle%>'); return true">#<%=defrank%>. <%=server.htmlencode(defname)%></a></p></td>
			<td><p class=small><a href="http://www.teamwarfare.com/viewteam.asp?team=<%=server.urlencode(attname)%>" onmouseover="(window.status='View team\'s profile'); return true" onmouseout="(window.status='<%=javatitle%>'); return true">#<%=attrank%>. <%=server.htmlencode(attname)%></a></p></td>
			<td><p class=small><%=result%></p></td><td><p class=small><%=ors.Fields("MatchDate").Value%></p></td><td align=center><p class=small>
			<% =server.htmlencode(winnername)%>
			   </p></td>
			</tr>
			<tr>
			<td><img src="http://www.teamwarfare.com/images/spacer.gif" height="1"></td>
			<td align=left height=20 colspan=3  bgcolor=<%=bgcone%>>
			<p class=small>&nbsp;<b><%=server.htmlencode(map1)%>:</b> <%=map1defscore%> - <%=map1attscore%><%
			if map1ot then
				response.write " in OT"
			end if
			if map1ft then
				response.write " by forfiet"
			end if
			%><br>
			&nbsp;<b><%=server.htmlencode(map2)%>:</b> <%=map2defscore%> - <%=map2attscore%><%
			if map2ot then
				response.write " in OT"
			end if
			if map2ft then
				response.write " by forfiet"
			end if
			if NOT ((map3usscore = 0) and (map3themscore=0) and not map3ot and not map3ft) then
			%><br>
			&nbsp;<b><%=server.htmlencode(map3)%>:</b> <%=map3defscore%> - <%=map3attscore%><%
			if map3ot then
				response.write " in OT"
			end if
			if map3ft then
				response.write " by forfiet"
			end if
			else
			%><br>
			&nbsp;<b><%=server.htmlencode(map3)%>:</b> not played
			<%
			end if
			%>
			</p>
			</td></tr>
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.MoveNext
		loop
	end if
	ors.Close 
	oConn.Close
	Set oConn=Nothing
	Set oRS = Nothing
	Set oRs1 = Nothing	
%>