<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("ladder"), """", "&quot;") 

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strLadderName, intLadderID
Dim PageNum, PerPage
Dim startNum, finishNum
Dim intTotalRecords, intCurrent, intTotalPages
Dim bgc
Dim strStatus
Dim strEnemyName, strMatchDate, strMap1, strMap2, strMap3
Dim pDate, newMDate, mm, dd, strLadderRules, intMatchID
Dim intMaps, strMapArray(6), i
' Paging
Dim pagetogo, start, finish
Dim intRants

strLadderName = Request.QueryString("ladder")

PageNum = Request.QueryString("page")
PerPage = Request.QueryString("perpage")

If Len(PageNum) = 0 Or Not(IsNumeric(PageNum)) then
	PageNum = 1
Else
	PageNum = cint(PageNum)
End If

If Len(PerPage) = 0 Or Not(IsNumeric(PerPage)) then
	PerPage = Request.Cookies("PerPage")("LadderView")
	If Len(PerPage) = 0 Or Not(IsNumeric(PerPage)) then
		PerPage = 25
	Else
		PerPage = cint(PerPage)
	End If
Else
	PerPage = cint(PerPage)
End If

intCurrent = 0

strSQL = "SELECT LadderID FROM tbl_ladders WHERE LadderName = '" & CheckString(strLadderName) & "'"
oRs.Open strSQL, oConn
If oRs.EOF Then
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "ladderlist.asp?error=1"
End If
oRs.NextRecordSet
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart(Server.HTMLEncode(strLadderName) & " Ladder")

strSQL = "EXECUTE ViewLadder @LadderName='" & CheckString(strLadderName) & "'"
oRS.PageSize = PerPage
oRS.CacheSize = PerPage
oRS.CursorLocation = adUseClient
oRs.Open strSQL, oConn, adOpenForwardOnly, adLockReadOnly ', adCmdTableDirect
If Not(oRS.EOF AND oRS.BOF) Then
	intTotalPages		= oRS.PageCount
	intTotalRecords		= oRS.RecordCount 
	If PageNum <= intTotalPages Then
		oRS.AbsolutePage	= PageNum
	Else
		oRs.AbsolutePage = 1
		PageNum = 1
	End If
	strLadderRules		= oRs.Fields("LadderRules").Value
	intLadderID			= oRS.Fields("LadderID").Value 
	intMaps				= oRS.Fields("Maps").Value
	bgc					= bgctwo

	%>
	<table align=center border=0 cellspacing=0 cellpadding=0 width=97% class="cssBordered">
	<tr BGCOLOR="#000000">
		<TH WIDTH=40 align=center>Rank</TH>
		<TH WIDTH=225 align=center>Team</TH>
		<TH WIDTH=75 align=CENTER>Record (Llamas)</TH>
		<TH width=375 align=center>Status</TH>
		<TH width=50 align=center>Rant</TH>
	</tr>
	<%
	bgc = bgctwo
	do while not ors.EOF AND oRs.AbsolutePage = PageNum
		strEnemyName = "Data Error"
		strMatchDate = "??"
		%>
		<tr bgcolor=<%=bgc%> height=40 valign=center>
		<td width=40 align=center><%= ors.Fields("Rank").Value%></td>
		<td width=225 align=left>&nbsp;<a href=viewTeam.asp?team=<% = server.urlencode(ors.Fields("TeamName").Value) %>><% =  Server.HTMLEncode(ors.Fields("TeamName").Value) & " " &  Server.HTMLEncode(ors.fields("TeamTag").value)%></a></td>
		<td width=75 align=CENTER><%=ors.Fields("Wins").Value & " / " & ors.Fields("Losses").Value & " - (" & ors.Fields("ForFeits").Value%>)</td>
		<TD WIDTH=375 ALIGN=CENTER><%
		strStatus = oRS.Fields("Status").Value 
		intMatchID = 0
		Select Case Left(uCase(strStatus), 6)
			Case "DEFEND", "ATTACK"
				If  Left(uCase(strStatus), 6)  = "DEFEND" Then
					strsql = "select m.MatchAttackerID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID,  m.MatchMap4ID,  m.MatchMap5ID, t.teamname, t.teamtag, MatchID, m.Rants "
					strsql = strsql & " from tbl_Matches m, tbl_teams t, lnk_t_l lnk "
					strsql = strsql & " where m.matchdefenderID = " & ors.Fields(4).Value 
					strsql = strsql & " and m.MatchLadderID=" & intLadderID
					strsql = strsql & " AND t.teamid = lnk.teamid "
					strsql = strsql & " AND lnk.tllinkid = m.MatchAttackerID "
					ors2.Open strSQL, oconn
					if not (ors2.EOF and ors2.BOF) then
						strEnemyName = ors2.Fields("TeamName").Value
						strMatchDate = ors2.Fields("MatchDate").Value
						strMapArray(1) = ors2.fields("MatchMap1ID").value
						strMapArray(2) = ors2.fields("MatchMap2ID").value
						strMapArray(3) = ors2.fields("MatchMap3ID").value
						strMapArray(4) = ors2.fields("MatchMap4ID").value
						strMapArray(5) = ors2.fields("MatchMap5ID").value
						intMatchID = oRs2.Fields("MatchID").Value
						intRants = oRs2.Fields("Rants").Value
					end if
					ors2.nextrecordset 
				Else
					strsql = "select m.MatchDefenderID, m.MatchDate, m.MatchMap1ID, m.MatchMap2ID, m.MatchMap3ID, m.MatchMap4ID,  m.MatchMap5ID, t.teamname, t.teamtag, MatchID, m.Rants "
					strsql = strsql & " from tbl_Matches m, tbl_teams t, lnk_t_l lnk "
					strsql = strsql & " where m.MatchAttackerID = " & ors.Fields(4).Value 
					strsql = strsql & " and m.MatchLadderID=" & intLadderID
					strsql = strsql & " AND t.teamid = lnk.teamid "
					strsql = strsql & " AND lnk.tllinkid = m.matchdefenderID "
					ors2.Open strSQL, oconn
					if not (ors2.EOF and ors2.BOF) then
						strEnemyName = ors2.Fields("TeamName").Value
						strMatchDate = ors2.Fields("MatchDate").Value
						strMapArray(1) = ors2.fields("MatchMap1ID").value
						strMapArray(2) = ors2.fields("MatchMap2ID").value
						strMapArray(3) = ors2.fields("MatchMap3ID").value
						strMapArray(4) = ors2.fields("MatchMap4ID").value
						strMapArray(5) = ors2.fields("MatchMap5ID").value
						intMatchID = oRs2.Fields("MatchID").Value
						intRants = oRs2.Fields("Rants").Value
					end if
					ors2.nextrecordset 				
				End If
				if strMatchDate <> "TBD" then
					newMDate = right(strMatchDate, len(strMatchDate)-instr(1, strMatchDate, ","))
					if len(newMDate) > 5 Then
						newMDate = Left(newmDate, (len(newMDate) - 4))
						newMDate = formatdatetime(newMDate, 2)
						mm=month(newmdate)
						dd=day(newmdate)
					End If
					pDate=mm & "/" & dd
				else
					pdate="TBD"
				end if
				Response.Write left(strStatus,3) & " v. <a href=viewTeam.asp?team=" & Server.URLEncode(strEnemyName) & ">"
				Response.Write Server.HTMLEncode(strEnemyName) & "</a> (" & pDate & ")"
				if pdate <> "TBD" then
					response.write "<br>("
					For i = 1 to intMaps
						If i > 1 Then
							Response.Write ", "
						End If
						Response.Write Server.HTMLEncode(strMapArray(i))
					Next
					Response.Write ")"
'					Server.HTMLEncode(strMap1) & ", " &  Server.HTMLEncode(strMap2) & ", " &  Server.HTMLEncode(strMap3) & ")"
				end if
			Case "IMMUNE", "DEFEAT", "RESTIN"
				Response.Write strStatus
			Case Else
				Response.Write "<B>Open</B>"
		End Select
		Response.Write "</TD>"
		If intMatchID > 0 Then
			Response.Write "<td align=""center""><a href=""viewmatch.asp?Ladder=" & Server.URLEncode(strLadderName & "") & "&MatchID=" & intMatchID & """>rant (" & intRants & ")</a></td>"
		Else
			Response.Write "<td> &nbsp;</td>"
		End If
		
		Response.Write "</TR>"
		ors.MoveNext

		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		intCurrent = intCurrent + 1
		If intCurrent = 10 Then
			Response.Write "</TABLE>"
			Call ContentEnd()
			Call ContentStart("")
			%>
			<table align=center border=0 cellspacing=0 cellpadding=0 width=97% class="cssBordered">
			<tr BGCOLOR="#000000">
				<TH WIDTH=40 align=center>Rank</TH>
				<TH WIDTH=225 align=center>Team</TH>
				<TH WIDTH=75 align=CENTER>Record (Llamas)</TH>
				<TH width=375 align=center>Status</TH>
				<TH width=50 align=center>Rant</TH>
			</tr>
			<%
		End If
	loop
	%>
	</table>

	<BR>
	<table ALIGN=CENTER border=0 cellspacing=0 cellpadding=0 Width=97%>
	<TR>
	<TD VALIGN=TOp>
	<UL>
	<LI><B><a href="/viewladderdetails.asp?ladder=<%=Server.URLEncode(strLadderName & "")%>">Ladder Information</A></B>
	<% If Not(IsNull(strLadderRules) or Len(Trim(strLadderRules)) = 0) Then %>
	<LI><B><a href="rules.asp?set=<%=Server.URLEncode(strLadderRules & "")%>" target="_blank"><%=strLadderName%> Rules</A></B>
	<% End If %>
	<LI><B><a href="/ladderhistory.asp?ladder=<%=Server.URLEncode(strLadderName & "")%>">Recent History</A></B>
	<LI><B><a href="/xml/viewladder.asp?ladder=<%=Server.URLEncode(strLadderName & "")%>&rank_start=1&rank_end=25">XML Version</A></B>
	<LI><B><a href="/xml/matchinfo.asp?ladder=<%=Server.URLEncode(strLadderName & "")%>&pendingdays=2&recentdays=2">XML History and Pending Matches</A></B>
	<LI><B><a href="/mapstats.asp?ladder=<%=Server.URLEncode(strLadderName & "")%>">Map Statistics</A></B>
	</UL>

	</TD>
	<TD ALIGN=RIGHT>
	<table ALIGN=RIGHT border=0 cellspacing=0 cellpadding=0 width=230 class="cssBordered">
	<tr>
	<TH BGCOLOR="#000000" COLSPAN=4>Total Teams: <%=intTotalRecords%></TH></tr>
	<tr height=25>
	<%
	if pageNum <> 1 then 
		%>
		<td bgcolor=<%=bgcone%> align=left colspan=2>&nbsp;<input type=button class=bright value="<-- Previous" style="width: 75px" onclick="window.location.href='viewladder.asp?page=<%=(pagenum - 1)%>&perpage=<%=perpage%>&ladder=<%=server.urlencode(strLadderName)%>'"></td>
		<% 		
	else 
		response.write "<td bgcolor=" & bgcone & " colspan=2><p class=small>&nbsp;</p></td>"
	end if
	if pageNum < intTotalPages then 
		%>
		<td bgcolor=<%=bgcone%> align=right colspan=2><input type=button class=bright value="Next -->"style="width: 75px" onclick="window.location.href='viewladder.asp?page=<%=(pagenum + 1)%>&perpage=<%=perpage%>&ladder=<%=server.urlencode(strLadderName)%>'">&nbsp;</td>
		<%
	else 
		response.write "<td bgcolor=" & bgcone & " colspan=2><p class=small>&nbsp;</p></td>"
	end if
	bgc=bgcone
	%>
	</tr>
	<% 
	intCurrent = intTotalRecords
	pagetogo = 0
	Do While intCurrent > 0
		If pagetogo Mod 4 = 0 Then
		 Response.Write "</tr><tr>"
		 if bgc=bgcone then 
		 	bgc=bgctwo
		 else 
		 	bgc=bgcone
		 end if
		End If
		pagetogo = pagetogo + 1
		intCurrent = intCurrent - perpage
		start = perpage*(pagetogo-1)
		finish = pagetogo*perpage
		if finish > intTotalRecords then
			finish = intTotalRecords
		end if
		if (pagetogo - pagenum = 0) then
			response.write "<td width=55 bgcolor=" & bgc & " align=center>" & start+1 & " - " & finish & "</td>"
		else
			%>
			<td bgcolor=<%=bgc%> width=55 align=center><a href="viewladder.asp?page=<%=pagetogo%>&perpage=<%=perpage%>&ladder=<%=server.urlencode(strLadderName)%>"><%=start+1%> - <%=finish%></a></td>
			<%
		end if
	loop
	if pagetogo Mod 4 = 0 Then
		Response.Write "</TR>"
	Else	
		response.write "<td bgcolor=" & bgc &" COLSPAN="& (4 - (pagetogo Mod 4)) &">&nbsp;</td></tr>"
	End If
	%>
	<tr>
	<td bgcolor=<%=bgctwo%> height=20 align=center colspan=4><p class=small>[ <b>Per Page: <%=perpage%></b> ] [ <b>Total Pages: <%=pagetogo%></b> ]</p></td></tr>

	</table>
	</TD></TR>
	</TABLE>
<% else %>
	<table align=center border=0 cellspacing=0 cellpadding=0 width=755 BGCOLOR="#444444">
	<TR><TD>
	<table align=center border=0 cellspacing=1 cellpadding=0 width=100%>
	<tr BGCOLOR="#000000">
		<TH WIDTH=40 align=center>Rank</TH>
		<TH WIDTH=225 align=center>Team</TH>
		<TH WIDTH=75 align=CENTER>Record (Llamas)</TH>
		<TH width=425 align=center>Status</TH>
	</tr>
	<TR>
		<TD Colspan=4 BGCOLOR="#000000">&nbsp;&nbsp;<i><font color=red>No teams have signed up for this ladder yet.</font></i></TD>
	</TR>
	</TABLE>
	</TD></TR>
	</TABLE>
<% end if %>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>