<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("ladder"), """", "&quot;") & " Ladder"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo

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

strSQL = "SELECT EloLadderID FROM tbl_elo_ladders WHERE EloLadderName = '" & CheckString(strLadderName) & "'"
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
Dim iCount
iCount = 0

strSQL = "EXECUTE ViewEloLadder @LadderName='" & CheckString(strLadderName) & "', @Unranked = '0'"
ORs.Open strSQL, oConn, adOpenForwardOnly, adLockReadOnly ', adCmdTableDirect
If Not(oRS.EOF AND oRS.BOF) Then
	strLadderRules		= oRs.Fields("EloRulesName").Value
	intLadderID			= oRS.Fields("EloLadderID").Value 
	bgc					= bgctwo

	%>
	<table align=center border=0 cellspacing=0 cellpadding=0 width="85%" class="cssBordered">
	<tr>
		<th colspan="6" bgcolor="#000000">Ranked Teams</th>
	</tr>
	<tr BGCOLOR="#000000">
		<TH WIDTH=40 align=center>Rung</TH>
		<TH WIDTH=40 align=center>Rating</TH>
		<TH align=center>Team</TH>
		<TH WIDTH=50 align=CENTER>Record</TH>
		<TH width=50 align=center>Streak</TH>
		<TH width=100 align=center>Active Matches</TH>
	</tr>
	<%
	bgc = bgctwo
	do while not ors.EOF
		iCount = iCount + 1
		strEnemyName = "Data Error"
		strMatchDate = "??"
		%>
		<tr bgcolor=<%=bgc%> valign=center>
		<td align=center><%= iCount + (PerPage) * (PageNum - 1)%></td>
		<td align=center><%= ors.Fields("Rating").Value%></td>
		<td align=left>&nbsp;<a href=viewTeam.asp?team=<% = server.urlencode(ors.Fields("TeamName").Value) %>><% =  Server.HTMLEncode(ors.Fields("TeamName").Value) & " " &  Server.HTMLEncode(ors.fields("TeamTag").value)%></a></td>
		<td align=center><%=ors.Fields("Wins").Value & " / " & ors.Fields("Losses").Value%></td>
		<TD align=center><%
		If oRs.Fields("WinStreak").Value > 0 Then
			Response.Write "<font color=""green"">+" & oRs.Fields("WinStreak").Value & "</font>"
		ElseIf oRs.Fields("LossStreak").Value > 0 Then
			Response.Write "<font color=""red"">-" & oRs.Fields("LossStreak").Value & "</font>"
		Else
			Response.Write "--"
		End If
		%>
		</td>
		<td align=center>
			<%
				if ors.Fields("ActiveMatches").Value = 0 then
					response.write "&nbsp;"
				else
					response.write ors.Fields("ActiveMatches").Value
				end if
			%></td>
		</tr>
		<%		
		ors.MoveNext
		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
	loop
	%>
	</table>
<% else %>
	<table align=center border=0 cellspacing=0 cellpadding=0 width=85% class="cssBordered">
	<tr BGCOLOR="#000000">
		<TH WIDTH=40 align=center>Rating</TH>
		<TH WIDTH=225 align=center>Team</TH>
		<TH WIDTH=100 align=CENTER>Record</TH>
		<TH width=100 align=center>Current Streak</TH>
	</tr>
	<TR>
		<TD Colspan=4 BGCOLOR="#000000">&nbsp;&nbsp;<i><font color=red>No ranked teams yet.</font></i></TD>
	</TR>
	</TABLE>
<% 
end if 
oRs.NextRecordSet
%>
<br /><br />
<%
strSQL = "EXECUTE ViewEloLadder @LadderName='" & CheckString(strLadderName) & "', @Unranked = '1'"
ORs.Open strSQL, oConn, adOpenForwardOnly, adLockReadOnly ', adCmdTableDirect
If Not(oRS.EOF AND oRS.BOF) Then
	strLadderRules		= oRs.Fields("EloRulesName").Value
	intLadderID			= oRS.Fields("EloLadderID").Value 
	bgc					= bgctwo

	%>
	<table align=center border=0 cellspacing=0 cellpadding=0 width="85%" class="cssBordered">
	<tr>
		<th colspan="6" bgcolor="#000000">Unranked Teams</th>
	</tr>
	<tr BGCOLOR="#000000">
		<TH WIDTH=40 align=center>Rating</TH>
		<TH align=center>Team</TH>
		<TH WIDTH=50 align=CENTER>Record</TH>
		<TH width=50 align=center>Streak</TH>
		<TH width=100 align=center>Active Matches</TH>
	</tr>
	<%
	bgc = bgctwo
	do while not ors.EOF
		strEnemyName = "Data Error"
		strMatchDate = "??"
		%>
		<tr bgcolor=<%=bgc%> valign=center>
		<td align=center><%= ors.Fields("Rating").Value%></td>
		<td align=left>&nbsp;<a href=viewTeam.asp?team=<% = server.urlencode(ors.Fields("TeamName").Value) %>><% =  Server.HTMLEncode(ors.Fields("TeamName").Value) & " " &  Server.HTMLEncode(ors.fields("TeamTag").value)%></a></td>
		<td align=center><%=ors.Fields("Wins").Value & " / " & ors.Fields("Losses").Value%></td>
		<TD align=center><%
		If oRs.Fields("WinStreak").Value > 0 Then
			Response.Write "<font color=""green"">+" & oRs.Fields("WinStreak").Value & "</font>"
		ElseIf oRs.Fields("LossStreak").Value > 0 Then
			Response.Write "<font color=""red"">-" & oRs.Fields("LossStreak").Value & "</font>"
		Else
			Response.Write "--"
		End If
		%>
		</td>
		<td align=center>
			<%
				if ors.Fields("ActiveMatches").Value = 0 then
					response.write "&nbsp;"
				else
					response.write ors.Fields("ActiveMatches").Value
				end if
			%></td>
		</tr>
		<%		
		ors.MoveNext
		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
	loop
	%>
	</table>
<% else %>
	<table align=center border=0 cellspacing=0 cellpadding=0 width=85% class="cssBordered">
	<tr BGCOLOR="#000000">
		<TH WIDTH=40 align=center>Rating</TH>
		<TH WIDTH=225 align=center>Team</TH>
		<TH WIDTH=100 align=CENTER>Record</TH>
		<TH width=100 align=center>Current Streak</TH>
	</tr>
	<TR>
		<TD Colspan=4 BGCOLOR="#000000">&nbsp;&nbsp;<i><font color=red>No unranked teams.</font></i></TD>
	</TR>
	</TABLE>
<% end if %>

	<BR>
	<table ALIGN=CENTER border=0 cellspacing=0 cellpadding=0 width=755>
	<TR>
	<TD VALIGN=TOp>
	<UL>
	<% If Not(IsNull(strLadderRules) or Len(Trim(strLadderRules)) = 0) Then %>
	<LI><B><a href="rules.asp?set=<%=Server.URLEncode(strLadderRules & "")%>" target="_blank"><%=strLadderName%> Rules</A></B>
	<% End If %>
	</UL>
	</td>
	</tr></table>

<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>