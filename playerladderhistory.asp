<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Ladder Match History"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strLadderName
strLadderName = Request.QueryString("Ladder")

Dim strResult, strEnemyName, bDefender, strTeamName
Dim map1, map1usScore, Map1ThemScore, map1OT, map1FT
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Match History for " & Server.HTMLEncode(strLadderName) & " Ladder") %>
	<table border=0 width=760 cellspacing=0 cellpadding=0 ALIGN=CENTER BGCOLOR="#444444">
	<TR><TD>
	<table border=0 width=100% cellspacing=1 cellpadding=2>
	<tr bgcolor="#000000">
		<TH>Defender</TH>
		<TH>Attacker</TH>
		<TH>Match Date</TH>
		<TH>Winner</TH>
	</TR>
<%	bgc=bgctwo
	strSQL="select top 20 * from vPlayerHistory where LadderName='" & CheckString(strLadderName) & "' AND MatchForfeit = 0 order by matchdate desc"
	ors.Open strSQL, oconn
	if not (ors.eof and ors.BOF) then
		do while not ors.EOF
			bDefender = False
			If oRS.Fields("WinnerDefending").Value Then
				strEnemyName = ors.Fields("LoserName").Value 
				strTeamName = oRS.Fields("WinnerName").Value 
				bDefender = True
			Else
				strEnemyName = ors.Fields("WinnerName").Value 
				strTeamName = oRS.Fields("LoserName").Value 
			End If

			map1 = ors.fields("MatchMap1").value
			map1usscore =ors.fields("MatchMap1DefenderScore").value
			map1themscore=ors.fields("MatchMap1AttackerScore").value
			map1ft=ors.fields("map1forfeit").value
			%>
			<tr bgcolor=<%=bgctwo%>>
			<td><a href="viewteam.asp?team=<%=server.urlencode(strTeamName)%>"><%=Server.HTMLEncode(strTeamName)%></a></td>
			<td><a href="viewteam.asp?team=<%=server.urlencode(strEnemyName)%>"><%=Server.HTMLEncode(strEnemyName)%></a></td>
			<td><%=ors.Fields("MatchDate").Value%></td><td align=center>
			<% if bDefender then
				response.write Server.HTMLEncode(strTeamName)
			   else
			   	response.write Server.HTMLEncode(strEnemyName)
			   end if
			   %>
			   </td>
			</tr>
			<tr BGCOLOR="#000000">
			<td><img src="/images/twl/spacer.gif" height="1"></td>
			<td align=left colspan=2 bgcolor=<%=bgcone%>>
			&nbsp;<b><%=Server.HTMLEncode(map1)%>:</b> <%=map1usscore%> - <%=map1themscore%><%
			if map1ot then
				response.write " in OT"
			end if
			if map1ft then
				response.write " by forfiet"
			end if
			%>
			</td>
			<TD>&nbsp;</TD>
			</tr>
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.MoveNext
		loop
	end if
	ors.NextRecordset 
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
	