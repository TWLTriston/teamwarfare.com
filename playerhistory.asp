<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("player"), """", "&quot;") 

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strPlayerName, intPlayerID, bDefender
Dim strResult, strEnemyName, map1, map1usscore, map1themscore, map1ft
strPlayerName = Request("Player")

If Len(strPlayerName) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
%><!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Match History for " & Server.HTMLEncode(strPlayerName)) %> 

<%	
strSQL = "select PlayerID from tbl_players where PlayerHandle='" & CheckString(strPlayerName) & "'"
oRs.Open strSQL, oConn
if not (ors.EOF and ors.BOF) then
	intPlayerID = ors.Fields(0).Value 
else
	oRS.Close 
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
end if
ors.NextRecordset 
%>

	<table border=0 width=760 cellspacing=0 cellpadding=0 ALIGN=CENTER BGCOLOR="#444444">
	<TR><TD>
	<table border=0 width=100% cellspacing=1 cellpadding=2>
	<tr bgcolor="#000000">
		<TH>Ladder</TH>
		<TH>Opponent</TH>
		<TH>Result</TH>
		<TH>Date</TH>
		<TH>Defender</TH>
	</tr>

<%
	bgc=bgctwo
	strsql="SELECT * FROM vPlayerHistory WHERE (WinnerName = '" & CheckString(strPlayerName) & "' OR LoserName = '" & CheckString(strPlayerName) & "') ORDER BY LadderName ASC, MatchDate DESC"
	ors.Open strSQL, oconn
	if not (ors.eof and ors.BOF) then
		do while not ors.EOF
			If ors.Fields("WinnerName") = strPlayerName Then
				strEnemyName = ors.Fields("LoserName").Value 
				strResult = "Win"
				If oRS.Fields("WinnerDefending").Value Then
					bDefender = True
				End If
			Else
				strEnemyName = ors.Fields("WinnerName").Value 
				strResult = "Loss"
				If oRS.Fields("WinnerDefending").Value Then
					bDefender = True
				End If
			End If

			If bDefender Then
				map1=ors.fields("MatchMap1").value
				map1usscore=ors.fields("MatchMap1DefenderScore").value
				map1themscore=ors.fields("MatchMap1AttackerScore").value
				map1ft=ors.fields("map1forfeit").value
			else
				map1=ors.fields("MatchMap1").value
				map1themscore=ors.fields("MatchMap1DefenderScore").value
				map1usscore=ors.fields("MatchMap1AttackerScore").value
				map1ft=ors.fields("map1forfeit").value
			end if
			
			%>
			<tr bgcolor=<%=bgctwo%>><TD><A href="viewplayerladder.asp?ladder=<%=server.URLEncode(oRs.Fields("LadderName").Value )%>"><%=Server.HTMLEncode(oRs.Fields ("LadderName").Value )%></A></TD>
			<td><a href="viewplayer.asp?player=<%=server.urlencode(strEnemyName)%>"><%=Server.HTMLEncode (strEnemyName)%></a></td>
			<td><%=strResult%></td><td><%=ors.Fields("MatchDate").Value%></td><td align=center>
			<% if bDefender then
				response.write Server.HTMLEncode (strPlayerName)
			   else
			   	response.write Server.HTMLEncode(strEnemyName)
			   end if
			   %>
			   </td>
			</tr>
			<%
			%>
			<tr>
			<td BGCOLOR="#000000"><img src="/images/spacer.gif" height="1"></td>
			<td align=left height=20 colspan=3 bgcolor=<%=bgcone%>>
			<%
			If oRs("MatchForfeit") = 1 Then
				Response.Write "Admin Forfeit"
			Else
			%>
				&nbsp;<b><%=Server.HTMLEncode(map1 & "")%>:</b> <%=map1usscore%> - <%=map1themscore%><%
				if map1ft then
					response.write " by forfeit"
				end if
			End If
			%></td>
			<TD BGCOLOR="#000000">&nbsp;</TD>
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
	ors.Close 
%>
	</table>
	</TD></TR>
	</TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>