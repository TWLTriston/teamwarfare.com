<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Join New Competition"

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

Dim jType, strPlayerName, intPlayerID
jType= Request.QueryString("JoinType")
strPlayerName = Request.QueryString ("player")
Dim bShown

If Not(bSysAdmin OR Session("uName") = strPlayerName) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Set oRS2 = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Join a Ladder with " & Server.HTMLEncode(strPlayerName)) %>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
	<TR><TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2>
	<form name=frmLadderJoin action=saveItem.asp method=post>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=2>Choose a Ladder</TH>
	</TR>
	<%
	strSQL = "select playerID from tbl_players where PlayerHandle='" & CheckString(strPlayerName) & "'"
	ors.Open strsql, oconn
	if not (ors.EOF and ors.BOF) then
		intPlayerID=ors.Fields(0).Value
	end if
	ors.close
	strSQL = "SELECT playerladdername, playerladderid FROM tbl_playerladders WHERE Active = 1 AND playerladderid Not IN "
	strSQL = strSQL & " ( SELECT playerladderid FROM lnk_p_pL lnk, tbl_players p WHERE p.PlayerID = lnk.PlayerID AND p.PlayerHandle='" & CheckString(strPlayerName) & "' AND IsActive = 1) "
	strSQL = strSQL & " ORDER BY playerladdername"
	ors.Open strSQL, oConn

	bgc=bgcone
	if not (ors.EOF and ors.BOF) then
		bshown=false
		do while not ors.EOF
			%>
			<tr bgcolor=<%=bgc%>><td align=right><INPUT id=radio1 name=LadderToJoin type=radio class=borderless value="<%= Server.HTMLEncode(ors.Fields("playerLadderName").Value) %>"></td><td width=200 align=left><%= Server.htmlencode(ors.Fields("PlayerLAddername").Value) %></td></tr>
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			bshown = true
			ors.MoveNext
		loop
		if bshown then
			%>
			<tr bgcolor=<%=bgc%> height=30><td colspan=2 align=middle><INPUT id=submit1 name=submit1 type=submit value='Join Ladder' class=bright ></td></tr>
			<%
		end if
	end if	
	ors.close
	%>
	</table>
	<input id=hidden name=PlayerName type=hidden value="<%=strPlayerName%>">
	<input id=hidden name=PlayerID type=hidden value=<%=intPlayerID%>>
	<input id=hidden name=SaveType type=hidden value=PlayerLadderJoin>
	</form>

</td></tr>
</TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>
