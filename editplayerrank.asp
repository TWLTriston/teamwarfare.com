<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Edit Player Rank"

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

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Edit Player Ranks") %>
<%
	if IsSysAdmin() then
		strsql="Select l.PlayerLadderID, PlayerladderName from TBL_playerladders l order by PlayerladderName"
	else
		strsql="Select l.PlayerLadderID, PlayerladderName from TBL_playerladders l, lnk_pL_a lnk where lnk.PlayerID = '" & session("PlayerID") & "' and l.playerladderid = lnk.playerladderid order by PlayerladderName"
	end if
	bgc=bgctwo
	%>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 ALIGN=CENTER WIDTH="400" BGCOLOR="#444444">
	<TR><TD>
		<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH="100%">
		<TR BGCOLOR="#000000">
			<TH>Choose a Ladder</TH>
		</TR>
	<%
	ors.open strsql, oconn
	if not (ors.eof and ors.bof) then
		do while not (ors.eof)
			response.write "<tr height=20 bgcolor=" & bgc & "><td>&nbsp;<a href=editplayerrank.asp?ladder=" & server.urlencode(ors.fields("PlayerLadderName").value) & ">" & Server.HTMLEncode(ors.fields("PlayerLadderName").value) & "</a></td></tr>"
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.movenext
		loop
	end if
	ors.close
	%>
	</TABLE>
	</TD></TR>
	</TABLE>
	<BR><BR>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 ALIGN=CENTER WIDTH="400" BGCOLOR="#444444">
	<TR><TD>
		<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH="100%">
		<form action=saveitem.asp method=post name=chooseplayer id=chooseplayer>
		<input type=hidden name=SaveType value=ChangePlayerRank>
		<TR BGCOLOR="#000000">
			<TH>Choose a Player</TH>
		</TR>
		<%
		if strLadder <> "" then
			strSQL = "SELECT PlayerHandle, PlayerLadderID, PPLLinkID, Rank FROM vPlayerLadder WHERE PlayerLadderName ='" & CheckString(strLadder) & "'"
			ors.open strsql, oconn
			if not (ors.eof and ors.bof) then
				Response.Write "<input type=hidden name=LadderID value=""" & oRS.Fields("PlayerLadderID") & """>"
				response.write "<tr BGCOLOR=" & bgcone & "><td align=center><select name=PPLLinkID>"
				do while not ors.eof
					response.write "<option value=" & ors.fields("PPLLinkID").value & ">" & ors.fields("PlayerHandle").value & " #" & ors.fields("Rank").value & "</OPTION>" & vbCrlF
					ors.movenext
				loop
				response.write "</select></td></tr>"
				response.write "<tr bgcolor=" & bgctwo & "><td align=center>New Rank: <input type=text class=bright name=NewRank id=newrank1></td></tr>"
				response.write "<tr bgcolor=" & bgcone & "><td align=center><input type=submit class=bright value='Confirm New Rank' style=""width:150"" name=submit1 value=submit1></td></tr>"
			End If
		end if
		%>
		</FORM>
		</TABLE>
		</TD></TR>
	</TABLE>
	<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>