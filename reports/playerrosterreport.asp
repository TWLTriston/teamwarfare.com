<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "Player Roster Report"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Player Roster Report")

If bSysAdmin Then
	strSQL = "SELECT PlayerLadderName from tbl_playerLadders order by PlayerLadderName"
Else
	strSQL = "SELECT PlayerLadderName from tbl_playerLadders pl, lnk_pl_a lnk WHERE lnk.PlayerID = '" & Session("PlayerID") & "' AND lnk.PlayerLadderID = pl.PlayerLadderID ORDER by PlayerLadderName	"
End if
		
ors.open strsql, oconn
if not(ors.eof and ors.bof) then
	Response.Write "<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 ALIGN=CENTER>"
	do while not(ors.eof)
		Response.Write "<TR>"
		Response.Write "<TD ALIGN=CENTER><p class=small><B><a href=""playerrosterreport.asp?ladderName=" & Server.URLEncode(oRS("PlayerLadderName")) & """>" & oRS("PlayerLadderName") & "</A></B></TD>"
		Response.Write "</TR>"
		ors.movenext
	loop
	Response.write "</TABLE>"
end if
ors.close

Call ContentEnd()
bgc = bgcone
If Request.QueryString("LadderName") <> "" Then 
	Call ContentStart(Request.QueryString("LadderName") & " Roster Report")
		strSQL = "SELECT tbl_players.PlayerHandle, lnk_p_pl.*, tbl_playerLadders.PlayerLadderID "
		strSQL = strSQL & " FROM tbl_players, lnk_p_pl, tbl_PlayerLadders "
		strSQL = strSQL & " WHERE tbl_players.playerid = lnk_p_pl.playerid "
		strSQL = strSQL & " AND lnk_p_pl.PlayerLadderID = tbl_playerLadders.PlayerLadderID "
		strSQL = strSQL & " AND tbl_playerLadders.PlayerLadderName = '" & CheckString(Request.QueryString("LadderName")) & "'"
		strSQL = strSQL & " AND lnk_p_pl.IsActive = 1 "
		strSQL = strSQL & " ORDER BY lnk_p_pl.LastLogin ASC "
		oRS.Open strSQL, oConn
		Response.Write "<TABLE BORDER=0 WIDTH=75% CELLSPACING=0 CELLPADDING=0 BGCOLOR=""#444444"" ALIGN=CENTER><TR><TD>"
		Response.Write "<TABLE BORDER=0 WIDTH=100% CELLSPACING=1 CELLPADDING=2 ALIGN=CENTER>"
		Response.Write "<TR BGCOLOR=""#000000""><TH>Player Name</TH><TH>Forfiets</TH><TH>Last Login Time</TH><TH>Join Date</TH></TR>"
		If Not(oRS.eof and oRS.bof) Then
			Do While Not(oRS.EOF)
				Response.Write "<TR BGCOLOR=""" & bgc & """>"
				Response.Write "<TD><A HREF=""/viewplayer.asp?player=" & server.URLEncode(oRS.Fields("PlayerHandle").Value ) & """>" & oRS.Fields("PlayerHandle").Value & "</A></TD>"
				Response.Write "<TD>" & oRS.Fields ("forfeits").Value & "</TD>"
				If oRS.Fields("LastLogin").Value  <> "" Then
					If DateDiff("d", Date(), oRS.Fields("LastLogin").Value) > 14 Then
						Response.Write "<TD><FONT STYLE=""color:FF0000;""><B>" & FormatDateTime(oRS.Fields("LastLogin").Value , vbShortDate) & "</B></FONT></TD>"
					Else
						Response.Write "<TD>" & FormatDateTime(oRS.Fields("LastLogin").Value , vbShortDate) & "</TD>"
					End If
				Else
					Response.Write "<Td><FONT STYLE=""color:FF0000;""><B>Never Logged In</B></FONT></TD>"
				End If
				If Not (IsNull(oRS.Fields("JoinDate").Value )) Then
					Response.Write "<TD>" & FormatDateTime(oRS.Fields("JoinDate").Value , vbShortDate) & "</TD>"
				Else
					Response.Write "<TD> &nbsp;</TD>"
				End If
				If bgc = bgcone then
					bgc = bgctwo
				else
					bgc = bgcone
				end if
				oRS.MoveNext
			Loop		
		Else
			Response.Write "<TR BGCOLOR=""#000000""><TD COLSPAN=5><I>No players on this ladder.</TD></TR>"
		End If
		Response.Write "</TABLE></TD></TR></TABLE>"
		oRS.Close	
	Call ContentEnd()
End If
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>