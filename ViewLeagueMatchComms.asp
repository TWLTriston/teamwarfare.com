<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Team Ladder Administration"

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

Dim intMatchID
intMatchID = Request.QueryString("MatchID")
If Not(bSysAdmin or IsAnyLeagueAdmin()) Then
	oConn.Close 
	Set oRS = Nothing
	Set oConn = Nothing
	Set oRs2 = Nothing
	response.clear
	response.redirect "/errorpage.asp?error=3"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->
<%
Call ContentStart("")

	Response.Write "<a name=""matchcomms""></a><br><br><center><font size=2><b><u>Match Communications</u></b></center>"
	strSQL = "select * from tbl_league_comms where LeagueMatchID='" & intMatchID & "' order by LeagueCommID desc"
	'Response.Write strSQL
	ors.Open strSQL, oconn
	%>
	<table align=center width=580 border=0 cellspacing=0 cellpadding=1>
	<%
	bgc=bgcone
	if not (ors.EOF and ors.bof) then
		do while not ors.EOF
			Response.Write "<tr bgcolor="& bgc & "><td colspan=2><hr></td></tr>" & vbCrLf
			Response.Write "<tr bgcolor="& bgc & "><td>Author: <b>" & ors.Fields(3).Value & " - Posted: " & FormatDateTime(ors.Fields(2).Value, 0) & "</td></tr>" & vbCrLf
			Response.write "<tr bgcolor="& bgc &"><td colspan=2>" & Replace(ors.Fields(4).Value, chr(10), "<br />") & "</td></tr>" & vbCrLf
			if bgc = bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.MoveNext
		loop
	end if
	%>
	</table>
	<%
	ors.Close 
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>