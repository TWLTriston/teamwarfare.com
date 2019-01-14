<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: News Desk"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin Or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim strHeadline
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("News Desk") %>
<table width=760 border=0 cellspacing=0 cellpadding=0 align=center BGCOLOR="#444444">
<TR><TD>
<table width=100% border=0 cellspacing=1 cellpadding=2 align=center>
<tr BGCOLOR="#000000"><TH colspan=3>Headlines</th></tr>
<TR BGCOLOR="#000000"><TD ALIGN=CENTER COLSPAN=3><A HREF="/newspage.asp?action=Add">Add News</A></TD>
<tr bgcolor="#000000"><th>Game</th><th>News</th><th>Modify</th></tr>
<%
	strSQL = "SELECT TOP 30 NewsID, NewsAuthor, NewsHeadline, NewsDate, GameName, NewsType from tbl_News LEFT OUTER JOIN tbl_games ON tbl_news.NewsType = tbl_games.GameID ORDER BY NewsID desc"
	oRs.Open strSQL, oConn
	bgc = bgctwo
	if not (ors.eof and ors.bof) then
		do while not ors.EOF 
			strHeadline = "<tr bgcolor=" & bgc & "><td>"
			If oRS.Fields("NewsType").Value = 0 Then
				strHeadline = strHeadline & "Announcement"
			Else
				strHeadline = strHeadline & oRS.Fields("GameName").Value
			End If
			strHeadline = strHeadline & "</td><td><b>" & Server.HTMLEncode(ors.Fields("NewsHeadline").Value) & "</b> (" & ors.Fields("NewsAuthor").Value & " by " & Server.HTMLEncode(ors.Fields(3).Value) & ")</td><td align=right><a href=newspage.asp?action=Edit&newsID=" & ors.Fields(0).Value & ">Edit</a> - <a href=newspage.asp?action=Delete&newsID=" & ors.Fields(0).Value & ">Delete</a></td></tr>"
			Response.Write strHeadline
			ors.MoveNext 
			if bgc = bgcone then
				bgc = bgctwo
			else
				bgc = bgcone
			end if 
		loop
	end if
	ors.close
%>
</table></TD></TR>
</TABLE>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

