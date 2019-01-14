<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle
strPageTitle = "TWL: " & Replace(Request.Querystring("tournament"), """", "&quot;") & " Tournament"

Dim strSQL, oConn, oRS, oRS1, oRS2, RSr
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

DIM Tournament, divID

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%

tournament = Request.QueryString("tournament")
If (tournament = "Call of Duty Tourney") THen
	Tournament = "Answer the Call"
End if

if tournament = "" then
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
end if

strSQL = "SELECT * FROM tbl_tournaments WHERE TournamentName='" & replace(tournament, "'", "''") & "'"
Set oRS = oConn.Execute(strSQL)
if oRS.eof and oRS.bof then
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
end if

Dim TournamentID, TournamentName, blnSignUp, blnLocked
TournamentID = oRS.fields("TournamentID").value
TournamentName = oRS.fields("TournamentName").value
blnsignUp = oRS.Fields("SignUp").Value
blnLocked = oRS.Fields("Locked").Value
oRs.NextRecordSet
%>
<!-- top box -->
<%
If TournamentName = "Operation Triple Threat" Then 
	Call ContentStart(TournamentName & " Tournament Sponsors")
	%>
<div align="center"> 
<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" 
	codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" WIDTH="728" HEIGHT="90" id="sc_banner" ALIGN=""> 
<PARAM NAME="movie" VALUE="/images/allinwonder9700_728x90.swf?clickTag=http://www.ati.com"> 
<PARAM NAME="quality" VALUE="high"> 
<PARAM NAME="bgcolor" VALUE="#000000"> 
<EMBED src="/images/allinwonder9700_728x90.swf?clickTag=http://www.ati.com" quality="high" bgcolor="#FFFFFF"  WIDTH="728" HEIGHT="90" NAME="rvs_banner" ALIGN="" TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/go/getflashplayer"></EMBED> 
</OBJECT> 
<br /><br />
</div>
	<center><a href="http://www.alienware.com/index.asp?from=Ravenshield 3 promo:01_banner_468x60"><img src="http://www.alienware.com/main/affiliate_pages/banners/01_banner_468x60.gif" width="468" height="60" border="0" alt="Alienware - Ultimate Gaming PC"></a><br />
	<br />
	</center>
	<%
	Call ContentEnd()
End If
If TournamentName = "Answer the Call" Then
	Call ContentStart(TournamentName & " Tournament Sponsors")
	%>
	<center>
	<OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
 codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0"
 WIDTH="468" HEIGHT="60" id="gearstore_banner" ALIGN="">
 <PARAM NAME=movie VALUE="/images/cod/gearstore_banner.swf">
  <PARAM NAME=quality VALUE=high> 
  <PARAM NAME=bgcolor VALUE=#FFFFFF> 
  <EMBED src="/images/cod/gearstore_banner.swf" quality=high bgcolor=#FFFFFF  
  	WIDTH="468" HEIGHT="60" NAME="gearstore_banner" ALIGN="" 
  	TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/go/getflashplayer"></EMBED>
  </object>
</center>
	<%
	Call ContentEnd()
	
End If
%>
<% Call ContentStart("Complete team list for " & Server.HTMLEncode(TournamentName)  & " Tournament")%>
<table border="0" cellspacing="0" cellpadding="4" width="90%" align="center">
<%
strSQL = "SELECT t.TeamName, t.TeamTag, l.TMLinkID FROM lnk_t_m l INNER JOIN tbl_teams t ON t.TeamID = l.TeamID  "
strSQL = strSQL & " WHERE TournamentID = '" & TournamentID & "'"
strSQL = strSQL & " ORDER BY TeamName ASC"
ors.open strsql, oconn
if not(ors.eof and ors.bof) then
	do while not(ors.eof)
		%>
		<tr>
			<td><a href="/viewteam.asp?team=<%=Server.URLEncode(oRs.fields("TeamName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("teamName").Value & " - " & ors.fields("teamtag").value)%></a> <%
			If bSysAdmin THen %>
				-- <a href="/saveitem.asp?SaveType=TournamentRemove&TMLinkID=<%=oRS.Fields("TMLinkID").Value%>&Tournament=<%=Server.URLEncode(TournamentName & "")%>">remove from tournament</a>
				<%
			End If
			%>
		</tr>
		<%
		ors.movenext
	loop
else
	%>
	<tr>
		<td>No teams have signed up yet.</td>
	</tr>
	<%	
end if
ors.nextrecordset
%>
</table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
