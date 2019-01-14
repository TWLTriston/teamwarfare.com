<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Match Communications"

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

Dim cMode, commid, comms, savetype
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Add a match communication")
	cMode = Request.QueryString("mode")
	if  cMode = "edit" then
		savetype= "EditLeagueCommunications"
		strSQL = "select * from tbl_league_comms where Leaguecommid=" & Request.QueryString("leaguecommid")
		oRs.Open strSQL, oConn
		If Not(oRS.EOF AND oRS.BOF) Then
			comms=ors.Fields("Comms").Value
			commid=Request.QueryString("leaguecommid")
		Else
			oRS.Close
			oConn.Close
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear 
			Response.Redirect "TeamLeagueAdmin.asp?ladder=" & server.urlencode(Request("League")) & "&team=" & server.urlencode(Request("Team"))
		End IF
		ors.Close  
	end if
	if cMode = "add" then
		savetype = "AddLeagueCommunications"
	end if 
	if cMode = "delete" then
		savetype="DeleteLeagueCommunications"
	end if
	
%>
<form name=frmComms method=post action=saveitem.asp>
<input type=hidden name=SaveType value=<%=savetype%>>
<input type=hidden name=MatchID value=<%=Request.QueryString("matchid")%>>
<input type=hidden name=commid value=<%=commid%>>
<input type=hidden name=commauthor value="<%=Server.HTMLEncode(session("uName"))%> - ( <%=Server.HTMLEncode(request.querystring("tag"))%> )">
<input type=hidden name=league value="<%=request("League")%>">
<input type=hidden name=team value="<%=request("Team")%>">
<input type=hidden name=commdate value=<%=date%>>
<input type=hidden name=commtime value=<%=time%>>
<table align=center width=97%>
<tr height=25 bgcolor=<%=bgcone%>><td align=center><p class=small><b><%=replace(savetype, "_", " ")%></b></p></td></tr>
<tr height=200 bgcolor=<%=bgctwo%>><td align=center><Textarea name=comms cols=60 rows=10 wrap=hard><%=Server.HTMLEncode(comms)%></textarea></td></tr>
<tr height=30 bgcolor=<%=bgcone%>><td align=center><input type=submit name=submit1 value=submit class=bright></td></tr></table>
</form>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

