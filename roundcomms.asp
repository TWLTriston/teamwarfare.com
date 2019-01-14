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

<% Call ContentStart("Add a match communication") %>
<%
	cMode = Request.QueryString("mode")
	if  cMode = "edit" then
		savetype= "Edit_Communications"
		strSQL = "select count(*) from tbl_round_comm where commid=" & Request.QueryString("commid")
		oRs.Open strSQL, oConn
		if ors.Fields(0) = 0 then
			oRS.Close
			oConn.Close
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "TeamTouramentAdmin.asp?tournament=" & server.urlencode(session("CurrentTournament")) & "&team=" & server.urlencode(session("CurrentTeam"))
		end if
		ors.Close 
		strSQL = "select * from tbl_round_comm where commid=" & Request.QueryString("commid")
		oRs.Open strSQL, oConn
		comms=ors.Fields("Comms").Value
		commid=Request.QueryString("commid")
		ors.Close  
	end if
	if cMode = "add" then
		savetype = "Add_Communications"
	end if 
	if cMode = "delete" then
		savetype="Delete_Communications"
	end if
	
%>
<form name=frmComms method=post action=/tournament/savetournament.asp>
<input type=hidden name=SaveType value=<%=savetype%>>
<input type=hidden name=RoundsID value=<%=Request.QueryString("RoundsID")%>>
<input type=hidden name=commid value=<%=commid%>>
<input type=hidden name=commauthor value="<%=Server.HTMLEncode(session("uName"))%> - ( <%=Server.htmlencode(request.querystring("tag"))%> )">
<input type=hidden name=commdate value=<%=date%>>
<input type=hidden name=commtime value=<%=time%>>
<table align=center width=97%>
<tr height=25 bgcolor=<%=bgcone%>><td align=center><p class=small><b><%=replace(savetype, "_", " ")%></b></p></td></tr>
<tr height=200 bgcolor=<%=bgctwo%>><td align=center><Textarea name=comms cols=60 rows=10 wrap=hard><%=Server.HTMLEncode(comms)%></textarea></td></tr>
<tr height=30 bgcolor=<%=bgcone%>><td align=center><input type=submit name=submit1 value=submit class=bright></td></tr></table>
</form>

        </td>
      </tr>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

