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

Dim intMatchID
Dim cMode, mrid, rant, savetype

If Not(Session("LoggedIn")) Then
	Response.Clear
	Response.Redirect "/errorpage.asp?error=2"	
End If
	
If Not(HasForumAccess()) Then
	Response.Clear
	Response.Redirect "/errorpage.asp?error=34"	
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Add a match rant")
	cMode = Request.QueryString("cmode")
	mrid=Request.QueryString("MRID")
	if  cMode = "edit" then
		savetype= "Edit_Match_Rant"
		strSQL = "select * from tbl_match_rants where MRID=" & Request.QueryString("MRID")
		oRs.Open strSQL, oConn
		If Not(oRS.EOF AND oRS.BOF) Then
			rant=ors.Fields("Rant").Value
		Else
			oRS.Close
			oConn.Close
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear 
			Response.Redirect "viewMatch.asp?Ladder=" & server.urlencode(Request("Ladder")) & "&MatchID=" & server.urlencode(Request("MatchID"))
		End IF
		ors.Close  
	end if
	if cMode = "add" then
		savetype = "Add_Match_Rant"
	end if 
%>
<form name=frmComms method=post action=saveitem.asp>
<input type=hidden name=SaveType value=<%=savetype%>>
<input type=hidden name=MatchID value=<%=Request.QueryString("MatchID")%>>
<input type=hidden name=MRID value=<%=MRID%>>
<input type=hidden name=Ranter value="<%=Server.HTMLEncode(session("PlayerID"))%>">
<input type=hidden name=Ladder value="<%=Server.HTMLEncode(request("Ladder") & "")%>">
<table align=center cellspacing="0" cellpadding="0" bgcolor="#444444">
<tr>
<td>
<table align=center cellspacing="1" cellpadding="4" bgcolor="#444444">

<tr height=25 bgcolor=<%=bgcone%>><td align=center><p class=small><b><%=replace(savetype, "_", " ")%></b></p></td></tr>
<tr height=200 bgcolor=<%=bgctwo%>><td align=center><Textarea name=Rant cols=60 rows=10><%=Server.HTMLEncode(rant)%></textarea></td></tr>
<tr height=30 bgcolor=<%=bgcone%>><td align=center><input type=submit name=submit1 value=submit class=bright></td></tr>
</table>
</td></tr>
</table>
</form>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

