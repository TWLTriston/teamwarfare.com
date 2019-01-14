<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Update TWL Menus"

Dim strSQL, oConn, oRS, oRs2, oRs3, oRs4
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")
Set oRS3 = Server.CreateObject("ADODB.RecordSet")
Set oRS4 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

if not(IsSysAdminLevel2()) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Update TWL Menus")
If Request.Querystring("s") = "1" Then
	Response.Write "-----------------------<br /><b>Menus are updated.</b><br />-----------------------<br /><br />"
End If
%>
Which menu do you want to update?<br />
<%

strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = 0 ORDER BY SortOrder, ShowMenuName "
oRs.Open strSQL, oConn
If Not(oRS.EOF AND oRs.BOF) Then
	Do While Not(oRs.EOF)
		%>
		<a href="updatemenus.asp?MenuID=<%=oRs.Fields("MenuID").Value%>&Menu=<%=Server.HTMLEncode(oRs.Fields("ShowMenuName").Value)%>"><%=oRs.FieldS("ShowMenuName").Value%></a><br />
		<%
		oRs.MoveNext
	Loop
End If
oRs.NextRecordSet
%>

<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>