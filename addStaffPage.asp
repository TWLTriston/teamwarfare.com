
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add a Staff Group"

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

If Not(IsSysAdminLevel2()) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim strVerbage, bIsEdit, strStaffGroupName, intStaffGroupID 
Dim strMethod
bIsEdit = cBool(Request.QueryString("IsEdit"))
If bIsEdit Then
	strVerbage = "Edit a game"
	strGameName = Request.QueryString("Game")
	strMethod = "Edit"

Else
	strVerbage = "Add a Staff Group"
	strMethod = "New"
End If
strPageTitle = strVerbage
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart(strVerbage)
%>
	<form name=frmAddStaffPage action=saveItem.asp method=post>
	<table align=center BACKGROUND="" BGCOLOR="#444444" CELLSPACING=0 CELLPADDING=0>
	<TR><TD>
	<table align=center CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
		<TR BGCOLOR="#000000"><TH COLSPAN=2><%=strVerbage%></TH></TR>
		<tr bgcolor=<%=bgcone%>><td align=right>Name:</td><td width=300>&nbsp;<INPUT id=StaffGroup name=StaffGroup style=" WIDTH: 250px" class=text value="<%=Server.HTMLEncode(strStaffGroup)%>"></td></tr>
<tr bgcolor=<%=bgctwo%>><td colspan=2 align=middle><INPUT id=submit1 name=submit1 type=submit value="Save Staff Group Information"></td></tr>
<input type=hidden name=SaveMethod value="<%=strMethod%>">
<input type=hidden value="<%=Server.HTMLEncode(intStaffGroupID)%>" name=GameID>
</TABLE>
</TD></TR>
</TABLE>
<input type=hidden name=SaveType value="StaffGroup">
</form>
<%
Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>