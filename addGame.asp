<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add a Game"

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

Dim strVerbage, bIsEdit, strGameName, intForumID, strGameAbbr, intGameID, intDisputeForumID 
Dim strMethod
bIsEdit = cBool(Request.QueryString("IsEdit"))
If bIsEdit Then
	strVerbage = "Edit a game"
	strGameName = Request.QueryString("Game")
	strMethod = "Edit"
	strSQL = "select * from tbl_games where GameName='" & CheckString(strGameName) & "'"
	oRs.Open strSQL, oConn
	if not (ors.eof and ors.BOF) then
		intForumID = oRS.Fields("ForumID").Value
		intDisputeForumID = oRS.Fields("DisputeForumID").Value
		strGameAbbr = oRS.Fields("GameAbbreviation").Value
		strGameName = oRS.Fields("GameName").Value
		intGameID = oRS.Fields("GameID").Value
		strMethod="Edit"
	end if
	ors.Close
Else
	strVerbage = "Add a game"
	strMethod = "New"
End If
strPageTitle = strVerbage
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart(strVerbage)
%>
	<form name=frmAddGame action=saveItem.asp method=post>
	<table align=center BACKGROUND="" BGCOLOR="#444444" CELLSPACING=0 CELLPADDING=0>
	<TR><TD>
	<table align=center CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
		<TR BGCOLOR="#000000"><TH COLSPAN=2><%=strVerbage%></TH></TR>
		<tr bgcolor=<%=bgcone%>><td align=right>Name:</td><td width=300>&nbsp;<INPUT id=GameName name=GameName style=" WIDTH: 250px" class=text value="<%=Server.HTMLEncode(strGameName)%>"></td></tr>
		<tr bgcolor=<%=bgctwo%>><td align=right>Abbreviation:</td><td>&nbsp;<INPUT id=GameAbbreviation name=GameAbbreviation style=" WIDTH: 100px" class=text value="<%=Server.HTMLEncode(strGameAbbr)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Linked Forum:</td><td>&nbsp;<SELECT NAME=ForumID Class=text>
				<%
					strSQL = "SELECT tbl_forums.ForumID, tbl_forums.ForumName "
					strSQL = strSQL & " FROM tbl_category, tbl_forums "
					strSQL = strSQL & " WHERE tbl_forums.CategoryID = tbl_category.CategoryID "
					strSQL = strSQL & " AND tbl_category.CategoryOrder >= 0 "
					strSQL = strSQL & " ORDER BY tbl_forums.ForumName ASC"
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						Do While Not(oRS.EOF)
							Response.Write "<OPTION VALUE=""" & oRS.Fields("ForumID").Value & """ "
							If cStr(oRS.Fields("ForumID").Value  & "") = cStr(intForumID & "") Then
								Response.Write " SELECTED "
							End If
							Response.Write ">" & Server.HTMLEncode(oRS.Fields("ForumName").Value & "") & "</OPTION>" & vbCrLf
							oRs.MoveNext
						Loop					
					End If
					oRs.NextRecordset
					%>
					</SELECT></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Dispute Forum:</td><td>&nbsp;<SELECT NAME=DisputeForumID Class=text>
				<%
					strSQL = "SELECT tbl_forums.ForumID, tbl_forums.ForumName "
					strSQL = strSQL & " FROM tbl_category, tbl_forums "
					strSQL = strSQL & " WHERE tbl_forums.CategoryID = tbl_category.CategoryID "
					strSQL = strSQL & " AND tbl_category.CategoryOrder >= 0 "
					strSQL = strSQL & " ORDER BY tbl_forums.ForumName ASC"
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						Do While Not(oRS.EOF)
							Response.Write "<OPTION VALUE=""" & oRS.Fields("ForumID").Value & """ "
							If cStr(oRS.Fields("ForumID").Value  & "") = cStr(intDisputeForumID & "") Then
								Response.Write " SELECTED "
							End If
							Response.Write ">" & Server.HTMLEncode(oRS.Fields("ForumName").Value & "") & "</OPTION>" & vbCrLf
							oRs.MoveNext
						Loop					
					End If
					oRs.NextRecordset
					%>
					</SELECT></td></tr>
<tr bgcolor=<%=bgctwo%>><td colspan=2 align=middle><INPUT id=submit1 name=submit1 type=submit value="Save Game Information"></td></tr>
<input type=hidden name=SaveMethod value="<%=strMethod%>">
<input type=hidden value="<%=Server.HTMLEncode(intGameID)%>" name=GameID>
</TABLE>
</TD></TR>
</TABLE>
<input type=hidden name=SaveType value="Games">
</form>
<%
Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>