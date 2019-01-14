<%
Option Explicit

Dim strPageTitle
strPageTitle = "TWL: Access Procected Forum"

Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Dim oConn, oRS, strSQL
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open Application("ConnectStr")

Set oRS = Server.CreateObject ("ADODB.RecordSet")

Dim strHeaderColor, strHighlight1, strHighlight2
Dim strTableBorder, strTableBG
strHeaderColor	= Application("HeaderColor")
strHighlight1	= bgcone
strHighlight2	= bgctwo
strTableBorder	= "#444444"
strTableBG		= "#000000"

Dim intForumID, strForumName, strUserName, strUserID, blnForumAdmin
Dim intThreadID

Call CheckCookie()
Dim blnLoggedIn, blnSysAdmin, blnForumModerator, bSysAdmin, bAnyLadderAdmin, bLoggedIn
blnLoggedIn = Session("LOggedIn")
blnSysAdmin = IsSysAdmin()
bSysAdmin = blnSysAdmin
Dim strForumPassword, intCategoryID, strCategoryName

If Not(blnLoggedIn) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Call AddError("You must be logged in before accessing protected forums.")
	Response.Clear
	Response.Redirect "default.asp"
End If

intThreadID = Request.QueryString("ThreadID")
intForumID = Request.QueryString("ForumID")
strSQL = "SELECT ForumName, ForumPassword, CategoryID, CategoryName = (SELECT CategoryName FROM tbl_category WHERE tbl_category.CategoryID = tbl_forums.CategoryID) FROM tbl_forums WHERE ForumID='" & intForumID & "'"
oRS.Open strSQL, oConn
If Not(oRS.EOF AND oRS.BOF) Then
	strForumName = oRS.Fields("ForumName").Value
	strForumPassword = oRS.Fields("ForumPassword").Value 
	intCategoryID = oRS.Fields("CategoryID").Value 
	strCategoryName = oRS.Fields("CategoryName").Value
Else
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Call AddError("Invalid forum id passed.")
	Response.Clear
	Response.Redirect "default.asp"
End If
oRS.NextRecordset 

blnForumModerator = IsForumModerator(intForumID)
intThreadID = Request.QueryString("ThreadID")
If blnSysAdmin Or blnForumModerator Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/forums/forumdisplay.asp?forumid=" & forumid
End If
%>
<!-- #include file="../include/i_funclib.asp" -->
<!-- #include file="../include/i_header.asp" -->
<% Call ContentStart("")%>
<table border=0 cellspacing=0 cellpadding=0 width=97%>
<% Call ShowErrors("<TR><TD>", "</TD></TR>") %>
<tr>
	<td>
		<table BORDER="0" cellspacing="0" cellpadding="0" width="100%">
		<tr>
			<td CLASS="pageheader"><%
			Response.Write "<a href=""default.asp"">TWL</A>"
			Response.Write " &raquo; "
			Response.Write "<a href=""default.asp#Category" & intCategoryID & """>" & strCategoryName & "</A>"
			Response.Write " &raquo; "
			Response.Write "<a href=""forumdisplay.asp?forumid=" & intForumID & """>" & strForumName & "</A>"
			Response.Write " &raquo; "
			Response.Write " Access Protected Forum"
			%> </td>
		</tr>
		</table>
	</td>
</tr>
<tr valign=center>
	<td align=center>
        <table border=0 cellspacing=0 cellpadding=0 bgcolor="<%=strTableBorder%>">
		<form id=enterpassword name=enterpassword action=forumengine.asp method=post>
		<input type=hidden name=ForumID value="<%=intForumID%>">
		<input type=hidden name=savetype value="GrantAccess">
		<input type=hidden name=ThreadID value="<%=intThreadID%>">
        <TR>
			<TD>
			<table border=0 cellspacing=1 cellpadding=4 width=100%>
			<TR>
				<TH COLSPAN=2 CLASS="category" BGCOLOR="<%=strTableBG%>">Access Protected Forum</TH>
			</TR>
			<TR bgcolor=<%=strHighlight1%>>
				<TD valign=middle align="right"><B>Password:</B></td>
				<TD><input name=ForumPassword id=ForumPassword type="Password" size=30 maxlength=50></td>
			</TR>
			<TR bgcolor=<%=strHighlight2%>>
				<TD colspan=2 align=center><input type=submit id=submit name=submit value="Grant Access"></td>
			</TR>
			</table>
			</TD>
		</TR>
		</form>	
		</TABLE>		
	</td>
</tr>
</TABLE>
<% Call ContentEnd() %>
<!-- #include file="../include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>

