<%
Option Explicit

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle, strMode, intThreadID

strMode = Trim(Request.QueryString("mode"))
If Len(strMode) = 0 Then
	strMode = "Add"
End If
strMode = uCase(Left(strMode, 1)) & lcase(Right(strMode, Len(strMode) - 1))

If strMode <> "Edit" AND strMode <> "Add" Then
	Call AddError("Unable to determine intended action. Please check linking url and try again.")
	Response.Clear
	Response.Redirect "/default.asp"
End If

If strMode = "Edit" Then
	intThreadID = Trim(Request.QueryString("threadid"))
	If Not(IsNumeric(intThreadID)) Or Len(intThreadID) = 0 Then
		Call AddError("Invalid thread id passed.")
		Response.Clear
		Response.Redirect "/default.asp"
	End If
End If

strPageTitle = "TWL: " & strMode & " Thread"
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Dim oConn, oRS, strSQL
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open Application("ConnectStr")

Set oRS = Server.CreateObject ("ADODB.RecordSet")

Dim intForumID, strForumName, strForumPassword
Dim intThreadAuthorID,strThreadAuthorName
Dim strThreadSubject, strThreadBody, blnThreadSignature, strSigChecked
Dim blnLoggedIn, blnSysAdmin
Dim blnForumAccess, blnForumModerator
Dim intCategoryID, strCategoryName
Dim blnForumLocked, blnThreadLocked, intReturnCode
Dim strThreadAuthorTitle, strThreadAuthorSignature

if uCase(Request.Cookies("PerPage")("ShowSig")) = "Y" Then
	blnThreadSignature  = True
End if

intForumID = Request("ForumID")
If Len(intForumID) = 0 OR Not(IsNumeric(intForumID)) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Call AddError("Invalid forum id. Please check linking url.")
	Response.Clear
	Response.Redirect "default.asp"
End If

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin, bForumAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

blnLoggedIn = Session("Loggedin")
blnSysAdmin = bSysAdmin

If Not(blnLoggedIn) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Call AddError("You must be logged in before you can post.")
	Response.Clear
	Response.Redirect "default.asp"
end if

blnForumAccess = HasForumAccess()
If Not(blnForumAccess) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Call AddError("Your posting priviledge has been revoked.")
	Response.Clear
	Response.Redirect "default.asp"
End If

strsql = "SELECT ForumName, ForumPassword, ForumLocked, CategoryID, CategoryName = (SELECT CategoryName FROM tbl_category WHERE tbl_category.CategoryID = tbl_forums.CategoryID) FROM tbl_forums WHERE ForumID='" & intForumID & "'"
oRS.Open strSQL, oConn
If Not(oRS.EOF AND oRS.BOF) Then
	strForumName = oRS.Fields("ForumName").Value
	strForumPassword = oRS.Fields("ForumPassword").Value 
	intCategoryID = oRS.Fields("CategoryID").Value 
	strCategoryName = oRS.Fields("CategoryName").Value  
	blnForumLocked = oRS.Fields("ForumLocked").Value 
Else
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Call AddError("Invalid forum id. Please check linking url.")
	Response.Clear
	Response.Redirect "default.asp"
End If
oRS.NextRecordset

blnForumModerator = IsForumModerator(intForumID)

Dim intTimeZoneDifference, strDate, strTime
Dim strCurrentTime, strCurrentDate
Dim strDateMask, bln24HourTime, dtmThreadPostTime
Dim blnThreadEdited, strThreadEditBy, dtmThreadEditTime
blnThreadEdited = False

intTimeZoneDifference = Session("intTimeZoneDifference")
strDateMask = "MM-DD-YYYY"
bln24HourTime = False

If strMode = "Edit" Then
	strsql = "SELECT *, PlayerTitle, PlayerSignature FROM tbl_Threads INNER JOIN tbl_players ON ThreadAuthorID = PlayerID WHERE ThreadID='" & intThreadID & "'"
	oRS.Open strSQL, oConn 
	If Not(oRS.EOF and oRS.BOF) Then
		intForumID					= oRS.Fields("ForumID").Value
		strThreadSubject			= oRS.Fields("ThreadSubject").Value
		strThreadBody				= oRS.Fields("ThreadBody").Value
		intThreadAuthorID			= oRS.Fields("ThreadAuthorID").Value
		blnThreadSignature			= cBool(oRS.Fields("ThreadSignature").Value)
		blnThreadLocked				= oRS.Fields("ThreadLocked").Value 
		strThreadAuthorName			= oRS.Fields("ThreadAuthorName").Value 
		strThreadAuthorTitle		= oRS.Fields("PlayerTitle").Value 
		dtmThreadPostTime			= oRS.Fields("ThreadPostTime").Value 
		strThreadAuthorSignature	= oRS.Fields("PlayerSignature").Value 
	End If
	oRS.Close 

	blnThreadEdited = True
	strThreadEditBy = Session("uName")
	dtmThreadEditTime = Now()

	If Not(blnSysAdmin OR blnForumModerator) AND cStr(intThreadAuthorID & "") <> cStr(Session("PlayerID") & "") Then
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Call AddError("You do not have access to edit other people's posts.")
		Response.Clear
		Response.redirect "default.asp"
	End If
Else
	strSQL = "SELECT PlayerTitle, PlayerSignature FROM tbl_players WHERE PlayerID = '" & Session("PlayerID") & "'"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF and oRS.BOF) Then
		strThreadAuthorName			= Session("uName")
		strThreadAuthorSignature	= oRS.Fields("PlayerSignature").Value 
		strThreadAuthorTitle		= oRS.Fields("PlayerTitle").Value 
		dtmThreadPostTime			= Now()
	End If
	oRS.NextRecordset
End If

If blnThreadLocked And Not(blnSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Call AddError("Thread is locked, no edits allowed.")
	Response.Clear
	Response.redirect "default.asp"
End If

If blnForumLocked And Not(blnSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Call AddError("Forum is locked, no posts allowed.")
	Response.Clear
	Response.redirect "default.asp"
End If

If Len(strForumPassword) > 0 AND Not(blnSysAdmin Or blnForumModerator) Then
	strSQL = "EXECUTE ForumsCheckPermissions '" & Session("PlayerID") & "', '" & intForumID & "', ''"
	oRS.Open strSQL, oConn
	If oRS.State = 1 Then
		If Not(oRS.EOF) Then
			intReturnCode = oRS.Fields("ReturnCode").Value
			Select Case intReturnCode
				Case 0 
					' Do nothing, success
				Case -1
					Call AddError(oRS.Fields("ErrorMessage").Value)
					oRS.Close 
					Set oRS = Nothing
					oConn.Close
					Set oConn = Nothing
					Response.Clear
					Response.Redirect "enterpassword.asp?threadid=" & intThreadID & "&forumid=" & intForumID
				Case Else
					Call AddError("There was a problem processing your request.")
					oRS.Close
					Set oRS = Nothing
					oConn.Close 
					Set oConn = Nothing
					Response.Clear
					Response.Redirect "default.asp"
			End Select
		End If
		oRS.Close
	End If
End If

Dim blnPreview, strPreviewBody, strPreviewSubject
blnPreview = False

If Request.Form("PreviewForm") = "1" Then
	blnPreview = True
	strPreviewBody = ForumEncode(Request.Form("ThreadBody"))
	strThreadBody = Request.Form("ThreadBody")
	strThreadSubject = Request.Form("ThreadSubject")
	strPreviewSubject = Request.Form("ThreadSubject")
	blnThreadSignature = cBool(Request.Form("ThreadSignature"))
End If

If blnThreadSignature Then 
	strSigChecked = " checked "
Else
	strSigChecked = ""
End If

%> 
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Dim strHeaderColor, strHighlight1, strHighlight2
Dim strBGC
strHeaderColor	= bgcheader
strHighlight1	= bgcone
strHighlight2	= bgctwo
%>
<% Call ContentStart("") %>
<script language="javascript">
<!--
function Preview(oForm) {
	oForm.action = "";
	oForm.PreviewForm.value = "1";
	oForm.submit();
}
//-->
</script>
		<table BORDER="0" cellspacing="0" cellpadding="0" width="100%">
		<tr>
			<td CLASS="pageheader"><%
			Response.Write "<a href=""default.asp"">TWL Forums</A>"
			Response.Write " &raquo; "
			Response.Write "<a href=""default.asp#Category" & intCategoryID & """>" & strCategoryName & "</A>"
			Response.Write " &raquo; "
			Response.Write "<a href=""forumdisplay.asp?forumid=" & intForumID & """>" & strForumName & "</A>"
			Response.Write " &raquo; "
			Response.Write " " & strMode & " Thread "
			%> </td>
		</tr>
		</table>

    <table border=0 cellspacing=0 cellpadding=0 width=100% class="cssBordered">
		<form id="frmThread" name="frmThread" action=forumengine.asp method=post>
		<input type=hidden name=PreviewForm id=PreviewForm value="0">
		<input type=hidden name=ForumID value="<%=intForumID%>">
		<input type=hidden name=SaveType value="Thread">
		<input type=hidden name=Mode value="<%=strMode%>">
		<input type=hidden name=ThreadID value="<%=intThreadID%>">
			<TR>
				<TH COLSPAN=2 CLASS="category" BGCOLOR="<%=bgcblack%>"><%=strMode & " Thread "%></TH>
			</TR>
			<TR bgcolor=<%=strHighlight2%>>
				<TD align="right"><B>Logged in as:</B></td>
				<TD><%=Server.HTMLEncode(Session("uName") & "")%></td>
			</TR>
			<TR bgcolor=<%=strHighlight1%>>
				<TD align="right"><B>Subject:</B></td>
				<% If strMode = "Add" Then %>
					<TD><input type="text" name=ThreadSubject id=ThreadSubject value="<%=Server.HTMLEncode(strThreadSubject & "")%>" maxlength=100 size=60></td></tr>
				<% Else %>
					<TD><%=Server.HTMLEncode(strThreadSubject & "")%></td>
					<input type="hidden" name=ThreadSubject id=ThreadSubject value="<%=Server.HTMLEncode(strThreadSubject & "")%>" 
				<% End If %>
			</TR>
			<TR bgcolor=<%=strHighlight2%>>
				<TD valign=top align="right"><B>Body:</B></td>
				<TD><textarea cols=60 rows=10 name="ThreadBody" id="ThreadBody"><%=Server.htmlencode(strThreadBody & "")%></textarea></td>
			</tr>
			<TR bgcolor=<%=strHighlight1%>>
				<TD colspan=2 align=center><input type="checkbox" class=borderless <%=strSigChecked%> ID="ThreadSignature" name="ThreadSignature" value="1"> Show Signature?
				<%
				If strMode = "Edit" Then Response.Write "<BR><input class=borderless type=checkbox id=""DeleteThread"" name=""DeleteThread"" value=""1""> Delete Thread?"
				%></td>
			</tr>
			<TR bgcolor=<%=strHighlight2%>>
				<TD colspan=2 align=center>
					<input type=submit id=thread name=thread value="<%=strMode%> Thread">
					<input type="button" id="preview" name="preview" value="Preview Thread" onclick="javascript:Preview(this.form);">
				</td>
			</tr>
	</form>	
	</table>
	<% If blnPreview Then %>
	<BR><BR>
    <table border=0 cellspacing=0 cellpadding=0 width=100% class="cssBordered">
		<TR bgcolor=<%=bgcblack%>>
			<TH COLSPAN=2>Preview Thread</TH>
		</TR>
		<TR bgcolor=<%=strHeaderColor%>>
			<TH CLASS="columnheader" ALIGN="LEFT">Author</TH>
			<TH CLASS="columnheader" ALIGN="LEFT">Thread</TH>
		</TR>
		<TR bgcolor="<%=strHighlight1%>"><TD width="15%" valign=top><b><%=server.htmlencode("" & strThreadAuthorName)%></b><BR><span class="usertitle"><%=strThreadAuthorTitle%></span></td>
			<TD valign=top>
			<table border=0 cellspacing=0 cellpadding=2 width=100% height=100%>
			<TR><TD><%
				Call FixDate(dtmThreadPostTime, intTimeZoneDifference, strDate, strTime, strDateMask, bln24HourTime)
				Response.Write "<span class=""smalldate"">" & strDate & "</span> <span class=""smalltime"">" & strTime & "</span>"
				%>
			</td></tr>
			<TR><TD><HR class=forum></TD></TR>
			<TR><TD>						
			<%
			Response.Write replace(strPreviewBody, chr(13), "<BR>")
			if blnThreadEdited then
				Response.Write "<BR><font class=edited>Post edited by " & Server.htmlencode("" & strThreadEditBy) & " at " & dtmThreadEditTime & "</font><BR>"
			end if
			if blnThreadSignature then
				If Not(IsNull(strThreadAuthorSignature)) Then
					Response.Write "<BR>" & Replace(strThreadAuthorSignature, chr(13), "<BR>") & ""
				End If
			end if
			%>
		</td></tr>
		</TABLE>
	</TD></TR></TABLE>
	<% End If %>
	<center><% Call DisplayForumFooter%></center>
<% Call ContentEnd()%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>