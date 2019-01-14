<%
Option Explicit

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle, strMode, intPostID

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
	intPostID = Trim(Request.QueryString("postid"))
	If Not(IsNumeric(intPostID)) Or Len(intPostID) = 0 Then
		Call AddError("Invalid post id passed.")
		Response.Clear
		Response.Redirect "/default.asp"
	End If
End If

strPageTitle = "TWL: " & strMode & " Reply"

Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Dim oConn, oRS, strSQL
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open Application("ConnectStr")

Set oRS = Server.CreateObject ("ADODB.RecordSet")

Dim blnShowSigs
blnShowSigs = Session("ShowSigs")
If Len(blnShowSigs) = 0 THen	
	blnShowSigs = True
End If

Dim intForumID, intThreadID, strForumName, strForumPassword
Dim intThreadAuthorID, strThreadSubject, strThreadBody, blnThreadSignature, strThreadAuthorName, strThreadAuthorTitle
Dim intPostAuthorID, strPostBody, blnPostSignature, strSigChecked
Dim blnLoggedIn, blnSysAdmin
Dim blnForumAccess, blnForumModerator
Dim dtmThreadPostTime, blnThreadEdited, dtmThreadEditTime
Dim strThreadEditBy, strThreadAuthorSignature
Dim intCategoryID, strCategoryName
Dim blnForumLocked, blnThreadLocked, intReturnCode

if uCase(Request.Cookies("PerPage")("ShowSig")) = "Y" Then
	blnPostSignature  = True
End if

intThreadID = Request.QueryString ("threadid")
If Len(intThreadID) = 0 OR Not(IsNumeric(intThreadID)) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Call AddError("Invalid thread id. Please check linking url.")
	Response.Clear
	Response.Redirect "default.asp"
End If

Call CheckCookie()
Dim bSysAdmin, bLoggedIn, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bLoggedIn = Session("LoggedIn")
blnLoggedIn = Session("LoggedIn")
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

'' Get thread infomration, so we can verify their access
strSQL = "SELECT tbl_threads.*, PlayerTitle, PlayerSignature FROM tbl_threads INNER JOIN tbl_players ON ThreadAuthorID = PlayerID WHERE ThreadID = '" & intThreadID & "'"
oRS.Open strSQL, oConn
If Not(oRS.EOF AND oRS.BOF) Then
	intForumID					= oRS.Fields("ForumID").Value
	strThreadSubject			= oRS.Fields("ThreadSubject").Value
	dtmThreadPostTime			= oRS.Fields("ThreadPostTime").Value
	strThreadBody				= oRS.Fields("ThreadBody").Value
	strThreadAuthorname			= oRS.Fields("ThreadAuthorName").Value
	strThreadAuthorTitle		= oRS.Fields("PlayerTitle").Value
	strThreadAuthorSignature	= oRS.Fields("PlayerSignature").Value 
	blnThreadEdited				= oRS.Fields("ThreadEdited").Value
	dtmThreadEditTime			= oRS.Fields("ThreadEditTime").Value
	strThreadEditBy				= oRS.Fields("ThreadEditBy").Value
	blnThreadSignature			= oRS.Fields("ThreadSignature").Value
	blnThreadLocked				= oRS.Fields("ThreadLocked").Value 
Else
	oRS.Close
	Set oRs = Nothing
	oConn.Close
	Set oConn = Nothing
	Call AddError("Invalid thread id. Please check linking url.")
	Response.Clear
	Response.Redirect "default.asp"
End If
oRS.NextRecordSet

strsql = "SELECT ForumName, ForumLocked, ForumPassword, CategoryID, CategoryName = (SELECT CategoryName FROM tbl_category WHERE tbl_category.CategoryID = tbl_forums.CategoryID) FROM tbl_forums WHERE ForumID='" & intForumID & "'"
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

If blnThreadLocked And Not(blnSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Call AddError("Thread is locked, no posts allowed.")
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

blnForumModerator = IsForumModerator(intForumID)
blnSysAdmin = IsSysAdmin()

Dim strPostAuthorName, strPostAuthorSignature, strPostAuthorTitle
Dim dtmPostTime, blnPostEdited, strPostEditBy, dtmPostEditTime
If strMode = "Edit" Then
	strSQL = "SELECT tbl_posts.*, PlayerHandle, PlayerTitle, PlayerSignature "
	strSQL = strSQL & " FROM tbl_posts INNER JOIN tbl_players ON PostAuthorID = PlayerID "
	strSQL = strSQL & " WHERE PostID='" & intPostID & "'"
	oRS.Open strSQL, oConn 
	If Not(oRS.EOF and oRS.BOF) Then
		intThreadID				= oRS.Fields("ThreadID").Value
		strPostBody				= oRS.Fields("PostBody").Value
		intPostAuthorID			= oRS.Fields("PostAuthorID").Value
		blnPostSignature		= cBool(oRS.Fields("PostSignature").Value)
		strPostAuthorName		= oRS.Fields("PlayerHandle").Value 
		strPostAuthorTitle		= oRS.Fields("PlayerTitle").Value 
		dtmPostTime				= oRS.Fields("PostTime").Value 
		strPostAuthorSignature	= oRS.Fields("PlayerSignature").Value 
	End If
	oRS.Close 

	blnPostEdited = True
	strPostEditBy = Session("uName")
	dtmPostEditTime = Now()

	If Not(blnSysAdmin OR blnForumModerator) AND cStr(intPostAuthorID & "") <> cStr(Session("PlayerID") & "") Then
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Call AddError("You do not have access to edit other people's posts.")
		Response.Clear
		Response.redirect "/errorpage.asp?error=3"
	End If
Else
	strSQL = "SELECT PlayerTitle, PlayerSignature FROM tbl_players WHERE PlayerID = '" & Session("PlayerID") & "'"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF and oRS.BOF) Then
		strPostAuthorName		= Session("uName")
		strPostAuthorSignature	= oRS.Fields("PlayerSignature").Value 
		strPostAuthorTitle		= oRS.Fields("PlayerTitle").Value 
		dtmPostTime				= Now()
	End If
	oRS.NextRecordset
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

If Len(Request.QueryString("QuoteThread")) > 0 AND IsNumeric(Request.QueryString("QuoteThread")) Then
	strSQL = "SELECT ThreadAuthorName, ThreadBody FROM tbl_threads WHERE ThreadID='" & Request.QueryString("QuoteThread") & "'"
	oRS.Open strSQL, oConn 
	If Not(oRS.EOF and oRS.BOF) Then
		strPostBody = "[quote=""" & ors.fields("ThreadAuthorName").value & """]"
		strPostBody = strPostBody & ors.fields("ThreadBody").value & "[/quote]"
	End If
	oRS.NextRecordset 
ElseIf Len(Request.QueryString("QuotePost")) > 0 AND IsNumeric(Request.QueryString("QuotePost")) Then
	strSQL = "SELECT PlayerHandle, PostBody FROM tbl_posts INNER JOIN tbl_players ON PostAuthorID = PlayerID WHERE PostID='" & Request.QueryString("QuotePost") & "'"
	oRS.Open strSQL, oConn 
	If Not(oRS.EOF and oRS.BOF) Then
		strPostBody = "[quote=""" & ors.fields("PlayerHandle").value & """]"
		strPostBody = strPostBody & ors.fields("PostBody").value & "[/quote]"
	End If
	oRS.NextRecordset 
End If

Dim intTimeZoneDifference, strDate, strTime
Dim strCurrentTime, strCurrentDate
Dim strDateMask, bln24HourTime

intTimeZoneDifference = 0

strDateMask = "MM-DD-YYYY"
bln24HourTime = False
Dim blnPreview, strPreviewBody, strPreviewSubject
blnPreview = False

If Request.Form("PreviewForm") = "1" Then
	blnPreview = True
	strPreviewBody = ForumEncode(Request.Form("PostBody"))
	strPostBody = Request.Form("PostBody")
	blnPostSignature = cBool(Request.Form("PostSignature"))
End If

If blnPostSignature Then 
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
			If Len(strThreadSubject) > 40 Then
				Response.Write "<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			End If
			Response.Write "<a href=""showthread.asp?forumid=" & intForumID & "&threadid=" & intThreadID & """>" & strThreadSubject & "</A>"
			Response.Write " &raquo; "
			Response.Write " " & strMode & " Reply "
			%> </td>
		</tr>
		</table>

		<form id="frmThread" name="frmThread" action=forumengine.asp method=post>
		<input type=hidden name=PreviewForm id=PreviewForm value="0">
		<input type=hidden name=ForumID value="<%=intForuMID%>">
		<input type=hidden name=SaveType value="Post">
		<input type=hidden name=Mode value="<%=strMode%>">
		<input type=hidden name=ThreadID value="<%=intThreadID%>">
		<input type=hidden name=PostID value="<%=intPostID%>">
     <table border=0 cellspacing=0 cellpadding=0 width=100% class="cssBordered">
			<TR>
				<TH COLSPAN=2 CLASS="category" BGCOLOR="<%=bgcblack%>"><%=strMode & " Reply "%></TH>
			</TR>
			<TR bgcolor=<%=strHighlight2%>>
				<TD align="right"><B>Logged in as:</B></td>
				<TD><%=Server.HTMLEncode(Session("uName") & "")%></td>
			</TR>
			<TR bgcolor=<%=strHighlight2%>>
				<TD valign=top align="right"><B>Body:</B></td>
				<TD><textarea cols=60 rows=10 name="PostBody" id="PostBody"><%=Server.htmlencode(strPostBody & "")%></textarea></td>
			</tr>
			<TR bgcolor=<%=strHighlight1%>>
				<TD colspan=2 align=center><input type="checkbox" class=borderless <%=strSigChecked%> ID="PostSignature" name="PostSignature" value="1"> Show Signature?
				<%
				If strMode = "Edit" Then Response.Write "<BR><input class=borderless type=checkbox id=""DeletePost"" name=""DeletePost"" value=""1""> Delete Post?"
				%></td>
			</tr>
			<TR bgcolor=<%=strHighlight2%>>
				<TD colspan=2 align=center>
					<input type=submit value="<%=strMode%> Post" id=submit1 name=submit1>
					<input type="button" id="preview" name="preview" value="Preview Post" onclick="javascript:Preview(this.form);">
				</td>
			</tr>
		</table>
		</form>			

		<% If blnPreview Then %>
		<br /><br />
		<table border=0 cellspacing=0 cellpadding=0 width=100% class="cssBordered">
			<TR bgcolor=<%=bgcblack%>>
				<TH COLSPAN=2>Preview Post</TH>
			</TR>
			<TR bgcolor=<%=strHeaderColor%>>
				<TH CLASS="columnheader" ALIGN="LEFT">Author</TH>
				<TH CLASS="columnheader" ALIGN="LEFT">Thread</TH>
			</TR>
			<TR bgcolor="<%=strHighlight1%>"><TD width="15%" valign=top><b><%=server.htmlencode("" & strPostAuthorName)%></b><BR><span class="usertitle"><%=strPostAuthorTitle%></span></td>
				<TD valign=top>
				<table border=0 cellspacing=0 cellpadding=2 width=100% height=100%>
				<TR><TD><%
					Call FixDate(dtmPostTime, intTimeZoneDifference, strDate, strTime, strDateMask, bln24HourTime)
					Response.Write "<span class=""smalldate"">" & strDate & "</span> <span class=""smalltime"">" & strTime & "</span>"
					%>
				</td></tr>
				<TR><TD><HR class=forum></TD></TR>
				<TR><TD>						
				<%
				Response.Write ForuMEncode2(strPreviewBody)
				if blnPostEdited AND strPostAuthorName <> "Triston" AND strPostAuthorName <> "Polaris" AND strPostAuthorName <> "ZedsDead" AND strPostAuthorName <> "TotalCarnage" then
					Response.Write "<BR><font class=edited>Post edited by " & Server.htmlencode("" & strPostEditBy) & " at " & dtmPostEditTime & "</font><BR>"
				end if
				if blnPostSignature then
					If Not(IsNull(strPostAuthorSignature)) Then
						Response.Write "<BR>" & ForumEncode2(strPostAuthorSignature) & ""
					End If
				end if
				%>
				</TD></TR>
			</TABLE>
		</TD></TR></TABLE>
		<% End If %>
		
		<br /><br />
		<table border=0 cellspacing=0 cellpadding=0 width=100% class="cssBordered">
				<TH COLSPAN=2 BGCOLOR="<%=bgcblack%>">First Post in Thread</TH>
				<TR bgcolor="<%=strHighlight1%>"><TD width="15%" valign=top><b><%=server.htmlencode(strThreadAuthorName)%></b><BR><span class="usertitle"><%=strThreadAuthorTitle%></span></td>
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
						Response.Write ForumEncode2(strThreadBody)
						If blnThreadEdited  AND strThreadAuthorName <> "Triston" AND strThreadAuthorName <> "Polaris" AND strThreadAuthorName <> "ZedsDead" AND strThreadAuthorName <> "TotalCarnage" Then
							Response.Write "<BR><font class=edited>Post edited by " & Server.htmlencode(strThreadEditBy) & " at " & dtmThreadEditTime & "</font><BR>"
						End If
						If blnThreadSignature and blnShowSigs Then
							If Not(IsNull(strThreadAuthorSignature)) Then
								Response.Write "<BR>" & ForumEncode2(strThreadAuthorSignature) & ""
							End If
						End If
						%>
						</TD></TR>
					</Table>
					</TD>
				</TR>
					
		</tABLE>
		
		<br /><br />
		
		<table border=0 cellspacing=0 cellpadding=0 width=100% class="cssBordered">
				<TH COLSPAN=2 BGCOLOR="<%=bgcblack%>">Last 5 Replies</TH>
				<%
				strSQL = "SELECT TOP 5 *, PlayerTitle, PlayerHandle, PlayerSignature "
				strSQL = strSQL & " FROM tbl_posts INNER JOIN tbl_players ON PlayerID = PostAuthorID "
				strSQL = strSQL & " WHERE ThreadID = " & intThreadID & " ORDER BY PostID DESC"
				oRS.Open strSQL, oConn
				If Not(oRS.EOF and oRS.BOF) Then
					strBGC = strHighlight1
					Do While Not(oRS.EOF)
						If Not(IsNull(ors.fields("PostBody").value)) Then
							strPostBody = ForumEncode2(ors.fields("PostBody").value)
						Else
							strPostBody = ""
						End If
						%>						
						<TR bgcolor="<%=strBGC%>"><TD width="15%" valign=top><b><%=server.htmlencode(ors.fields("PlayerHandle").value)%></b><BR><span class="usertitle"><%=ors.fields("PlayerTitle").value%></span></td>
							<TD valign=top>
							<table border=0 cellspacing=0 cellpadding=2 width=100% height=100%>
								<TR><TD><%
									Call FixDate(ors.fields("PostTime").value, intTimeZoneDifference, strDate, strTime, strDateMask, bln24HourTime)
									Response.Write "<span class=""smalldate"">" & strDate & "</span> <span class=""smalltime"">" & strTime & "</span>"
									%>
								</td></tr>
								<TR><TD><HR class=forum></TD></TR>
								<TR><TD>						
								<%
								Response.Write strPostBody
								If oRS.Fields("PostEdited").Value AND ors.fields("PostEditBy").value <> "Triston" AND ors.fields("PostEditBy").value <> "Polaris"  AND ors.fields("PostEditBy").value <> "ZedsDead"  AND ors.fields("PostEditBy").value <> "TotalCarnage" Then
									Response.Write "<BR><font class=edited>Post edited by " & Server.htmlencode(ors.fields("PostEditBy").value) & " at " & ors.fields("PostEditTime").value & "</font><BR>"
								End If
								If oRS.Fields("PostSignature").Value and blnShowSigs Then
									Response.Write "<BR>" & ForumEncode2(ors.fields("PlayerSignature").value) & ""
								End If
								%>
								</TD></TR>
							</Table>
							</TD>
						</TR>
						<%
						If strBGC = strHighlight1 Then
							strBGC = strHighlight2
						Else
							strBGC = strHighlight1
						End If
						oRS.MoveNext
					Loop
				Else
					%>
					<TR BGCOLOR="<%=strHighlight2%>">
						<TD>No replies have been posted.</TD>
					</TR>
					<%
				End If
				oRS.Close 
				%>
			</tABLE>

		<% Call DisplayForumFooter%>

<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>