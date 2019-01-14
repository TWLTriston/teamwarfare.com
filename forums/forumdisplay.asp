<%
Option Explicit

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle
strPageTitle = "TWL: Forums"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()
Call ForumCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim blnLoggedIn, blnSysAdmin
blnLoggedIn = Session("LoggedIn")
blnSysAdmin = IsSysAdmin()

Dim intCategoryID, strCategoryName, blnForumLocked
Dim intForumID, strForumName, strForumPassword, intCounter, blnForumModerator
Dim intPerPage, intPageNum, intPages, intThreadPerPage, intItemPages
Dim intShowDays, intReturnCode

intForumID = Request.QueryString ("ForumID")

If Len(intForumID) = 0 Or Not(IsNumeric(intForumID)) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Call AddError("Invalid forum id passed.")
	Response.Clear
	Response.Redirect "default.asp"
End If

strsql = "SELECT ForumName, ForumPassword, ForumLocked, CategoryID, CategoryName = (SELECT CategoryName FROM tbl_category WHERE tbl_category.CategoryID = tbl_forums.CategoryID) FROM tbl_forums WHERE ForumID='" & CheckString(intForumID) & "'"
oRS.Open strSQL, oConn
If Not(oRS.EOF AND oRS.BOF) Then
	strForumName = oRS.Fields("ForumName").Value
	strPageTitle = "TWL Forums: " & strForumName
	strForumPassword = oRS.Fields("ForumPassword").Value 
	intCategoryID = oRS.Fields("CategoryID").Value 
	strCategoryName = oRS.Fields("CategoryName").Value 
	blnForumLocked = oRS.Fields("ForumLocked").Value
Else
	oRS.Close 
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Call AddError("Invalid forum id passed.")
	Response.Clear
	Response.Redirect "default.asp"
End If
oRS.NextRecordset 

if intCategoryID = 3 AND NOT(IsSysAdminLevel2()) Then
	Response.Clear
	Response.Redirect "default.asp"
End If

blnForumModerator = IsForumModerator(intForumID)

If Len(strForumPassword) > 0 AND Not(blnSysAdmin Or blnForumModerator) Then
	strSQL = "EXECUTE ForumsCheckPermissions '" & Session("PlayerID") & "', '" & CheckString(intForumID) & "', ''"
	oRS.Open strSQL, oConn
	If oRS.State = 1 Then
		If Not(oRS.EOF) Then
			intReturnCode = CInt(oRS.Fields("ReturnCode").Value)
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
					Response.Redirect "enterpassword.asp?forumid=" & intForumID
				Case Else
					Call AddError("There was a problem processing your request.")
					oRS.Close
					Set oRS = Nothing
					oConn.Close 
					Set oConn = Nothing
					Response.Clear
					Response.Redirect "default.asp"
			End Select
		Else
			Call AddError("There was a problem processing your request.")
			oRS.Close
			Set oRS = Nothing
			oConn.Close 
			Set oConn = Nothing
			Response.Clear
			Response.Redirect "default.asp"
		End If
		oRS.Close
	Else
		Call AddError("There was a problem processing your request.")
		Set oRS = Nothing
		oConn.Close 
		Set oConn = Nothing
		Response.Clear
		Response.Redirect "default.asp"
	End If
End If

intShowDays = Request.Querystring("ShowDays")
intPerPage = Request.Querystring("PerPage")
intPageNum = Request.Querystring("Page")
If Len(intPageNum) = 0 OR Not(IsNumeric(intPageNum)) Then
	intPageNum = 1
Else
	intPageNum = cInt(intpageNum)
End If

If Len(intShowDays) = 0 OR Not(IsNumeric(intShowDays)) Then
	intShowDays = 500
End If

If Len(intPerPage) = 0 OR Not(IsNumeric(intPerPage)) Then
	intPerPage = 35
End If

If Len(intThreadPerPage) = 0 OR Not(IsNumeric(intThreadPerPage)) Then
	intThreadPerPage = 20
End If

Dim dtmDate, intTimeZoneDifference, strDate, strTime
intTimeZoneDifference = Session("intTimeZoneDifference")

Dim strDateMask, bln24HourTime, blnVerticalBars, strColumnColor1, strColumnColor2
strDateMask = "MM-DD-YYYY"
bln24HourTime = False

Dim strCurrentTime, strCurrentDate
Call FixDate(Now(), intTimeZoneDifference, strCurrentDate, strCurrentTime, strDateMask, bln24HourTime)

If blnLoggedIn Then
	Call UpdateForumVisit()
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

blnVerticalBars = False
If blnVerticalBars Then
	strColumnColor1 = ""
	strColumnColor2 = ""
Else
	strColumnColor1 = strHighlight1
	strColumnColor2 = strHighlight2
End If
%>
<% Call ContentStart("") %>
<% Call ForumAds() %>
<table width="100%">
<% Call ShowErrors("<TR><TD>", "</TD></TR>") %>
<tr>
	<td>
		<table width="100%">
		<tr>
			<td CLASS="pageheader"><%
			Response.Write "<a href=""default.asp"">TWL Forums</A>"
			Response.Write " &raquo; "
			Response.Write "<a href=""default.asp#Category" & intCategoryID & """>" & strCategoryName & "</A>"
			Response.Write " &raquo; "
			Response.Write strForumName
			%> </td>
		</tr>
		<TR>
			<td align="right" class="littlelinks"><%
					If blnForumLocked AND Not(bSysAdmin) Then
						Response.Write "Forum Locked"
					Else
						Response.Write "<a href=""newthread.asp?forumid=" & intForumID & """>New Thread</a>"
					End If %>
					</td>
		</tr>
		</table>
	</td>
</tr>
<TR>
	<TD>&nbsp;</TD>
</TR>
<TR>
	<TD>
        <table width="100%" class="cssBordered">
		<%
		oRS.PageSize = intPerPage
		oRS.CacheSize = intPerPage
		oRS.CursorLocation = adUseClient
		'strsql = "EXEC ForumDisplay '" & ForumID & "'" 
		strSQL = " SELECT ThreadLocked, ThreadSticky, ThreadAuthorName, ThreadAuthorID, "
		strSQL = strSQL & "ThreadPostCount, ThreadViewCount, ThreadID, ThreadSubject, ThreadLastPostTime, ThreadLastPosterName FROM tbl_threads WHERE ForumID = '" & CheckString(intForumID) & "'"
		'strSQL = strSQL & " AND DateDiff(D, ThreadLastPostTime, GetDate()) < " & intShowDays & " ORDER BY ThreadSticky DESC, ThreadLastPostTime DESC"
		strSQL = strSQL & " ORDER BY ThreadSticky DESC, ThreadLastPostTime DESC"
		oRS.Open strSQL, oConn, 3, 3
		If Not(oRS.EOF AND oRS.BOF) Then
			intPages		= oRS.PageCount
			If intPageNum <= intPages Then
				oRS.AbsolutePage = intPageNum
			Else
				oRs.AbsolutePage = 1
				intPageNum = 1
			End If
			
			strBGC = strHighlight1
			%>
			<tr bgcolor="<%=strHeaderColor%>">
				<TH CLASS="columnheader">&nbsp;</TH>
				<TH CLASS="columnheader" width="80%">Thread</TH>
				<TH CLASS="columnheader">Author</TH>
				<TH CLASS="columnheader">Replies</TH>
				<TH CLASS="columnheader">Views</TH>
				<TH CLASS="columnheader" nowrap="nowrap">Last Post</TH>
			</TR>
			<%
			intCounter = 0
			Do While Not(oRS.EOF) AND intCounter < intPerPage
				intCounter = intCounter + 1
				Response.Write "<TR bgcolor=" & strBGC & " valign=""middle""><TD width=10 BGCOLOR=""" & strColumnColor1 & """>"
				If oRS.Fields("ThreadLocked").Value OR blnForumLocked Then
						Response.Write "<img src='/images/locked.gif' alt='locked' border=""0"" />"
				Else
					If isnull(session("CookieTime")) then
						Response.Write "<img src='/images/lighton.gif' alt='new posts' border=""0"" />"
					Else
						if cdate(session("CookieTime")) < ors.fields("ThreadLastPostTime").value then
							Response.Write "<img src='/images/lighton.gif' alt='new posts' border=""0"" />"
						else
							Response.Write "<img src='/images/lightoff.gif' alt='no new posts' border=""0"" />"
						end if						
					End If
				End If
				Response.Write "<TD BGCOLOR=""" & strColumnColor2 & """>"
				If oRs.Fields("ThreadSticky").Value Then
					Response.Write "<B>Sticky:</B> "
				End If				
				Response.Write "<a href='showthread.asp?forumid=" & intForumID & "&amp;threadid=" & ors.fields("ThreadID").value & "'>" & server.HTMLEncode(Ors.fields("ThreadSubject")) & "</a>"
				Call WriteThreadPages(intForumID, ors.fields("ThreadID").value, ors.fields("ThreadPostCount").value, intThreadPerPage)
				Response.Write "</td>"
				Response.Write "<TD BGCOLOR=""" & strColumnColor1 & """ nowrap align=center><a href=""/viewplayer.asp?player=" & Server.URLEncode(oRS.Fields("ThreadAuthorName").Value) & """>" & Server.HTMLEncode(ors.fields("ThreadAuthorName").value & "") & "</A></td>"
				Response.Write "<TD BGCOLOR=""" & strColumnColor2 & """ align=center>" & ors.fields("ThreadPostCount").value & "</td>"
				Response.Write "<TD BGCOLOR=""" & strColumnColor1 & """ align=center>" & ors.fields("ThreadViewCount").value & "</td>"
				Call FixDate(ors.fields("ThreadLastPostTime").value, intTimeZoneDifference, strDate, strTime, strDateMask, bln24HourTime)
				Response.Write "<TD BGCOLOR=""" & strColumnColor2 & """ nowrap align=""right""><span class=""smalldate"">" & strDate & "</span>"
				Response.Write "&nbsp;<span class=""smalltime"">" & strTime & "</span><br /><span class=""note"">by <b>" & Server.HTMLEncode (ors.fields("ThreadLastPosterName").value & "") & "</b></span></td>"
'				Response.Write "<td bgcolor=""" & strColumnColor2 & """ align=center>" & ors.fields("ThreadLastPostTime").value & "<br /><font color='#AAAAAA'>by: " & Server.HTMLEncode(ors.fields("ThreadLastPosterName").Value & "") & "</td></tr>"
				if strBGC = strHighlight1 then
					strBGC = strHighlight2
				else
					strBGC = strHighlight1
				end if
				ors.movenext
			loop
		Else
			%>
			<tr bgcolor="<%=bgcblack%>">
				<tH width="10">&nbsp;</tH>
				<tH width="45%">Thread Title</tH>
				<tH width="15%">Author</th>
				<tH width="8%">Replies</th>
				<TH width="7%">Views</th>
				<tH width="25%">Last Post</th>
			</TR>
			<TR>
				<TD COLSPAN=6 BGCOLOR="<%=strHighlight1%>"><I>No posts have been made yet</I></TD>
			</TR>
			<%
		End If
		oRS.Close
		%>
		</table>
	</TD>
</TR>
<TR>
	<TD>&nbsp;</TD>
</TR>
<TR>
	<TD>
		<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH=100%>
		<TR>
			<TD class="littlelinks"><%
					If blnForumLocked Then
						Response.Write "Forum Locked"
					Else
						Response.Write "<a href=""newthread.asp?forumid=" & intForumID & """>New Thread</a>"
					End If %>
				</TD>
		</TR>
		<% Call ListPages(intPageNum, intPages, intForumID) %>
		</TABLE>
	</TD>
</TR>
<TR>
	<TD>&nbsp;</TD>
</TR>
<% Call DisplayThreadLegend() %>
</table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing

Function ListPages(byVal iPageNum, byVal iTotalPages, byVal iForumID)
	Dim i
	If iTotalPages > 1 Then
		Response.Write "<TR>"
		Response.Write "<TD align=""right"" CLASS=""pagelist"">"
		Response.Write "Pages (" & iTotalPages & "): <B>"
		If iPageNum > 5 Then
			Response.Write " <a alt=""First Page"" href=""forumdisplay.asp?forumid=" & iForumID & "&amp;page=1"">&laquo; First</A> ... "
		End If
		If iPageNum > 1 Then
			Response.Write " <a alt=""Previous Page"" href=""forumdisplay.asp?forumid=" & iForumID & "&amp;page=" & iPageNum - 1 & """>&laquo;</A> "
		End If
		For i = iPageNum - 5 To iPageNum + 5
			If i > 0 Then
				If i = iPageNum Then
					Response.Write " <span class=""currentpage"">[" & i & "]</span>"
				ElseIf i <= iTotalPages Then
					Response.Write " <a href=""forumdisplay.asp?forumid=" & iForumID & "&amp;page=" & i & """>" & i & "</a>"
				End If				
			End If
		Next
		If iPageNum < iTotalPages Then
			Response.Write " <a title=""Next Page"" href=""forumdisplay.asp?forumid=" & iForumID & "&amp;page=" & iPageNum + 1 & """>&raquo;</A> "
		End If
		If iPageNum + 5 < iTotalPages Then
			Response.Write " ... <a title=""Last Page"" href=""forumdisplay.asp?forumid=" & iForumID & "&amp;page=" & iTotalpages & """>Last &raquo;</A>"
		End If
		Response.Write "</B>"
		Response.Write "</TD></TR>"
	End If
End Function

Function WriteThreadPages(byVal iForumID, byVal iThreadID, byVal iPosts, byVal iPerPage)
	Dim iThisPages, i
	If (iPosts) > iPerPage Then
		Response.Write "<br />&nbsp;<font class=""pagelist"">("
		iThisPages = (iPosts) / iPerPage
		if (iThisPages) <> fix(iThisPages) then
			iThisPages = fix(iThisPages) + 1
		end if
'		Response.Write thispages
		For i = 1 to iThisPages
			If i <= 10 Then
				Response.Write " <a href=""showthread.asp?forumid=" & iForumID & "&amp;perpage=" & iPerPage & "&amp;threadid=" & iThreadID & "&amp;page=" & i & """>" & i & "</a>"
			End If
		Next
		If iThisPages > 10 Then
			Response.Write " ... <a href=""showthread.asp?forumid=" & iForumID & "&amp;perpage=" & iPerPage & "&amp;threadid=" & iThreadID & "&amp;page=" & iThisPages & """>Last page</a>"
		End If
		Response.Write " )</font>"
	End If
End Function
%>