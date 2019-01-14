<%
Option Explicit

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle
strPageTitle = "TWL: View Thread"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open Application("ConnectStr")

Set oRS = Server.CreateObject ("ADODB.RecordSet")

Dim strHeaderColor, strHighlight1, strHighlight2
Dim strTableBorder, strTableBG, strBGC
strHeaderColor	= Application("HeaderColor")
strHighlight1	= bgcone
strHighlight2	= bgctwo
strTableBorder	= "#444444"
strTableBG		= "#000000"

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin, bForumAdmin
Dim blnShowSigs
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
blnShowSigs = Session("ShowSigs")
If Len(blnShowSigs) = 0 THen	
	blnShowSigs = True
End If
Dim blnLoggedIn, blnSysAdmin, blnForumModerator
blnLoggedIn = Session("Loggedin")
blnSysAdmin = IsSysAdmin()

Dim intThreadID, intForumID, strThreadSubject
Dim strThreadBody, dtmThreadPostTime, strThreadAuthorTitle
Dim blnThreadEdited, dtmEditTime, strThreadEditBy, blnThreadSignature
Dim intPerPage, intPageNum, intPages
Dim blnThreadLocked, blnThreadSticky
Dim strPostBody
DIm strThreadAuthorName, intThreadAuthorID, strThreadAuthorSignature
DIm dtmThreadEditTime
Dim strForumName, strForumPassword
Dim intCounter
Dim intCategoryID, strCategoryName, blnForumLocked
Dim intReturnCode

intThreadID = Request.QueryString("ThreadID")
If Len(intThreadID) = 0 Or Not(IsNumeric(intThreadID)) Then
	Call AddError("Invalid thread id passed.")
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "default.asp"
Else
	intThreadID = clng(intThreadID)
End If
Dim intContributor
strSQL = "SELECT tbl_threads.*, PlayerTitle, Contributor, PlayerSignature FROM tbl_Threads INNER JOIN tbl_players on ThreadAuthorID=PlayerID WHERE ThreadID = '" & intThreadID & "'"
oRS.Open strSQL, oConn
If Not(oRS.EOF and oRS.BOF) Then
	intContributor = oRs.Fields("Contributor").Value
	intForumID = ors.fields("ForumID").value
	blnForumModerator = IsForumModerator(intForumID)
	strThreadSubject = ors.fields("ThreadSubject").value
	strPageTitle = "TWL: " & strThreadSubject
	dtmThreadPostTime = ors.fields("ThreadPostTime").value
	strThreadBody = ors.fields("ThreadBody").value
	intThreadAuthorID = oRS.Fields("ThreadAuthorID").Value 
	strThreadAuthorName = ors.fields("ThreadAuthorName").value
	strThreadAuthorTitle = ors.fields("PlayerTitle").value
	blnThreadEdited = ors.fields("ThreadEdited").value
	dtmThreadEditTime = ors.fields("ThreadEditTime").value
	strThreadEditBy = ors.fields("ThreadEditBy").value
	If blnThreadEdited AND (strThreadEditBy = "Triston" OR strThreadEditBy = "Polaris") Then
		blnThreadEdited = false
	End If
	blnThreadSignature = ors.fields("ThreadSignature").value
	blnThreadLocked = oRS.Fields("ThreadLocked").Value 
	blnThreadSticky = ors.Fields("ThreadSticky").Value 
	strThreadAuthorSignature = oRS.Fields("PlayerSignature").Value 
Else
	Call AddError("Invalid thread id passed.")
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "default.asp"
End If
ors.NextRecordset 

strsql = "SELECT ForumName, ForumPassword, ForumLocked, CategoryID, CategoryName = (SELECT CategoryName FROM tbl_category WHERE tbl_category.CategoryID = tbl_forums.CategoryID) FROM tbl_forums WHERE ForumID='" & intForumID & "'"
oRS.Open strSQL, oConn
If Not(oRS.EOF AND oRS.BOF) Then
	strForumName = oRS.Fields("ForumName").Value
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

blnForumModerator = IsForumModerator(intForumID)

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

strSQL = "EXECUTE ForumsThreadViewed "
strSQL = strSQL & "'" & intThreadID & "'"
oConn.Execute(strSQL)

intPageNum = Request.querystring("page")
If Len(intPageNum) = 0 Or  Not(IsNumeric(intPageNum)) then
	intPageNum = 1
Else
	intPageNum = cInt(intPageNum)
End If

intPerPage = Request.querystring("PerPage")
If Len(intPerPage) = 0 Or Not(IsNumeric(intPerPage)) then
	intPerPage = 20
Else
	intPerPage = cInt(intPerPage)
End If

Dim intTimeZoneDifference, strDate, strTime
Dim strCurrentTime, strCurrentDate
Dim strDateMask, bln24HourTime

intTimeZoneDifference = Session("intTimeZoneDifference")
strDateMask = "MM-DD-YYYY"
bln24HourTime = False

Call FixDate(Now(), intTimeZoneDifference, strCurrentDate, strCurrentTime, strDateMask, bln24HourTime)
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<% Call ContentStart("")%>
<%
oRS.PageSize = intPerPage
oRS.CacheSize = intPerPage
oRS.CursorLocation = adUseClient

strsql = "SELECT tbl_posts.*, Contributor, PlayerID, PlayerTitle, playerHandle, PlayerSignature FROM tbl_posts INNER JOIN tbl_players ON PlayerID = PostAuthorID WHERE ThreadID = " & intThreadID & " ORDER BY PostID ASC"
oRS.Open strSQL, oConn, 3, 3
If Not(oRS.EOF and oRS.BOF) Then
	intPages		= oRS.PageCount
	If Request("lastpage") = "1" THen
		oRs.AbsolutePage = intPages
		intPageNum = intPages
	ElseIf intPageNum <= intPages Then
		oRS.AbsolutePage = intPageNum
	Else
		oRs.AbsolutePage = 1
		intPageNum = 1
	End If
Else
	intPages = 1
End If
intCounter = 0

%>
<table border=0 cellspacing=0 cellpadding=0 width=760 ALIGN=CENTER>
<% Call ShowErrors("<TR><TD>", "</TD></TR>") %>
<tr>
	<td>
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
			Response.Write Server.HTMLEncode(strThreadSubject & "")
			%> </td>
		</tr>
		</table>
	</td>
</tr>
<% Call ListPages(intPageNum, intPages, intThreadID, intForumID) %>
<tr>
  <td>
    <table border=0 cellspacing=0 cellpadding=0 width=100% bgcolor="<%=strTableBorder%>">
    <TR><TD>
		<table border=0 cellspacing=1 width=100% cellpadding=4>
		<TR bgcolor=<%=strHeaderColor%>>
			<TH CLASS="columnheader" ALIGN="LEFT">Author</TH>
			<TH CLASS="columnheader">
				<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH=100%>
				<TR>
					<TH CLASS="columnheader" ALIGN="LEFT">Thread</TH>
					<TH CLASS="columnheader" ALIGN=RIGHT><%
						If blnForumLocked AND Not(bSysAdmin) Then
							Response.Write "Forum Locked"
						ElseIf blnThreadLocked AND Not(bSysAdmin) Then
							Response.Write "<a class=""columnheader"" href=""newthread.asp?forumid=" & intForumID & """>New Thread</a> / "
							Response.Write "Thread Locked"
						Else
							Response.Write "<a class=""columnheader"" href=""newthread.asp?forumid=" & intForumID & """>New Thread</a> / "
							Response.Write "<a class=""columnheader"" href=""newpost.asp?threadid=" & intThreadID & "&forumid=" & intForumID & """>Reply to Thread</a>"
						End If
						%></TH>
				</TR>
				</TABLE>
			</TH>
		</tr>
			<%
				intCounter = 0
				strBGC = strHighlight1
				if intPageNum = 1 then
					%>
					<TR bgcolor="<%=strHighlight1%>"><TD width="15%" valign=top><b><%=server.htmlencode("" & strThreadAuthorName)%></b><BR><span class="usertitle"><%=strThreadAuthorTitle%>
					<%
					If intContributor = 1 Then
						Response.Write "<br /><a href=""/contributors.asp"">TWL Contributor</a>"
					End If
					%>
					</span></td>
					<TD valign=top>
					<table border=0 cellspacing=0 cellpadding=2 width=100% height=100%>
					<TR><TD><%
						Call FixDate(dtmThreadPostTime, intTimeZoneDifference, strDate, strTime, strDateMask, bln24HourTime)
						Response.Write "<span class=""smalldate"">" & strDate & "</span> <span class=""smalltime"">" & strTime & "</span>"
						%> <span class="postoptions"> / <a href="/viewplayer.asp?player=<%=server.URLEncode(strThreadAuthorName)%>">profile</a> <% If blnLoggedIn Then %> / <a href="newpost.asp?threadid=<%=intThreadID%>&quotethread=<%=intThreadID%>&forumid=<%=intForumID%>">quote</a> <% End If %>
						<%
						If blnSysAdmin or (cStr(intThreadAuthorID & "") = cStr(Session("PlayerID") & "")) Or blnForumModerator Then
							Response.Write " / <a href=""newthread.asp?mode=edit&threadid=" & intThreadID & "&forumid=" & intForumID & """>edit</a>"
						End If
						If blnSysAdmin or blnForumModerator Then
							Response.Write " / <a href=""forumengine.asp?SaveType=Thread&mode=Delete&threadid=" & intThreadID & "&forumid=" & intForumID & """>delete</a>"
						End If
					%></span>
					</td></tr>
					<TR><TD><HR class=forum></TD></TR>
					<TR><TD>						
					<%
					Response.Write replace(strThreadBody, chr(13), "<BR>")
					if blnThreadEdited then
						Response.Write "<BR><font class=edited>Post edited by " & Server.htmlencode("" & strThreadEditBy) & " at " & dtmThreadEditTime & "</font><BR>"
					end if
					if blnThreadSignature AND blnShowSigs then
						If Not(IsNull(strThreadAuthorSignature)) Then
							Response.Write "<BR>" & Replace(strThreadAuthorSignature, chr(13), "<BR>") & ""
						End If
					end if
					%>
					</TD></TR>
					</Table>
					</td></tr>
					<%
					strBGC = strHighlight2
				end if
				if not (ors.eof and ors.bof) then
					do while not(ors.eof) AND intCounter < intPerPage
						intCounter = intCounter + 1
						intContributor = oRs.Fields("Contributor").Value
						If Not(IsNull(ors.fields("PostBody").value)) then
							strPostBody = replace(ors.fields("PostBody").value, chr(13), "<BR>")
						else
							strPostBody = ""
						end if
						%>						
						<TR bgcolor="<%=strBGC%>"><TD width="15%" valign=top><b><%=server.htmlencode(ors.fields("playerHandle").value)%></b><BR><span class="usertitle"><%=ors.fields("PlayerTitle").value%>
						<%
						If intContributor = 1 Then
							Response.Write "<br /><a href=""/contributors.asp"">TWL Contributor</a>"
						End If
						%>
						</span></td>
						<TD valign=top>
						<table border=0 cellspacing=0 cellpadding=2 width=100% height=100%>
							<TR><TD><%
								Call FixDate(ors.fields("PostTime").value, intTimeZoneDifference, strDate, strTime, strDateMask, bln24HourTime)
								Response.Write "<span class=""smalldate"">" & strDate & "</span> <span class=""smalltime"">" & strTime & "</span>"
								%> <span class="postoptions"> / <a href="/viewplayer.asp?player=<%=server.URLEncode(oRS.Fields("PlayerHandle").Value)%>">profile</a> <% 
									If blnLoggedIn Then %> / <a href="newpost.asp?threadid=<%=intThreadID%>&quotepost=<%=oRS.Fields("PostID").Value%>&forumid=<%=intForumID%>">quote</a> <% End If %>
							<%
								If blnSysAdmin or (cStr(oRS.Fields("PlayerID").Value & "") = cStr(Session("PlayerID") & "")) Or blnForumModerator Then
									Response.Write " / <a href=""newpost.asp?mode=edit&threadid=" & intThreadID & "&postid=" & oRS.Fields("PostID").Value & "&forumid=" & intForumID & """>edit</a>"
								End If
								If blnSysAdmin or blnForumModerator Then
									Response.Write " / <a href=""forumengine.asp?SaveType=Post&mode=Delete&threadid=" & intThreadID & "&postid=" & oRS.Fields("PostID").Value & "&forumid=" & intForumID & """>delete</a>"
								End If
							%></span>
							</td></tr>
							<TR><TD><HR class=forum></TD></TR>
							<TR><TD>						
							<%
							Response.Write strPostBody
							If oRS.Fields("PostEdited").Value AND (ors.fields("PostEditBy").value <> "Polaris" AND ors.fields("PostEditBy").value <> "Triston") Then
								Response.Write "<BR><font class=edited>Post edited by " & Server.htmlencode(ors.fields("PostEditBy").value) & " at " & ors.fields("PostEditTime").value & "</font><BR>"
							End If
							If oRS.Fields("PostSignature").Value AND blnShowSigs Then
								Response.Write "<BR>" & replace(ors.fields("PlayerSignature").value, chr(13), "<BR>") & ""
							End If
							%>
							</TD></TR>
						</Table>
						</td></tr>
						<%			
						if strBGC = strHighlight1 then
							strBGC = strHighlight2
						else
							strBGC = strHighlight1
						end if
						ors.movenext
					loop
				end if
				ors.close
			%>
			<TR bgcolor=<%=strHeaderColor%>>
				<TH CLASS="columnheader" COLSPAN=2 ALIGN="RIGHT"><%
						If blnForumLocked AND Not(bSysAdmin) Then
							Response.Write "Forum Locked"
						ElseIf blnThreadLocked AND Not(bSysAdmin) Then
							Response.Write "<a class=""columnheader"" href=""newthread.asp?forumid=" & intForumID & """>New Thread</a> / "
							Response.Write "Thread Locked"
						Else
							Response.Write "<a class=""columnheader"" href=""newthread.asp?forumid=" & intForumID & """>New Thread</a> / "
							Response.Write "<a class=""columnheader"" href=""newpost.asp?threadid=" & intThreadID & "&forumid=" & intForumID & """>Reply to Thread</a>"
						End If
				%></TH>
			</tr>
			</table>
		</td>
	</tr>
	</table>
	</td>
</tr>
<% call ListPages(intPageNum, intPages, intThreadID, intForumID) %>
<tr>
	<td>
		<table BORDER="0" cellspacing="0" cellpadding="0" width="100%">
		<tr>
			<td CLASS="pageheader"> <%
			Response.Write "<a href=""default.asp"">TWL Forums</A>"
			Response.Write " &raquo; "
			Response.Write "<a href=""default.asp#Category" & intCategoryID & """>" & strCategoryName & "</A>"
			Response.Write " &raquo; "
			Response.Write "<a href=""forumdisplay.asp?forumid=" & intForumID & """>" & strForumName & "</A>"
			Response.Write " &raquo; "
			Response.Write Server.HTMLEncode(strThreadSubject & "")
			%> </td>
		</tr>
		</table>
	</td>
</tr>
<%Call DisplayForumFooter()%>
<%
If blnSysAdmin Or blnForumModerator Then
	Response.Write "<TR><TD COLSPAN=2 ALIGN=CENTER CLASS=""littlelinks"">"
	If blnThreadSticky Then
		Response.Write "<a href=""forumengine.asp?savetype=UnStickyThread&threadid=" & intThreadID & "&forumid=" & intForumID & """>Unsticky Thread</A> / "
	Else
		Response.Write "<a href=""forumengine.asp?savetype=StickyThread&threadid=" & intThreadID & "&forumid=" & intForumID & """>Sticky Thread</A> / "
	End If
	If blnThreadLocked Then
		Response.Write "<a href=""forumengine.asp?savetype=UnlockThread&threadid=" & intThreadID & "&forumid=" & intForumID & """>Unlock Thread</A>"
	Else
		Response.Write "<a href=""forumengine.asp?savetype=LockThread&threadid=" & intThreadID & "&forumid=" & intForumID & """>Lock Thread</A>"
	End If
	Response.Write "</TD></TR>"
	If blnSysAdmin Then
		Response.Write "<tr><td colspan=""2"" align=""center"" class=""littlelinks"">"
		strSQL = "SELECT ForumName, ForumID FROM tbl_Forums ORDER BY ForumName "
		oRS.Open strSQL, oConn
		If Not(orS.EOF AND oRS.BOF) Then
			%>
			<form name="frmMoveThread" id="frmMoveThread" action="forumengine.asp" method="post">
			<input type="hidden" name="SaveType" id="SaveType" value="MoveThread" />
			<input type="hidden" name="ThreadID" id="SaveType" value="<%=intThreadID%>" />
			<select name="selForumID" id="selForumID">
			<%
			Do While Not(oRS.EOF)
				Response.Write "<option value=""" & oRS.Fields("ForumID").Value & """"
				If oRS.Fields("ForumID").Value = intForumID Then
					Response.Write " selected=""selected"""
				End If
				Response.Write ">" & Server.HTMLEncode(oRS.Fields("FOrumName").Value & "") & "</option>" & vbCrLf
				oRS.MoveNext
			Loop
			%>&nbsp;<input type="submit" value="Move It" />
			</form>
			<%
		End If
		oRS.Close
		Response.Write "</td></tr>"
	End If
End If
%>	
</table>
<% Call ContentEnd()%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = NothingFunction ListPages(byVal iPageNum, byVal iTotalPages, byVal iThreadID, byVal iForumID)
	Dim i
	If iTotalPages > 1 Then
		Response.Write "<TR><TD><IMG SRC=""/images/spacer.gif"" HEIGHT=5></TD></TR>"
		Response.Write "<TR>"
		Response.Write "<TD CLASS=""pagelist"">"
		Response.Write "Pages (" & iTotalPages & "): <B>"
		If iPageNum > 5 Then
			Response.Write " <a alt=""First Page"" href=""showthread.asp?forumid=" & iForumID & "&threadid=" & iThreadID & "&page=1"">&laquo; First</A> ... "
		End If
		If iPageNum > 1 Then
			Response.Write " <a alt=""Previous Page"" href=""showthread.asp?forumid=" & iForumID & "&threadid=" & iThreadID & "&page=" & iPageNum - 1 & """>&laquo;</A> "
		End If
		For i = iPageNum - 5 To iPageNum + 5 
			If i > 0 Then
				If i = iPageNum Then
					Response.Write " <span class=""currentpage"">[" & i & "]</span>"
				ElseIf i <= iTotalPages Then
					Response.Write " <a href=""showthread.asp?forumid=" & iForumID & "&threadid=" & iThreadID & "&page=" & i & """>" & i & "</a>"
				End If				
			End If
		Next
		If iPageNum < iTotalPages Then
			Response.Write " <a alt=""Next Page"" href=""showthread.asp?forumid=" & iForumID & "&threadid=" & iThreadID & "&page=" & iPageNum + 1 & """>&raquo;</A> "
		End If
		If iPageNum + 5 < iTotalPages Then
			Response.Write " ... <a alt=""Last Page"" href=""showthread.asp?forumid=" & iForumID & "&threadid=" & iThreadID & "&page=" & iTotalpages & """>Last &raquo;</A>"
		End If
		Response.Write "</B>"
		Response.Write "</TD></TR>"
		Response.Write "<TR><TD><IMG SRC=""/images/spacer.gif"" HEIGHT=5></TD></TR>"
	End If
End Function
%>
