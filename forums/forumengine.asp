<%
Option Explicit

Server.ScriptTimeout = 5
'' This is an engine page. It does not display to the user.
Dim strPageTitle

Dim oConn, oRS, strSQL
Set oConn = Server.CreateObject("ADODB.Connection")
Dim oCmd, intReturnValue

oConn.Open Application("ConnectStr")

Set oRS = Server.CreateObject ("ADODB.RecordSet")

Call CheckCookie()

Dim blnLoggedIn, blnSysAdmin, blnForumModerator
blnLoggedIn = Session("LoggedIn")
blnSysAdmin = IsSysAdmin()

%>
<!-- #include file="../include/i_funclib.asp" -->
<!-- #include file="../include/adovbs.inc" -->
<%
Dim strSaveType
strSaveType = Request.Form("SaveType")
If Len(strSaveType) = 0 Then
	strSaveType = Request.QueryString("SaveType")
End If

If Len(strSaveType) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Call AddError("Unable to determine intended action.")
	Response.Clear
	Response.Redirect "default.asp"
End If

Dim intThreadID, intForumID, intThreadAuthorID
Dim strThreadSubject, strThreadBody
Dim blnThreadSignature, blnDeleteThread
Dim strMode, intReturnCode

Dim intPostID, intPostAuthorID
Dim strPostBody, blnPostSignature, blnDeletePost

Dim strPassword, intPlayerID

'-----------------------
' Threads
'-----------------------
If Request("SaveType") = "Thread" then
	intThreadID			= Trim(Request("ThreadID"))
	intForumID			= Trim(Request("ForumID"))
	strThreadSubject	= Trim(Request.Form("ThreadSubject"))
	strThreadBody		= Trim(Request.Form("ThreadBody"))
	blnThreadSignature	= Trim(Request.Form("ThreadSignature"))
	blnDeleteThread		= Trim(Request.Form("DeleteThread"))
	strMode				= Trim(Request("Mode"))
	
	If blnDeleteThread = "1" Then
		strMode = "Delete"
	End If
	
	If blnThreadSignature = "1" Then
		blnThreadSignature = 1
	Else
		blnThreadSignature = 0
	End If
	
	'' Require Login
	If Not(blnLoggedIn) Then
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Call AddError("You must login before posting.")
		Response.Clear
		Response.Redirect "/login.asp"
	End If
	
	'' Check posting privs
	If Not(HasForumAccess()) Then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Call AddError("Your posting priviledge has been revoked.")
		Response.Clear
		Response.Redirect "default.asp"
	End If
	
	If Len(strThreadSubject) = 0 AND strMode = "Add" Then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Call AddError("No empty posting.")
		Response.Clear
		Response.Redirect "default.asp"
	End If
	strThreadBody = ForumEncode(strThreadBody)
	
	Select Case strMode
		Case "Add"
			Set oCmd = Server.CreateObject("ADODB.Command")
			With oCmd
        .CommandText = "ForumSpamProtection"
        .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RetVal", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@PlayerID", adInteger, adParamInput, 4, Session("PlayerID"))
        .Parameters.Append .CreateParameter("@ActivityType", adChar, adParamInput, 1, "T")
			End With
      oCmd.Execute
      intReturnValue = oCmd(0)
      Set oCmd = Nothing

      If intReturnValue = 1 Then
				Call AddError("You may only post a new thread in the forums every 30 seconds.")
				oConn.Close
				Set oConn = Nothing
				Response.Clear
				Response.Redirect "forumdisplay.asp?forumid=" & intForumID
			End If
			strSQL = "EXECUTE ForumsNewThread "
			strSQL = strSQL & "'" & CheckString(intForumID) & "', "
			strSQL = strSQL & "'" & CheckString(strThreadSubject) & "', "
			strSQL = strSQL & "'" & Session("PlayerID") & "', "
			strSQL = strSQL & "'" & CheckString(blnThreadSignature) & "', "
			strSQL = strSQL & "'" & CheckString(strThreadBody) & "' "
			oRS.Open strSQL, oConn
			If oRS.State = 1 Then
				If Not(oRS.EOF) Then
					intReturnCode = oRS.Fields("ReturnCode").Value
					Select Case intReturnCode
						Case 0 
							' Success
							intThreadID = oRS.Fields("ThreadID").Value 
							oRS.Close 
							Set oRS = Nothing
							oConn.Close
							Set oConn = Nothing
							Response.Clear
							Response.Redirect "showthread.asp?threadid=" & intThreadID & "&forumid=" & intForumID
						Case -1
							Call AddError(oRS.Fields("ErrorMessage").Value)
							oRS.Close 
							Set oRS = Nothing
							oConn.Close
							Set oConn = Nothing
							Response.Clear
							Response.Redirect "forumdisplay.asp?forumid=" & intForumID
						Case Else
							Call AddError("There was a problem processing your request.")
							oRS.Close
							Set oRS = Nothing
							oConn.Close 
							Set oConn = Nothing
							Response.Clear
							Response.Redirect "default.asp"							
					End Select
				End IF
				oRS.Close
			Else
				Set oRS = Nothing
				oConn.Close
				Set oConn = Nothing
				Call AddError("There was an error processing your request, please try again later.")
				Response.Clear
				Response.Redirect "default.asp"
			End If
		Case "Edit"
			blnForumModerator = IsForumModerator(intForumID)
			If Not(blnSysAdmin Or blnForumModerator) Then
				intThreadAuthorID = 0
				strSQL = "SELECT ThreadAuthorID FROM tbl_threads WHERE ThreadID = '" & intThreadID & "'"
				oRS.Open strSQL, oConn
				If Not(oRS.EOF and oRS.BOF) Then
					intThreadAuthorID = oRS.Fields("ThreadAuthorID").Value
				End If
				oRS.NextRecordset 
			End If
			If blnSysAdmin Or blnForumModerator Or cStr(intThreadAuthorID & "") = cStr(Session("PlayerID") & "") Then
				strSQL = "EXECUTE ForumsEditThread "
				strSQL = strSQL & "'" & CheckString(intThreadID) & "', "
				strSQL = strSQL & "'" & CheckString(Session("uName")) & "', "
				strSQL = strSQL & "'" & CheckString(blnThreadSignature) & "', "
				strSQL = strSQL & "'" & CheckString(strThreadBody) & "' "
				oRS.Open strSQL, oConn
				If oRS.State = 1 Then
					If Not(oRS.EOF) Then
						intReturnCode = oRS.Fields("ReturnCode").Value
						Select Case intReturnCode
							Case 0 
								' Success
								oRS.Close 
								Set oRS = Nothing
								oConn.Close
								Set oConn = Nothing
								Response.Clear
								Response.Redirect "showthread.asp?threadid=" & intThreadID & "&forumid=" & intForumID
							Case -1
								Call AddError(oRS.Fields("ErrorMessage").Value)
								oRS.Close 
								Set oRS = Nothing
								oConn.Close
								Set oConn = Nothing
								Response.Clear
								Response.Redirect "forumdisplay.asp?forumid=" & intForumID
							Case Else
								'' How did we get anything else!?
								Set oRS = Nothing
								oConn.Close
								Set oConn = Nothing
								Call AddError("There was an error processing your request, please try again later.")
								Response.Clear
								Response.Redirect "default.asp"
							End Select
					End If
					oRS.NextRecordset 
				Else
					Set oRS = Nothing
					oConn.Close
					Set oConn = Nothing
					Call AddError("There was an error processing your request, please try again later.")
					Response.Clear
					Response.Redirect "default.asp"
				End If
			Else
				Set oRS = Nothing
				oConn.Close
				Set oConn = Nothing
				Call AddError("You do not have access to edit that post.")
				Response.Clear
				Response.Redirect "default.asp"
			End If	
		Case "Delete"
			strSQL = "EXECUTE ForumsDeleteThread "
			strSQL = strSQL & "'" & intThreadID & "' "
			oRS.Open strSQL, oConn
			If oRS.State = 1 Then
				If Not(oRS.EOF) Then
					intReturnCode = oRS.Fields("ReturnCode").Value
					Select Case intReturnCode
						Case 0 
							' Success
							oRS.Close 
							Set oRS = Nothing
							oConn.Close
							Set oConn = Nothing
							Response.Clear
							Response.Redirect "forumdisplay.asp?forumid=" & intForumID
						Case -1
							Call AddError(oRS.Fields("ErrorMessage").Value)
							oRS.Close 
							Set oRS = Nothing
							oConn.Close
							Set oConn = Nothing
							Response.Clear
							Response.Redirect "forumdisplay.asp?forumid=" & intForumID
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
			Else
				Set oRS = Nothing
				oConn.Close
				Set oConn = Nothing
				Call AddError("There was an error processing your request, please try again later.")
				Response.Clear
				Response.Redirect "default.asp"
			End If
		Case Else
			Call AddError("Unable to determine intended action.")
			oConn.Close
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "forumdisplay.asp?forumid=" & intForumID
	End Select	
End If

'----------------------------
' Posts
'----------------------------
If strSaveType = "Post" then
	intThreadID			= Trim(Request("ThreadID"))
	intPostID			= Trim(Request("PostID"))
	intForumID			= Trim(Request("ForumID"))
	strPostBody			= Trim(Request.Form("PostBody"))
	blnPostSignature	= Trim(Request.Form("PostSignature"))
	blnDeletePost		= Trim(Request.Form("DeletePost"))
	strMode				= Trim(Request("Mode"))
	
	If blnDeletePost = "1" Then
		strMode = "Delete"
	End If
	
	If blnPostSignature = "1" Then
		blnPostSignature = 1
	Else
		blnPostSignature = 0
	End If
	
	'' Require Login
	If Not(blnLoggedIn) Then
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Call AddError("You must login before posting.")
		Response.Clear
		Response.Redirect "/login.asp"
	End If
	
	'' Check posting privs
	If Not(HasForumAccess()) Then
		oConn.Close
		Set oConn = Nothing
		Set oRs = Nothing
		Call AddError("Your posting priviledge has been revoked.")
		Response.Clear
		Response.Redirect "default.asp"
	End If
	
	strPostBody = ForumEncode(strPostBody)
'	posterip = Request.ServerVariables("REMOTE_ADDR")
	Select Case strMode
		Case "Add"
			Set oCmd = Server.CreateObject("ADODB.Command")
			With oCmd
        .CommandText = "ForumSpamProtection"
        .ActiveConnection = oConn
        .CommandType = adCmdStoredProc
        .Parameters.Append .CreateParameter("RetVal", adInteger, adParamReturnValue)
        .Parameters.Append .CreateParameter("@PlayerID", adInteger, adParamInput, 4, Session("PlayerID"))
        .Parameters.Append .CreateParameter("@ActivityType", adChar, adParamInput, 1, "P")
			End With
      oCmd.Execute
      intReturnValue = oCmd(0)
      Set oCmd = Nothing

      If intReturnValue = 1 Then
				Call AddError("You may only post a new reply in the forums every 30 seconds.")
				oConn.Close
				Set oConn = Nothing
				Response.Clear
				Response.Redirect "forumdisplay.asp?forumid=" & intForumID
			End If

			strSQL = "EXECUTE ForumsNewPost "
			strSQL = strSQL & "'" & intThreadID & "', "
			strSQL = strSQL & "'" & intForumID & "', "
			strSQL = strSQL & "'" & Session("PlayerID") & "', "
			strSQL = strSQL & " " & blnPostSignature & ", "
			strSQL = strSQL & " " & CInt(blnSysAdmin) & ", "
			strSQL = strSQL & "'" & CheckString(strPostBody) & "' "
			oRS.Open strSQL, oConn
			If oRS.State = 1 Then
				If Not(oRS.EOF) Then
					intReturnCode = oRS.Fields("ReturnCode").Value
					Select Case intReturnCode
						Case 0 
							' Success
							oRS.Close 
							Set oRS = Nothing
							oConn.Close
							Set oConn = Nothing
							Response.Clear
							Response.Redirect "showthread.asp?threadid=" & intThreadID & "&forumid=" & intForumID & "&lastpage=1"
						Case -1
							Call AddError(oRS.Fields("ErrorMessage").Value)
							oRS.Close 
							Set oRS = Nothing
							oConn.Close
							Set oConn = Nothing
							Response.Clear
							Response.Redirect "showthread.asp?threadid=" & intThreadID & "&forumid=" & intForumID
						Case Else
							Call AddError("There was a problem processing your request.")
							oRS.Close
							Set oRS = Nothing
							oConn.Close 
							Set oConn = Nothing
							Response.Clear
							Response.Redirect "showthread.asp?threadid=" & intThreadID & "&forumid=" & intForumID
					End Select
				End If
				oRS.Close
			Else
				Set oRS = Nothing
				oConn.Close
				Set oConn = Nothing
				Call AddError("There was an error processing your request, please try again later.")
				Response.Clear
				Response.Redirect "default.asp"
			End If
		Case "Edit"
			blnForumModerator = IsForumModerator(intForumID)
			If Not(blnSysAdmin Or blnForumModerator) Then
				intPostAuthorID = 0
				strSQL = "SELECT PostAuthorID FROM tbl_posts WHERE PostID = '" & intPostID & "'"
				oRS.Open strSQL, oConn
				If Not(oRS.EOF and oRS.BOF) Then
					intPostAuthorID = oRS.Fields("PostAuthorID").Value
				End If
				oRS.NextRecordset 
			End If
			If blnSysAdmin Or blnForumModerator Or cStr(intPostAuthorID & "") = cStr(Session("PlayerID") & "") Then
				strSQL = "EXECUTE ForumsEditPost "
				strSQL = strSQL & "'" & intPostID & "', "
				strSQL = strSQL & "'" & CheckString(Session("uName")) & "', "
				strSQL = strSQL & " " & blnPostSignature & ", "
				strSQL = strSQL & "'" & CheckString(strPostBody) & "' "
				oRS.Open strSQL, oConn
				If oRS.State = 1 Then
					If Not(oRS.EOF) Then
						intReturnCode = oRS.Fields("ReturnCode").Value
						Select Case intReturnCode
							Case 0 
								' Success
								oRS.Close 
								Set oRS = Nothing
								oConn.Close
								Set oConn = Nothing
								Response.Clear
								Response.Redirect "showthread.asp?threadid=" & intThreadID & "&forumid=" & intForumID & "&lastpage=1"
							Case -1
								Call AddError(oRS.Fields("ErrorMessage").Value)
								oRS.Close 
								Set oRS = Nothing
								oConn.Close
								Set oConn = Nothing
								Response.Clear
								Response.Redirect "forumdisplay.asp?forumid=" & intForumID
							Case Else
								'' How did we get anything else!?
								Set oRS = Nothing
								oConn.Close
								Set oConn = Nothing
								Call AddError("There was an error processing your request, please try again later.")
								Response.Clear
								Response.Redirect "default.asp"
							End Select
					End If
					oRS.NextRecordset 
				Else
					Set oRS = Nothing
					oConn.Close
					Set oConn = Nothing
					Call AddError("There was an error processing your request, please try again later.")
					Response.Clear
					Response.Redirect "default.asp"
				End If
			Else
				Set oRS = Nothing
				oConn.Close
				Set oConn = Nothing
				Call AddError("You do not have access to edit that post.")
				Response.Clear
				Response.Redirect "default.asp"
			End If	
		Case "Delete"
			strSQL = "EXECUTE ForumsDeletePost "
			strSQL = strSQL & "'" & intPostID & "' "
			oRS.Open strSQL, oConn
			If oRS.State = 1 Then
				If Not(oRS.EOF) Then
					intReturnCode = oRS.Fields("ReturnCode").Value
					Select Case intReturnCode
						Case 0 
							' Success
							oRS.Close 
							Set oRS = Nothing
							oConn.Close
							Set oConn = Nothing
							Response.Clear
							Response.Redirect "showthread.asp?threadid=" & intThreadID
						Case -1
							Call AddError(oRS.Fields("ErrorMessage").Value)
							oRS.Close 
							Set oRS = Nothing
							oConn.Close
							Set oConn = Nothing
							Response.Clear
							Response.Redirect "showthread.asp?threadid=" & intThreadID
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
			Else
				Set oRS = Nothing
				oConn.Close
				Set oConn = Nothing
				Call AddError("There was an error processing your request, please try again later.")
				Response.Clear
				Response.Redirect "default.asp"
			End If
		Case Else
			Call AddError("Unable to determine intended action.")
			oConn.Close
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear
			Response.Redirect "showthread.asp?forumid=" & intForumID & "&threadid=" & intThreadID
	End Select	
End If

'--------------------------------
' Check forum password permission
'--------------------------------
If strSaveType = "GrantAccess" Then
	intThreadID	= Request.Form("ThreadID")
	intForumID	= Request.Form("ForumID")
	strPassword	= Request.Form("ForumPassword")
	intPlayerID	= Session("PlayerID")

	'' Require Login
	If Not(blnLoggedIn) Then
		oConn.Close
		Set oConn = Nothing
		Set oRS = Nothing
		Call AddError("You must login before posting.")
		Response.Clear
		Response.Redirect "/login.asp"
	End If
	
	strSQL = "EXECUTE ForumsCheckPermissions '" & intPlayerID & "', '" & intForumID & "', '" & CheckString(strPassword) & "'"
	oRS.Open strSQL, oConn
	If oRS.State = 1 Then
		If Not(oRS.EOF) Then
			intReturnCode = oRS.Fields("ReturnCode").Value
			Select Case intReturnCode
				Case 0 
					' Success
					' Call AddError(oRS.Fields("Access").Value)
					oRS.Close 
					Set oRS = Nothing
					oConn.Close
					Set oConn = Nothing
					Response.Clear
					If Len(intThreadID) > 0 Then
						Response.Redirect "showthread.asp?threadid=" & intThreadID & "&forumid=" & intForumID
					Else
						Response.Redirect "forumdisplay.asp?forumid=" & intForumID
					End If
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
					Response.Redirect "enterpassword.asp?threadid=" & intThreadID & "&forumid=" & intForumID
			End Select
		End If
		oRS.Close
	End If
End If

'--------------------------------
' Lock Thread
'--------------------------------
If strSaveType = "LockThread" Then
	intThreadID = Request.QueryString("ThreadID")
	intForumID = Request.QueryString("ForumID")
	If blnSysAdmin Or IsForumModerator(intForumID) Then
		strsql = "UPDATE tbl_threads SET ThreadLocked = 1 WHERE ThreadID = '" & intThreadID & "'"
		oConn.Execute strSQL
	End If
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "showthread.asp?forumid=" & intForumID & "&threadid=" & intThreadID 
End If

'--------------------------------
' UnLock Thread
'--------------------------------
If strSaveType = "UnlockThread" Then
	intThreadID = Request.QueryString("ThreadID")
	intForumID = Request.QueryString("ForumID")
	If blnSysAdmin Or IsForumModerator(intForumID) Then
		strsql = "UPDATE tbl_threads SET ThreadLocked = 0 WHERE ThreadID = '" & intThreadID & "'"
		oConn.Execute strSQL
	End If
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "showthread.asp?forumid=" & intForumID & "&threadid=" & intThreadID 
end if

'--------------------------------
' Sticky Thread
'--------------------------------
If strSaveType = "StickyThread" Then
	intThreadID = Request.QueryString("ThreadID")
	intForumID = Request.QueryString("ForumID")
	If blnSysAdmin Or IsForumModerator(intForumID) Then
		strsql = "UPDATE tbl_threads SET ThreadSticky = 1 WHERE ThreadID = '" & intThreadID & "'"
		oConn.Execute strSQL
	End If
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "showthread.asp?forumid=" & intForumID & "&threadid=" & intThreadID 
end if

'--------------------------------
' UnSticky Thread
'--------------------------------
if Request.QueryString("SaveType") = "UnStickyThread" then
	intThreadID = Request.QueryString("ThreadID")
	intForumID = Request.QueryString("ForumID")
	If blnSysAdmin Or IsForumModerator(intForumID) Then
		strsql = "UPDATE tbl_threads SET ThreadSticky = 0 WHERE ThreadID = '" & intThreadID & "'"
		oConn.Execute strSQL
	End If
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "showthread.asp?forumid=" & intForumID & "&threadid=" & intThreadID 
End If

'--------------------------------
' Move Thread
'--------------------------------
if strSaveType = "MoveThread" then
	intThreadID = Request.Form("ThreadID")
	intForumID = Request.Form("selForumID")
	If blnSysAdmin Then
		strsql = "EXECUTE ForumsMoveThread @ThreadID='" & intThreadID & "', @ForumID ='" & intForumID & "'"
		oConn.Execute strSQL
	End If
	'Response.Write strSQL
	'Response.End
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "showthread.asp?forumid=" & intForumID & "&threadid=" & intThreadID 
End If

oConn.Close
set oConn = nothing
set oRs = nothing
Response.Clear
Response.Redirect "default.asp"
%>
