<!-- #include file="i_styles.asp" //-->
<!-- #include file="i_securitylib.asp" //-->
<%
''''''''''''''''''''''''''''''''''''''''''
' General Functions
''''''''''''''''''''''''''''''''''''''''''
Function CheckString(byVal strData)
	CheckString = strData
	CheckString = Replace(strData, "'", "''")
End Function
Function SearchString(byVal strData)
	SearchString = strData
	SearchString = Replace(SearchString, "[", "[ [ ]")
	SearchString = Replace(SearchString, "%", "[ % ]")
	SearchString = Replace(SearchString, "_", "[_]")
End Function

Function TrackSession()
	Dim strTrackIPAddress
	Dim oRSIPBan
	strTrackIPAddress = Request.ServerVariables("REMOTE_ADDR")
	If Not(Session("Tracked")) Then
		sSQL = "EXECUTE CheckIPStatus '" & CheckString(strTrackIPAddress) & "'"
		Set oRSIPBan = Server.CreateObject("ADODB.RecordSet")
		oRSIPBan.Open sSQL, oConn
		If orsIPBan.State = 1 Then
			If Not(oRSIPBan.EOF AND oRSIPBan.BOF) Then
				If (oRSIPBan.Fields("ReturnCode").Value = "Banned") Then
					oConn.Close
					Set oConn = Nothing
					Set oRSIPBan = Nothing
					Response.clear
					Response.Redirect "/banned.asp"
				End If
			End If
			oRSIPBan.Close
		End if
		Set oRSIPBan = Nothing
	End If
	If Session("LoggedIn") AND Not(Session("Tracked")) Then
		Session("Tracked") = True
		Dim sSQL
		sSQL = "INSERT INTO tbl_tracker ( "
		sSQL = sSQL & " REMOTE_ADDR, "
		sSQL = sSQL & " REMOTE_HOST, "
		sSQL = sSQL & " HTTP_USER_AGENT, "
		sSQL = sSQL & " PlayerHandle, "
		sSQL = sSQL & " TimeLogged "
		sSQL = sSQL & " ) VALUES ( "
		sSQL = sSQL & "'" & CheckString(strTrackIPAddress) & "', "
		sSQL = sSQL & "'" & CheckString(Request.ServerVariables("REMOTE_HOST")) & "', "
		sSQL = sSQL & "'" & CheckString(Request.ServerVariables("HTTP_USER_AGENT")) & "', "
		sSQL = sSQL & "'" & CheckString(Session("uName")) & "', GetDate()) "
		oConn.Execute(sSQL)
	End If
	''' Update active sessions
	sSQL = "EXECUTE UpdateSession @PlayerHandle = '" & CheckString(Session("uName")) & "', @LastPageView = '" & CheckString(Left(Request.ServerVariables ("PATH_INFO") & "?" & Request.QueryString, 500)) & "', @SessionID = '" & CheckString(Session.SessionID) & "'"
	oConn.Execute(sSQL)
End Function

Function ForumCookie()
	If Session("LoggedIn") then
		Dim sSQL, objRSs
		Set objRSs = Server.CreateObject("ADODB.RecordSet")
		CheckCookie()
		If Session("CookieTime") = "" or IsNull(session("CookieTime")) then
				sSQL = "select ForumLastVisit from tbl_players where playerhandle='" & CheckString(session("uName")) & "'"
				objRSs.Open sSQL, oConn
				if not(objRSs.bof and objRSs.eof) then
					session("CookieTime") = objRSs.fields(0).value
				end if
				objRSs.close
		end if
		sSQL = "update tbl_players set ForumLastVisit = GetDate() WHERE playerhandle='" & CheckString(session("uName")) & "'"
		oConn.execute sSQL
		Set objRSs = Nothing
	End If
End Function

Function ForumFooter()
	ForumFooter = vbcrlf & "<TR bgcolor=""#000000""><TD align=""center"" colspan=""6""><p class=small>[ <a href=""/forums/"">forum home</a> ] [ <a href=""/forums/forumcodes.asp?code=smiley"">smiley legend</a> ] [ <a href=""/forums/forumcodes.asp?code=forumcode"">forum codes</a> ]</P></TD></TR>"
End Function

Function IsForumAdmin(byVal ForumID)
	Dim sSQL, objRSs
	Set objRSs = Server.CreateObject("ADODB.RecordSet")
	CheckCookie()
	IsForumAdmin = False
	sSQL = "SELECT * FROM lnk_f_p where PlayerID='" & Session("PlayerID") & "' and forumid=" & forumid
	objRSs.Open sSQL, oConn
	if not(objRSs.bof and objRSs.eof) then
		IsForumAdmin=true
	end if
	objRSs.close
	Set objRSs = Nothing
End Function

''''''''''''''''''''''''
' Old Functions 12/1/2001
''''''''''''''''''''''''

Function MailTeamCaptains(LinkID, Text, Subject, MailAdmin, MailSys, LadderID)
	Dim sec, sec2
	Set Sec = Server.Createobject("ADODB.RecordSet")
	Set Sec2 = Server.Createobject("ADODB.RecordSet")

	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.RemoteHost  = "216.250.243.215"
	Mailer.FromName    = "Teamwarfare"
	mailer.FromAddress = "automailer@teamwarfare.com"
	Mailer.Subject     = subject
	Mailer.BodyText    = text
'	response.write "Subject: " & Subject & "<BR>"
'	response.write "Text: " & text & "<BR>"
	
	strsql= "select p.PlayerHandle, p.PlayerEmail from tbl_players p, Lnk_T_P_L lnk where p.PlayerID = lnk.PlayerID AND lnk.TLLinkID='" & LinkID & "' and lnk.isAdmin=1"
	sec.open strsql,oconn
	if not (sec.eof and sec.bof) then
		do while not(sec.eof)
			Mailer.AddRecipient sec.fields(0).value, sec.fields(1).value
			sec.movenext
		loop
	end if
	sec.close
	if mailadmin then
		strsql = "Select p.PlayerHandle, p.PlayerEmail from tbl_players p, Lnk_L_A lnk where lnk.PlayerID = p.PlayerID AND lnk.LadderID = " & ladderID
		sec.open strsql,oconn
		if not (sec.eof and sec.bof) then
			do while not(sec.eof)
				Mailer.AddRecipient sec.fields(0).value, sec.fields(1).value
				'Response.write "Recipient: " & sec2.fields(0).value & " - " & sec2.fields(1).value & "<BR>"
				sec.movenext
			loop
		end if
		sec.close
	end if
	if mailsys then
		strsql = "Select p.PlayerHandle, p.PlayerEmail from sysadmins s, tbl_players P WHERE p.PlayerID = s.AdminID and s.SendEmail = 1"
		sec.open strsql,oconn
		if not (sec.eof and sec.bof) then
			do while not(sec.eof)
				Mailer.AddRecipient sec.fields(0).value, sec.fields(1).value
				'Response.write "Recipient: " & sec2.fields(0).value & " - " & sec2.fields(1).value & "<br>"
				sec.movenext
			loop
			end if
		sec.close
	end if
	on error resume next
	Mailer.SendMail
	set mail = nothing
			
	MailTeamcaptains = true
End Function

Function MailPlayersOnLadder(LinkID, Name, Email, Text, Subject, MailAdmin, MailSys, LadderID)
	Dim sec, sec2
	Set Sec = Server.Createobject("ADODB.RecordSet")
	Set Sec2 = Server.Createobject("ADODB.RecordSet")

	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.RemoteHost  = "216.250.243.215"
	Mailer.FromName    = "Teamwarfare"
	mailer.FromAddress = "automailer@teamwarfare.com"
	Mailer.Subject     = subject
	Mailer.BodyText    = text
'	response.write "Subject: " & Subject & "<BR>"
'	response.write "Text: " & text & "<BR>"
	
	Mailer.AddRecipient Name, Email

	if mailadmin then
		strsql = "Select PlayerID from Lnk_pL_A where PlayerLadderID = " & ladderID
		sec.open strsql,oconn
		if not (sec.eof and sec.bof) then
			do while not(sec.eof)
				strsql="Select PlayerHandle, PlayerEmail from tbl_players where playerID=" & sec.fields(0).value
				sec2.open strsql,oconn
				if not (sec2.eof and sec2.bof) then
					Mailer.AddRecipient sec2.fields(0).value, sec2.fields(1).value
					'Response.write "Recipient: " & sec2.fields(0).value & " - " & sec2.fields(1).value & "<BR>"
					end if
				sec2.close
				sec.movenext
			loop
		end if
		sec.close
	end if
	if mailsys then
		strsql = "Select p.PlayerHandle, p.PlayerEmail from sysadmins s, tbl_players P WHERE p.PlayerID = s.AdminID and s.SendEmail = 1"
		sec.open strsql,oconn
		if not (sec.eof and sec.bof) then
			do while not(sec.eof)
					Mailer.AddRecipient sec.fields(0).value, sec.fields(1).value
					'Response.write "Recipient: " & sec2.fields(0).value & " - " & sec2.fields(1).value & "<br>"
				sec.movenext
			loop
			end if
		sec.close
	end if
	on error resume next
	Mailer.SendMail
	set mail = nothing
			
	MailPlayersOnLadder = true
End Function

Function AddMatchComms(text, MatchID)
	Dim sec, sec2
	Set Sec = Server.Createobject("ADODB.RecordSet")
	strsql = "select * from tbl_comms where commdead < 1 and matchID=" & matchID & " order by CommID desc"
	sec.open strsql,oconn
	if not (sec.eof and sec.bof) then
		text = text & vbcrlf & "Current Match Communications:" & vbcrlf & vbcrlf
		do while not(sec.eof)
			text = text & sec.fields("CommAuthor").value & vbcrlf & "--------------" & vbcrlf
			text = text & sec.fields("Commdate").value & vbcrlf & "--------------" & vbcrlf
			text = text & sec.fields("comms").value & vbcrlf
			sec.movenext
		loop
	else
		text = text & "No current match communications." & vbcrlf
	end if
	sec.close
	AddMatchComms = text
End Function

Function IsTeamCaptainByLinkID(TeamLinkID)
	Dim sec
	Set Sec = Server.CreateObject("ADODB.RecordSet")
	IsTeamCaptainByLinkID = False
	CheckCookie()
	membername=session("uName")
	TLLinkID=TeamLinkID
	strsql="select PlayerID from tbl_players where PlayerHandle='" & replace(membername, "'", "''") & "'"
	sec.open strsql,oconn
	if not (sec.bof and sec.eof) and TLLinkID <> "" then
		playeridseccheck=sec.fields(0).value
		sec.close
		strsql="select isadmin from lnk_T_P_L where TLLinkID=" & TLLinkID & " and PlayerID=" & playeridseccheck
		sec.open strsql,oconn
		if not (sec.eof and sec.bof) then
			if sec.fields(0).value = 1 then 
				IsTeamCaptainByLinkID=true
			end if
		end if
		sec.close
	else
		sec.close
	end if	
End Function

Function IsServerAdmin(ServerID)
	Dim sec
	Set Sec = Server.CreateObject("ADODB.RecordSet")
	IsServerAdmin = False
	CheckCookie()
	memberName = session("uName")
	strSQL = "SELECT sp.SPLinkID FROM lnk_s_p sp, tbl_players p WHERE p.PlayerID = sp.PlayerID AND p.PlayerHandle = '" & Replace(MemberName, "'", "''") & "' AND sp.ServerID ='" & ServerID & "'"
	sec.open strsql, oconn
	if not(sec.bof and sec.eof) then
		IsServerAdmin = True
	End if
	Sec.Close
End Function

Sub AddError(byVal strErrorString)
	Session("ErrorList") = Session("ErrorList") & "|" & strErrorString
End Sub

Sub ShowErrors(byVal strPrepending, byVal strAppending)
	If Len(Session("ErrorList")) > 0 Then
		If Not(IsNull(strPrepending)) Then
			Response.Write strPrepending
		End If
		Response.Write "<FONT CLASS=""error"">Error:"
		Response.Write Replace(Session("ErrorList") & "", "|", "<BR>")
		Response.Write "</FONT>"
		If Not(IsNull(strAppending)) Then
			Response.Write strAppending
		End If
		Session("ErrorList") = ""
	End If	
End Sub

Sub DisplayForumFooter()
	Response.Write "<TR>"
	Response.Write "<TD class=""littlelinks"" align=""center"" colspan=""6"">"
	If blnSysAdmin Then 
		Response.Write "<a href=""/forums/admin"">forum admin</a> / "
	End If
	Response.Write "<a href=""/forums/"">forum home</a>"
	Response.Write " / <a href=""/forums/emoticons.asp"">emoticon legend</a>"
	Response.Write " / <a href=""/forums/forumcodes.asp"">forum codes</a></TD>"
	Response.Write "</TR>"
End Sub

Sub DisplayNewForumFooter()
	Response.Write "<span class=""cssSmall"">"
	If blnSysAdmin Then 
		Response.Write "<a href=""/forums/admin"">forum admin</a> / "
	End If
	Response.Write "<a href=""/forums/"">forum home</a>"
	Response.Write " / <a href=""/forums/emoticons.asp"">emoticon legend</a>"
	Response.Write " / <a href=""/forums/forumcodes.asp"">forum codes</a></span>"
End Sub

Function HasForumAccess()
	Dim oForumRS, strSQL
	HasForumAccess = False
	If Session("LoggedIn") Then
		Set oForumRS = Server.CreateObject("ADODB.RecordSet")
		strSQL = "SELECT ForumAccess FROM tbl_players WHERE PlayerID = '" & Session("PlayerID") & "'"
		oForumRS.Open strSQL, oConn
		If Not(oForumRS.EOF AND oForumRS.BOF) Then
			HasForumAccess = cBool(oForumRS.Fields("ForumAccess").Value)
		End If
		oForumRS.Close
		Set oForumRS = Nothing 
	End If
End Function

Function IsForumModerator(byVal intForumID)
	Dim oModeratorRS, strSQL
	IsForumModerator = False
	If Session("LoggedIn") Then
		Set oModeratorRS = Server.CreateObject("ADODB.RecordSet")
		strSQL = "SELECT ForumID FROM lnk_f_p WHERE ForumID = '" & intForumID & "' AND PlayerID = '" & Session("PlayerID") & "'"
		oModeratorRS.Open strSQL, oConn
		If Not(oModeratorRS.EOF AND oModeratorRS.BOF) Then
			IsForumModerator = True
		End If
		oModeratorRS.Close 
		Set oModeratorRS = Nothing
	End If	
End Function

Function ForumEncodeOld(byVal strBody)
	Dim strSQL, oEncoderRS
	Dim strEncoded
	Dim intStartURL, intEndURL, strURL, strRightOfURL, strLeftOfURL
	Dim intCounter, intNextSpace, intNextLine, blnParse
	Dim intStartName, intEndName, strURLName
	
	' BB Codes
	strEncoded = strBody
	strEncoded = Replace(strEncoded, "[b]", "<B>", 1, -1, 1)
	strEncoded = Replace(strEncoded, "[i]", "<I>", 1, -1, 1)
	strEncoded = Replace(strEncoded, "[u]", "<U>", 1, -1, 1)
	strEncoded = Replace(strEncoded, "[/b]", "</B>", 1, -1, 1)
	strEncoded = Replace(strEncoded, "[/i]", "</I>", 1, -1, 1)
	strEncoded = Replace(strEncoded, "[/u]", "</U>", 1, -1, 1)
	strEncoded = Replace(strEncoded, "[img]", "<div style=""overflow: auto; width: 630px;""><img src=""", 1, -1, 1)
	strEncoded = Replace(strEncoded, "[/img]", """></div>", 1, -1, 1)
	strEncoded = Replace(strEncoded, "[quote]", "<blockquote class=""forumquote"">", 1, -1, 1)
	strEncoded = Replace(strEncoded, "[/quote]", "</blockquote>", 1, -1, 1)
	strEncoded = Replace(strEncoded, "cableone.net", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "redcoat.net", "BannedURL", 1, -1, 1)
'	strEncoded = Replace(strEncoded, "tribalpharmacy.com", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "commax.com", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "commax.net", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "mywebpages.comcast.net", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "orange.comax.com", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "villagephotos.com", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "0catch.com", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "shadow-lands.net", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "fordestore.com", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "msnusers.com", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "passport.com", "NoCookieHere", 1, -1, 1)
	strEncoded = Replace(strEncoded, "csports.net", "NoCookieHere", 1, -1, 1)
 	
	strEncoded = Replace(strEncoded, "<body", "******", 1, -1, 1)
	strEncoded = Replace(strEncoded, "onload", "******", 1, -1, 1)
	strEncoded = Replace(strEncoded, "onerror", "******", 1, -1, 1)
	strEncoded = Replace(strEncoded, "onreadystatechange", "******", 1, -1, 1)
	strEncoded = Replace(strEncoded, "<script", "******", 1, -1, 1)
	strEncoded = Replace(strEncoded, "<noscript", "******", 1, -1, 1)
	strEncoded = Replace(strEncoded, "<meta", "******", 1, -1, 1)
	strEncoded = Replace(strEncoded, ".cookie", "******", 1, -1, 1)
	strEncoded = Replace(strEncoded, "iframe", "******", 1, -1, 1)

	' parse the easy one [url] [/url]
	intStartURL = inStr(1, strEncoded, "[url]", 1)
	If intStartURL > 0 Then
		While intStartURL > 0
			intEndURL = inStr(intStartURL, strEncoded, "[/url]", 1)
			If intEndURL > 0 Then
				strURL = Mid(strEncoded, intStartURL + 5, intEndURL - intStartURL - 5)
				strLeftOfURL = Left(strEncoded, intStartURL - 1)
				strRightOfURL = Right(strEncoded, Len(strEncoded) - intEndURL - 5)
				If Left(lCase(strURL), 7) <> "http://" Then
					strURL = "http://" & strURL
				End If
				strEncoded = strLeftOfURL & "<a href=""" & strURL & """>" & Server.HTMLEncode(strURL & "") & "</A>" & strRightOfURL
			End If
			intStartURL = inStr(intStartURL + 1, strEncoded, "[url]", 1)
		Wend
	End If

	'' Looking for "www.teamwarfare.com" style urls
	intStartURL = inStr(1, strEncoded, "www", 1)
	If intStartURL > 0 Then
		While intStartURL > 0
			blnParse = False
			If intStartURL = 1 Then
				' First item in the post
				blnParse = True
			ElseIf Mid(strEncoded, intStartURL - 1, 1) = " " Or Mid(strEncoded, intStartURL - 1, 1) = chr(10) Then
				' beginning of line or space preceeds it
				blnParse = True
			Else
				blnParse = False
			End If
			
			If blnParse Then
				intNextSpace = inStr(intStartURL, strEncoded, " ", 1)
				intNextLine = inStr(intStartURL, strEncoded, vbCrLf)
				If (intNextLine < intNextSpace AND intNextLine > 0) Or (intNextSpace = 0 And intNextLine > 0) Then
					intEndURL = intNextLine
				ElseIf (intNextSpace < intNextLine AND intNextSpace > 0) Or (intNextLine = 0 And intNextSpace > 0) Then
					intEndURL = intNextSpace
				Else
					intEndURL = Len(strEncoded) + 1
				End If
'				Response.Write "intStartURL: " & intStartURL & "<BR><BR>"
'				Response.Write "intEndURL: " & intEndURL & "<BR><BR>"
'				Response.Write "Len(strEncoded): " & Len(strEncoded) & "<BR><BR>"
				strURL = Mid(strEncoded, intStartURL, intEndURL - intStartURL)
				strLeftOfURL = Left(strEncoded, intStartURL - 1)
				If Len(strEncoded) - intEndURL > 0 Then
					strRightOfURL = Right(strEncoded, Len(strEncoded) - intEndURL + 1)
				Else
					strRightOfURL = ""
				End If
				strURL = "http://" & strURL
				strEncoded = strLeftOfURL & "<a href=""" & strURL & """>" & Server.HTMLEncode(strURL & "") & "</A>" & strRightOfURL
			End If
			intStartURL = inStr(intStartURL + 1, strEncoded, "www", 1)
		Wend
	End If

	'' Looking for "http://www.teamwarfare.com" style url's
	intStartURL = inStr(1, strEncoded, "http://", 1)
	If intStartURL > 0 Then
		While intStartURL > 0
			blnParse = False
			If intStartURL = 1 Then
				' beginning of post
				blnParse = True
			ElseIf Mid(strEncoded, intStartURL - 1, 1) = " " Or Mid(strEncoded, intStartURL - 1, 1) = chr(10) Then
				' Either on a new line or seperated by a space
				blnParse = True
			Else
				blnParse = False
			End If
			
			If blnParse Then
				intNextSpace = inStr(intStartURL, strEncoded, " ", 1)
				intNextLine = inStr(intStartURL, strEncoded, vbCrLf)
				If (intNextLine < intNextSpace AND intNextLine > 0) Or (intNextSpace = 0 And intNextLine > 0) Then
					intEndURL = intNextLine
				ElseIf (intNextSpace < intNextLine AND intNextSpace > 0) Or (intNextLine = 0 And intNextSpace > 0) Then
					intEndURL = intNextSpace
				Else
					intEndURL = Len(strEncoded) + 1
				End If
				strURL = Mid(strEncoded, intStartURL, intEndURL - intStartURL)
				strLeftOfURL = Left(strEncoded, intStartURL - 1)
				If Len(strEncoded) - intEndURL > 0 Then
					strRightOfURL = Right(strEncoded, Len(strEncoded) - intEndURL + 1)
				Else
					strRightOfURL = ""
				End If
				strEncoded = strLeftOfURL & "<a href=""" & strURL & """>" & Server.HTMLEncode(strURL & "") & "</A>" & strRightOfURL
			End If
			intStartURL = inStr(intStartURL + 1, strEncoded, "http://", 1)
		Wend
	End If
'	Response.Write Replace(strEncoded, chr(13), "<BR>")
	' the hard one, name and url in one replace
	intStartURL = inStr(1, strEncoded, "[url=""", 1)
	If intStartURL > 0 Then
		While intStartURL > 0
			intEndURL = inStr(intStartURL, strEncoded, """]", 1)
			blnParse = False
			intStartName = intEndURL + 2
			intEndName = inStr(intStartURL, strEncoded, "[/url]", 1)
			If intEndURL > 0 And intEndName > 0 AND intEndName < Len(strEncoded) Then
				blnParse = True
			End If
			
			If blnParse Then
'				Response.Write "intStartURL:" & intStartURL & "<BR>"
'				Response.Write "intEndURL:" & intEndURL & "<BR>"
'				Response.Write "intStartName:" & intStartName & "<BR>"
'				Response.Write "intEndName:" & intEndName & "<BR>"
'				Response.Write "Len(strEncoded):" & Len(strEncoded) & "<BR>"
				strURL = Mid(strEncoded, intStartURL + 6, intEndURL - intStartURL - 6)
				strURLName = Mid(strEncoded, intStartName, intEndName - intStartName)
				strLeftOfURL = Left(strEncoded, intStartURL - 1)
				If (Len(strEncoded) - intEndName - 5) > 0 Then
					strRightOfURL = Right(strEncoded, Len(strEncoded) - intEndName - 5)
				Else 
					strRightOfURL = ""
				End If
				If Left(lCase(strURL), 7) <> "http://" Then
					strURL = "http://" & strURL
				End If
				strEncoded = strLeftOfUrl & "<a href=""" & strURL & """>" & strURLName & "</A>" & strRightOfURL
			End If
			intStartURL = inStr(intStartURL + 1, strEncoded, "[url=""", 1)
		Wend
	End If

	'' Parse body for smileys and such
	Set oEncoderRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT EmoticonSearch, EmoticonImage FROM tbl_emoticons "
	oEncoderRS.Open strSQL, oConn
	If Not(oEncoderRS.EOF and oEncoderRS.BOF) Then
		Do While Not(oEncoderRS.EOF)
			strEncoded = Replace(strEncoded, uCase(oEncoderRS.Fields("EmoticonSearch").Value), uCase(oEncoderRS.Fields("EmoticonImage").Value),1,-1,1)
			oEncoderRS.MoveNext
		Loop
	End If
	oEncoderRS.Close 
	Set oEncoderRS = Nothing
	ForumEncode = strEncoded
End Function

Function RantEncode(strBody)
	Dim oEncoderRS, strSQL, strEncoded
	strEncoded = Server.HTMLEncode(strBody)
	'' Parse body for smileys and such
	Set oEncoderRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT EmoticonSearch, EmoticonImage FROM tbl_emoticons "
	oEncoderRS.Open strSQL, oConn
	If Not(oEncoderRS.EOF and oEncoderRS.BOF) Then
		Do While Not(oEncoderRS.EOF)
			strEncoded = Replace(strEncoded, Server.HTMLEncode(uCase(oEncoderRS.Fields("EmoticonSearch").Value)), lCase(oEncoderRS.Fields("EmoticonImage").Value),1,-1,1)
			oEncoderRS.MoveNext
		Loop
	End If
	oEncoderRS.Close 
	Set oEncoderRS = Nothing
	RantEncode  = strEncoded
End Function

Sub UpdateForumVisit()
	If Session("LoggedIn") Then
		Dim strSQL
		strSQL = "EXECUTE ForumsUpdateForumVisit "
		strSQL = strSQL & "'" & Session("UserID") & "'"
		oConn.Execute(strSQL)
	End If
End Sub

Sub FixDate(byVal dtmDate, byVal intTimeZoneDiff, byRef strDate, byRef strTime, byVal strDateMask, byVal bln24HourTime)
	Dim strMonth, strDay, strYear, strHour, strMinute
	Dim strAMPM
	If IsNull(dtmDate) Or Not(IsDate(dtmDate)) Then
		Exit Sub
	End If
	dtmDate = cDate(dtmDate)
	dtmDate = DateAdd("h", intTimeZoneDiff, dtmDate)
	strMonth = Month(dtmDate)
	strDay = Day(dtmDate)
	strYear = Year(dtmDate)
	strHour = Hour(dtmDate)
	strMinute = Minute(dtmDate)
	If Not(bln24HourTime) Then
		If strHour > 12 Then
			strHour = strHour - 12
			strAMPM = "PM"
		ElseIf strHour = 0 Then
			strHour = 12
			strAMPM = "AM"
		ElseIf strHour = 12 Then
			strAMPM = "PM"
		Else 
			strAMPM = "AM"
		End If
		If Len(strHour) = 1 Then
			strHour = "0" & strHour
		End If
		If Len(strMinute) = 1 Then
			strMinute = "0" & strMinute
		End If
		strTime = strHour & ":" & strMinute & " " & strAMPM
	Else
		If Len(strHour) = 1 Then
			strHour = "0" & strHour
		End If
		If Len(strMinute) = 1 Then
			strMinute = "0" & strMinute
		End If
		strTime = strHour & ":" & strMinute
	End If					
	If Len(strMonth) = 1 Then
		strMonth = "0" & strMonth
	End If
	If Len(strDay) = 1 Then
		strDay = "0" & strDay
	End if
					
	strDate = strDateMask
	strDate = Replace(strDate, "MM", strMonth)
	strDate = Replace(strDate, "DD", strDay)
	strDate = Replace(strDate, "YYYY", strYear)
End Sub

Sub DisplayForumLegend()
	Response.Write "<tr>"
	Response.Write "<td ALIGN=CENTER CLASS=""legend""><img src=""/images/lighton.gif"" border=""0""> New posts "
	Response.Write "<img src=""/images/lightoff.gif"" border=""0""> No new posts "
	Response.Write "<img src=""/images/locked.gif"" border=""0""> A closed forum "
	Response.Write "</td>"
	Response.Write "</tr>"
End Sub

Sub DisplayThreadLegend()
	Response.Write "<tr>"
	Response.Write "<td ALIGN=CENTER CLASS=""legend""><img src=""/images/lighton.gif"" border=""0""> New posts "
	Response.Write "<img src=""/images/lightoff.gif"" border=""0""> No new posts "
	Response.Write "<img src=""/images/locked.gif"" border=""0""> Closed thread "
	Response.Write "</td>"
	Response.Write "</tr>"
End Sub

'-------------------------
'Regular Expressions here we come
'-------------------------

Function ParseURLs1(strInput)
	' Put http:// where needed
	Dim sMatchPattern, sReplacement
	sMatchPattern = "(\s|^|=|])((([a-z0-9_-]+:[a-z0-9_-]+\@)?((www|ftp|[a-z0-9]+(-\+[a-z0-9])*)\.))([a-z0"
	sMatchPattern = sMatchPattern & "-9]+(\-+[a-z0-9]+)*\.)+[a-z]{2,7}(:\d+)?(/~[a-z0-9_%\-]+)?(/[a-z0-9_%-\.]+)*(/[a-z0-9_%-]+(\.[a-z0-9"
	sMatchPattern = sMatchPattern & "]+)?(\#[a-z0-9_.]+)?)*(\?([a-z0-9_.%-]+)=[a-z0-9_.%/-]*)?(&([a-z0-9_.%-]+)=[a-z0-9_.%/-]*)*/?)"
	sReplacement = "$1http://$2"
		
	Dim oReg ' Create variable.
	Set oReg = New RegExp   ' Create a regular expression.
	oReg.Pattern = sMatchPattern   ' Set pattern.
	oReg.IgnoreCase = True   ' Set case insensitivity.
	oReg.Global = True   ' Set global applicability.

	ParseURLs1 = oReg.Replace(strInput, sReplacement)
	Set oReg = Nothing
End Function

Function ParseURLs2(strInput)
	If IsNull(strInput) Then
		ParseURLs2 = ""
	Else 
		' Convert valid url's to links
		Dim sMatchPattern, sReplacement
		sMatchPattern = "(\s|^)((((new|(ht|f)tp)s?://)([a-z0-9_-]+:[a-z0-9_-]+\@)?((www|ftp|[a-z0-9]+(-\+[a-z0-9])*)\.)?)([a-z0"
		sMatchPattern = sMatchPattern & "-9]+(\-+[a-z0-9]+)*\.)+[a-z]{2,7}(:\d+)?(/~[a-z0-9_%\-]+)?(/[a-z0-9_%-\.]+)*(/[a-z0-9_%-]+(\.[a-z0-9"
		sMatchPattern = sMatchPattern & "]+)?(\#[a-z0-9_.]+)?)*(\?([a-z0-9_.%-+]+)=[a-z0-9_.%/+-]*)?(&([a-z0-9_.%+-]+)=[a-z0-9_.%/+-]*)*/?)"
		sReplacement = "$1<a href=""$2"" target=""TWLOutput"">$2</a>"
			
		Dim oReg ' Create variable.
		Set oReg = New RegExp   ' Create a regular expression.
		oReg.Pattern = sMatchPattern   ' Set pattern.
		oReg.IgnoreCase = True   ' Set case insensitivity.
		oReg.Global = True   ' Set global applicability.
		ParseURLs2 = oReg.Replace(strInput, sReplacement)
		Set oReg = Nothing
	End If
End Function

Function ParseEmails(strInput)
	' Convert valid emails to links
	Dim sMatchPattern, sReplacement
	sMatchPattern = "(\s|^)([a-z0-9_\.-]+@([a-z0-9]+([\.\-][a-z0-9]+)*\.)+[a-z]{2,7})"
	sReplacement = "$1<a href=""mailto:$2"">$2</a>"
		
	Dim oReg, Match, Matches   ' Create variable.
	Set oReg = New RegExp   ' Create a regular expression.
	oReg.Pattern = sMatchPattern   ' Set pattern.
	oReg.IgnoreCase = True   ' Set case insensitivity.
	oReg.Global = True   ' Set global applicability.

	ParseEmails = oReg.Replace(strInput, sReplacement)
	Set oReg = Nothing
End Function

Function RegExpReplace(strInput, strMatch, strReplace)
	Dim oReg ' Create variable.
	Set oReg = New RegExp   ' Create a regular expression.
	oReg.Pattern = strMatch ' Set pattern.
	oReg.IgnoreCase = True   ' Set case insensitivity.
	oReg.Global = True   ' Set global applicability.
	RegExpReplace = oReg.Replace(strInput, strReplace)
	Set oReg = Nothing
End Function

Function ParseForumCode(strInput)
	Dim strOutput
	strOutput = strInput
	strOutput = RegExpReplace(strOutput, "(\[)(/?[biu]?)(\])", "<$2>")
	strOutput = RegExpReplace(strOutput, "(\[quote\])", "<blockquote class=""forumquote"">")
	strOutput = RegExpReplace(strOutput, "(\[quote="")(.*?)(""\])", "<blockquote class=""forumquote""><span class=""originalposter"">Originally posted by: $2</span><br />")
	strOutput = RegExpReplace(strOutput, "(\[/quote\])", "</blockquote>")
'	strOutput = RegExpReplace(strOutput, "(\<img.*>)", "<div style=""overflow: auto; width: 630px;"">$1</div>")
	strOutput = RegExpReplace(strOutput, "(\[img\])([\w\W]*?)(\[/img\])", "<div style=""overflow: auto; width: 630px;""><img src=""$2"" border=""0"" /></div>")
	strOutput = RegExpReplace(strOutput, "(\[url=([""])?)([\w\W]*?)(("")?\])([\w\W]*?)(()\[/url\])", "<a href=""$3"" target=""TWLOutput"">$6</a>")
	strOutput = RegExpReplace(strOutput, "(\[url\])([\w\W]*?)(\[/url\])", "<a href=""$2"" target=""TWLOutput"">$2</a>")
	ParseForumCode = strOutput
End Function

Function ParseCookieBans(strInput)
	Dim strOutput, strBanned
	strOutput = strInput
	strBanned = "(commax\.com)"
	strBanned = strBanned & "|(commax\.net)"
	strBanned = strBanned & "|(cableone\.net)"
	strBanned = strBanned & "|(mywebpages\.comcast\.net)"
	strBanned = strBanned & "|(orange\.comax\.com)"
	strBanned = strBanned & "|(0catch\.com)"
	strBanned = strBanned & "|(shadow\-lands\.net)"
	strBanned = strBanned & "|(villagephotos\.com)"
	strBanned = strBanned & "|(fordestore\.com)"
	strBanned = strBanned & "|(msnusers\.com)"
	strBanned = strBanned & "|(redcoat\.net)"
	strBanned = strBanned & "|(passport\.com)"
	strBanned = strBanned & "|(csports\.net)"
	strOutput = RegExpReplace(strOutput, strBanned, "NoCookieHere")
	ParseCookieBans = strOutput
End Function

Function ParseForbiddenWords(strInput) 
	Dim strOutput, strBanned
	strOutput = strInput
	strBanned = "(<(/)?((no)?script|body|meta|link|style))"
	strBanned = strBanned & "|(on(error|(un)?load|readystatechange|mouse(over|off)))"
	strBanned = strBanned & "|(\.cookie)"
	strBanned = strBanned & "|(iframe)"
	strBanned = strBanned & "|(marquee)"
	strBanned = strBanned & "|(style=)"
	strOutput = RegExpReplace(strOutput, strBanned, "**")
	ParseForbiddenWords = strOutput
End Function

Function ParseEmoticons(strInput) 
	Dim strOutput
	strOutput = strInput
	strOutput = RegExpReplace(strOutPut,"(:)(alien|transform|angel|blink|cigar|cool|cry|lol|party|threeeye|lover|mad|fu|read|roll|rotate|sex|lick|spin|sleep|withstupid|beat|rolleyes|chinese)(:)", "<img src=""smilies/$2.gif"" />")
	Dim oEncoderRS, strSQL, strEncoded
	'' Parse body for smileys and such
	Set oEncoderRS = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT EmoticonRegExp, EmoticonImage FROM tbl_emoticons WHERE EmoticonRegExp <> '' "
	oEncoderRS.Open strSQL, oConn
	If Not(oEncoderRS.EOF and oEncoderRS.BOF) Then
		Do While Not(oEncoderRS.EOF)
			If Len(oEncoderRS.Fields("EmoticonRegExp").Value) > 0 Then
				strOutput = RegExpReplace(strOutput, oEncoderRS.Fields("EmoticonRegExp").Value, oEncoderRS.Fields("EmoticonImage").Value)
			End If
			oEncoderRS.MoveNext
		Loop
	End If
	oEncoderRS.Close 
	Set oEncoderRS = Nothing
	ParseEmoticons = strOutput
End Function

Function ForumEncode2(strInput)
	Dim sOut
	sOut = strInput
	sOut = ParseURLs2(sOut)
	sOut = ParseForumCode(sOut)
	sOut = ParseEmoticons(sOut)
	sOut = ParseEmails(sOut)
	sOut = RegExpReplace(sOut, "(\n)", "<br />$1") ' Add line breaks
	ForumEncode2 = sOut
End Function

Function ForumEncode(byVal strBody)
	Dim sOut
	If IsNull(strBody) Or Len(strBody) = 0 Then
		strBody = ""
	End If
	sOut = strBody
	sOut = ParseCookieBans(sOut)
	sOut = ParseForbiddenWords(sOut)
	sOut = ParseURLs1(sOut)
	ForumEncode = sOut
End Function

Function Log2(intX)
   Log2 = Log(intX) / Log(2)
End Function

%>
