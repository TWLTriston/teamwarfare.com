<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle
strPageTitle = "TWL: " & Replace(Request.Querystring("tournament"), """", "&quot;")  & " Tournament"

Dim strSQL, oConn, oRS, oRS1, oRS2, RSr
Dim bgcone, bgctwo
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim strTournamentName, intTournamentID
strTournamentName = Request.QueryString("Tournament")

Dim strPage, strContent, intForumID, strRulesName, intHasPrizes, intHasSponsors, blnSignUp, blnLocked, intTeamsPerDiv, strContentMain
Dim chrTournamentStyle, chrFinalsStyle, strHeaderURL
strPage = lCase(Request.QueryString("page") & "")
If Len(strPage) = 0 Then
	strPage = "main"
End If

Dim intDivisionID

strSQL = "SELECT TournamentID, HasSponsors, HasPrizes, ForumID, RulesName, Signup, Locked, TeamsPerDiv, HeaderURL, TournamentStyle, FinalsStyle "
If strPage = "main" Or strPage = "sponsors" Or strPage = "prizes" Or strPage = "schedule" Then
	strSQL = strSQL & ", Content" & CheckString(strPage)
End If
strSQL = strSQL & ", ContentMain FROM tbl_tournaments WHERE TournamentName = '" & CheckString(strTournamentName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intTournamentID = oRs.Fields("TournamentID").Value
	If strPage = "main" Or strPage = "sponsors" Or strPage = "prizes" Or strPage = "schedule" Then
		strContent = oRs.Fields("Content" & strPage)
	End If
	intForumID = oRs.Fields("ForumID").Value
	strRulesName = oRs.Fields("RulesName").Value
	intHasPrizes = oRs.Fields("HasPrizes").Value
	intHasSponsors = oRs.Fields("HasSponsors").Value	
	blnSignUp = oRS.Fields("SignUp").Value
	blnLocked = oRS.Fields("Locked").Value
	strHeaderURL =  oRs.Fields("HeaderURL").Value
	intTeamsPerDiv = oRs.Fields("TeamsPerDiv").Value
	strContentMain = oRs.Fields("ContentMain").Value & ""
	chrTournamentStyle = oRs.Fields("TournamentStyle").Value
	chrFinalsStyle = oRs.Fields("FinalsStyle").Value
Else
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
oRs.NextRecordSet

If strPage = "main" AND Len(strContentMain) = 0 Then
	strPage = "brackets"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<!-- #include file="sponsors.asp" -->

<% Call ContentStart(Server.HTMLEncode(strTournamentName) & " Tournament")%>
<center>
<%
If Len(strHeaderURL) > 0 Then
	%>
	<img src="<%=strHeaderURL%>" alt="" border="0" /><br /><br />
	<%	
End If
%>
	<%
	If Len(strContentMain) > 0 Then
		If lCase(strPage) = "main" Then 
			Response.Write "<b>Main</b> / "
		Else 
			Response.Write "<a href=""default.asp?tournament=" & Server.URLENcode(strTournamentName) & """>Main</a> / "
		End If
	End If
	If lCase(strPage) = "signup" Then 
		Response.Write " <b>Sign-up</b>"
	Else 
		Response.Write " <a href=""default.asp?tournament=" & Server.URLENcode(strTournamentName) & "&page=signup"">Sign-up</a>"
	End If
	If Len(strRulesName) > 0 Then
		If lCase(strPage) = "rules" Then 
			Response.Write " / <b>Rules</b>"
		Else 
			Response.Write " / <a href=""default.asp?tournament=" & Server.URLENcode(strTournamentName) & "&page=rules"">Rules</a>"
		End If
	End If
	If intHasPrizes = "1" Then
		If lCase(strPage) = "prizes" Then 
			Response.Write " / <b>Prizes</b>"
		Else 
			Response.Write " / <a href=""default.asp?tournament=" & Server.URLENcode(strTournamentName) & "&page=prizes"">Prizes</a>"
		End If
	End If
	If intHasSponsors = "1" Then
		If lCase(strPage) = "sponsors" Then 
			Response.Write " / <b>Sponsor List</b>"
		Else 
			Response.Write " / <a href=""default.asp?tournament=" & Server.URLENcode(strTournamentName) & "&page=sponsors"">Sponsor List</a>"
		End If
	End If
	If lCase(strPage) = "schedule" Then 
		Response.Write " / <b>Schedule</b>"
	Else 
		Response.Write " / <a href=""default.asp?tournament=" & Server.URLENcode(strTournamentName) & "&page=schedule"">Schedule</a>"
	End If
	If lCase(strPage) = "brackets" Then 
		Response.Write " / <b>Brackets</b>"
	Else 
		Response.Write " / <a href=""default.asp?tournament=" & Server.URLENcode(strTournamentName) & "&page=brackets"">Brackets</a>"
	End If
	If lCase(strPage) = "teamlist" Then 
		Response.Write " / <b>Team List</b>"
	Else 
		Response.Write " / <a href=""default.asp?tournament=" & Server.URLENcode(strTournamentName) & "&page=teamlist"">Team List</a>"
	End If
	If Len(intForumID) > 0 Then 
		%>
	 	/ <a href="/forums/forumdisplay.asp?ForumID=<%=Server.URLEncode(intForumID)%>">Forums</a>
	 	<% 
	End If
	If (IsTournamentAdmin(strTournamentName) Or bSysAdmin) Then
		Response.Write " / <a href=""displayservers.asp?tournament=" & Server.URLENcode(strTournamentName) & """>Server Admin</a>"
	End If
	 %>
</center>		
<br />

<%
Select Case lCase(strPage)
	Case "rules"
		strsql = "select c.rulename as Chapter, c.fldauto, c.GeneralRuleID "
		strsql = strsql & " from tbl_chapter c "
		strsql = strsql & " where c.rulename = '" & CheckString(strRulesName) & "' "
		strsql = strsql & " ORDER BY C.orderingField "
		%>
    <table width="90%" border="0">
		<tr><td>
		<%
		Dim intGeneralRuleID, CurrentChapter, quessql
		ors.open strsql, oconn
		if ors.eof and ors.bof then
			ors.close 
			Response.Write "Unable to find requested rule set."
		else
			do while not(ors.eof)
				intGeneralRuleID = ors.fields("GeneralRuleID").value
				CurrentChapter = ors.fields("Chapter").value
				Response.Write "<P class=small><B><center>" & CurrentChapter & "</center></b></p>"
				quessql = "select q.rulename as question, q.answer from tbl_question q "
				quesSQL = quesSQL & "where q.chapter_fldauto = '" & ors.fields("fldauto").value & "' order by q.orderingfield"
		'		Response.Write quessql
				ors2.open quesSQL, oconn
				if not(ors2.eof and ors2.bof) then
					do while not(ors2.eof)
						Response.Write "<P class=small><B>" & ors2.fields("Question").value & "</B></P>"
						Response.Write "<p class=small>" & Replace(ors2("answer").Value,vbCrlf,"<br>") & "</P>"
						ors2.movenext
					loop
				end if
				ors2.nextrecordset
				ors.movenext
			loop
			ors.nextrecordset
		
				strsql = "select c.rulename as Chapter, c.fldauto "
				strsql = strsql & " from tbl_chapter c "
				strsql = strsql & " where c.fldauto = '" & intGeneralRuleID & "'"
				strsql = strsql & " ORDER BY C.orderingField "
				ors.open strsql, oconn
				if not(ors.eof and ors.bof) then
					do while not(ors.eof)
						CurrentChapter = ors.fields("Chapter").value
						Response.Write "<P class=small><B><center>" & CurrentChapter & "</center></b></p>"
						quessql = "select q.rulename as question, q.answer from tbl_question q "
						quesSQL = quesSQL & "where q.chapter_fldauto = '" & ors.fields("fldauto").value & "'"
						ors2.open quesSQL, oconn
						if not(ors2.eof and ors2.bof) then
							do while not(ors2.eof)
								Response.Write "<P class=small><B>" & ors2.fields("Question").value & "</B></P>"
								Response.Write "<p class=small>" & Replace(ors2("answer").Value,vbCrlf,"<br>") & "</P>"
								ors2.movenext
							loop
						end if
						ors2.nextrecordset
						ors.movenext
					loop
				end if
		end if
		%>
		</td>
		</tr>
		</table>
		<%
	' End Case Rules
	Case "signup"
		%>
		<!-- #include file="signup.asp" //-->
		<%
	Case "teamlist"
		%>
		<!-- #include file="teamlist.asp" //-->
		<%
	Case "brackets"
		intDivisionID = Request.QueryString("div")
		If Len(intDivisionID) = 0 Then
			intDivisionID = 1
		ElseIf Not(IsNumeric(intDivisionID)) Then
			intDivisionID = 1
		End If
		
		If (chrTournamentStyle = "S" AND cInt(intDivisionID) <> 0) OR (cInt(intDivisionID) = 0 AND chrFinalsStyle = "S") Then
			%>
			<!-- #include file="brackets.asp" //-->
			<%
		Else
			%>
			<!-- #include file="bracket_doubleelim_8.asp" //-->
			<%
		End If
	Case "main", "prizes", "sponsors", "schedule"
		%>
		<table width="97%" align=center border=0 cellpadding="0">
		  <tr>
			<td><%=strContent%></td>
		  </tr>
		</table>
		<%
End Select
%>		
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
