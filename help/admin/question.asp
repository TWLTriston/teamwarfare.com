<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Help Administration"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim rulename 
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% 
if not (bAnyLadderAdmin or bSysAdmin) then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
end if
    
	If Request.QueryString("save")="yes" Then
		Dim sID
		sID = Request.QueryString("fldAuto")
		If sID = "" Then
			sID = 0
		Else
			sID = CInt( sID)
		End If
		oRS.Open "select * from tbl_question where fldAuto = " & sID, oConn, 1, 3
		Select Case Request.QueryString("action")
			Case "new"
				oRS.AddNew
				oRS("question").Value = Request.Form("question")
				oRS("rulename").Value = Request.Form("rulename")
				oRS("answer").Value = Request.Form("answer")
				oRS("orderingfield").Value = "0"
				oRS("chapter_fldAuto").Value = Request.QueryString("chapter_fldAuto")
			Case "edit"
				oRS("question").Value = Request.Form("question")
				oRS("rulename").Value = Request.Form("rulename")
				oRS("answer").Value = Request.Form("answer")
			Case "del"
				oRS("IsActive").Value = 0
				' oRS.Delete
		End Select
		oRS.Update
		Response.Redirect "questions.asp?chapter_fldAuto=" & Request.QueryString("chapter_fldAuto")
	End If

Dim name, descr

If Request.QueryString("action") = "edit" Then
	Set oRS2 = oConn.Execute( "select * from tbl_question where fldAuto = " & Request.QueryString("fldAuto") )
	name = oRS2("question").Value
	rulename = oRS2("rulename").Value
	descr = oRS2("answer").Value
	oRS2.Close
Else
End If

Dim oRSChapter, sFAQName, nChapterID, sRuleName 
Set oRSChapter = oConn.Execute ("select * from tbl_chapter where fldAuto = " & Request.QueryString("chapter_fldAuto") )
nChapterID = Request.QueryString("chapter_fldAuto")
sFAQName = oRSChapter("name")
sRuleName = oRSChapter("rulename")
oRSChapter.Close    
Set oRSChapter = Nothing

Call ContentStart("Questions For: " & sFAQName & " / " & sruleName)
%>
<%
Dim sURL
sURL = "question.asp?save=yes&action=" & Request.QueryString("action")
sURL = sURL & "&fldAuto=" & Request.QueryString("fldAuto") 
sURL = sURL & "&chapter_fldAuto=" & Request.QueryString("chapter_fldAuto") 
%>
    <form method="POST" action="<%=sURL%>" id=form1 name=form1>
    <table border="0" bgcolor="#444444" cellspacing="0" cellpadding="0">
	<tr>
		<td>
            <table border="0" width="100%" cellspacing="1" cellpadding="4">
				<tr valign=top>
					<td bgcolor="<%=bgcone%>" align="right">Question:</td>
					<td bgcolor="<%=bgcone%>"><input type="text" name="question" size="40" value="<%=Server.HTMLencode(name & "")%>"></td>
				</tr>
				<tr valign=top>
					<td bgcolor="<%=bgctwo%>" align="right">Rule Page Name:</td>
					<td bgcolor="<%=bgctwo%>"><input type="text" name="rulename" size="40" value="<%=Server.HTMLencode(rulename & "") %>" /></td>
				</tr>
				<tr valign=top>
					<td bgcolor="<%=bgcone%>" align="right">Answer:</td>
					<td bgcolor="<%=bgcone%>"><textarea rows="7" name="answer" cols="60"><%=Server.HTMLEncode(descr & "")%></textarea></td>
				</tr>
				<tr valign=top>
					<td align=center colspan=2 bgcolor="#000000"><input type="submit" value="Submit" name="B1"></td>
				</tr>
				<tr>
					<td bgcolor="#000000" colspan="2">
					<br />
					<a href="questions.asp?chapter_fldAuto=<%=nChapterID%>"><b>Back to current Topic ( <%=sFAQName%> )</b></a>
                    <br />
                    <a href="default.asp"><b>Help Admin</b></a>
			</table>                            
		</td>
	</tr>
	</table>
            </form>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS2 = Nothing
Set oRS = Nothing
%>