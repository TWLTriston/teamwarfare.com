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
Dim nRuleName, sName, oRSFaq
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
		Set oRS = Server.CreateObject("ADODB.Recordset")
		'Change ordering...
		Dim sID, nFldAuto
		sID = Request.QueryString("chapter_fldAuto")
		sID = CInt( sID)
		oRS.Open "select * from tbl_question where isactive = 1 and chapter_fldAuto = " & sID, oConn, 1, 3
		While Not oRS.EOF
			nFldAuto = oRS("fldAuto")
			oRS("orderingfield") = Request.Form("order" & nFldAuto)
			oRS.Update
			oRS.MoveNext
		Wend
		oRS.Close
	End If

Dim name, descr

Dim oRSQuestions
Set oRSQuestions = oConn.Execute("select * from tbl_question where chapter_fldAuto=" & Request.QueryString("chapter_fldAuto") & " order by orderingfield" )

Dim oRSChapter, sFAQName, nFaqID
Set oRSChapter = oConn.Execute ("select * from tbl_chapter where fldAuto = " & Request.QueryString("chapter_fldAuto") )
sFAQName = oRSChapter("name")
nRuleName = oRSChapter("rulename")
nFaqID = oRSChapter("faq_fldAuto")
oRSChapter.Close    
Set oRSChapter = Nothing
'Get some info on current FAQ
   
Call ContentStart("Questions For: " & sFAQName & " / " & NruleName)
%>
<%
Dim sURL
sURL = "questions.asp?chapter_fldAuto=" & Request.QueryString("chapter_fldAuto") & "&save=yes"
%>                             
<form method="POST" action="<%=sURL%>" name="FrontPage_Form1">
<table border="0" width="97%" bgcolor="#444444" cellspacing="0" cellpadding="0">
  <tr>
    <td>
                        <table border="0" width="100%" cellspacing="1" cellpadding="4">
                          <tr>
                            <th bgcolor="#000000">Question</th>
                            <th bgcolor="#000000">Answer</th>
                            <th bgcolor="#000000">Order Number</th>
                            <th bgcolor="#000000">Actions</th>
                          </tr>
<%
Dim nVal 
nVal = 1
bgc = bgcone
While Not oRSQuestions.EOF
	sName = "order" & oRSQuestions("fldAuto").Value
	if bgc=bgcone then
		bgc=bgctwo
	else
		bgc = bgcone
	end if
%>                        
                          <tr valign=top>
                            <td bgcolor="<%=bgcone%>"><%=oRSQuestions("question")%> / <%=orsquestions("rulename")%></td>
                            <td bgcolor="<%=bgctwo%>"><%=Replace(oRSQuestions("answer") & "",vbCrlf,"<br>")%></td>
                            <td bgcolor="<%=bgcone%>" align="center">
                            <input type="text" name="<%=sName%>" size="3" value="<%=nVal%>"></td>
                            <td bgcolor="<%=bgctwo%>"><a href="question.asp?action=edit&amp;fldAuto=<%=oRSQuestions("fldAuto")%>&amp;chapter_fldAuto=<%=Request.QueryString("chapter_fldAuto")%>">Modify</a> -<a href="question.asp?save=yes&amp;action=del&amp;fldAuto=<%=oRSQuestions("fldAuto")%>&amp;chapter_fldAuto=<%=Request.QueryString("chapter_fldAuto")%>"> Delete</a></td>
                          </tr>
<%
nVal = nVal + 1
oRSQuestions.MoveNext
Wend
oRSQuestions.Close
Set oRSQuestions = Nothing
%>                        
                             
<tr>
	<td colspan="4" bgcolor="#000000">
		<br />
		&nbsp;&nbsp;&nbsp;<input type="submit" value="Change question ordering" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;
      	<a href="question.asp?action=new&amp;chapter_fldAuto=<%=Request.QueryString("chapter_fldAuto")%>">Add new question</a>
      	<br />
                             
		<%
		Set oRSFAQ = oConn.Execute ("select * from tbl_faq where fldAuto = " & nFaqID )
		sFAQName = oRSFAQ("name")
		oRSFAQ.Close    
		Set oRSFAQ = Nothing
		%>
		<br />
		
		<a href="chapters.asp?faq_fldAuto=<%=nFaqId%>"><b>Back to current Topic ( <%=sFAQName%> )</b></a>
		<br />
		<a href="/help/admin/default.asp"><b>Back to Admin Help</b></a>
	</td>
</tr>
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