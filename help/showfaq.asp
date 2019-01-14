<% Option Explicit %>
<% Response.Buffer = true %>
<%
Dim oConn
Dim nFAQID, sql, oFaq, sCurrentSubCat
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack
nFAQId = Request.QueryString("fldAuto")
if nFAQId = "" Then
	Response.Redirect "/help/"
End If

set oConn = Server.CreateObject("ADODB.Connection")
oConn.connectionstring=Application("ConnectStr")
oConn.Open 

sql = "select question, question.fldAuto as qid, faq.name as faqname, chapter.name as chaptername "
sql = sql & "from tbl_question as question, tbl_faq as faq, tbl_chapter as chapter "
sql = sql & "where question.chapter_fldAuto=chapter.fldAuto AND chapter.faq_fldAuto=faq.fldAuto "
sql = sql & "AND faq.fldAuto = " & nFAQId & " order by chapter.orderingfield, question.orderingfield"
Set oFAQ = oconn.Execute(sql)

If Len(Session("StyleID")) = 0 Then
	Session("StyleID") = 1
End If
Call SetStyle()
%>
<!-- #include virtual="/include/i_funclib.asp" //-->
<html>
<head>
<title>TeamWarfare Help</title>
<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
</head>
<body bgcolor="#000000" leftmargin="2" topmargin="00" marginwidth="000" marginheight="0000" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">

<table width=100% height=100% cellpadding=2>
<TR valign=center>
	<TD align=center>
	<table border=1 cellspacing=0 cellpadding=0 HEIGHT=100% width=100% bordercolor="#444444" bgcolor="<%=bgcone%>">
	<TR valign=top>
		<TD>
		<table border="0" width="100%">
  			<tr>
    			<td align=left>&nbsp;<b>TWL Help System</b></td>
    			<TD align=right>
    				<a href="default.asp">Back to Main</a>
				</tD>
    		</tr>
    	</TABLE>
    	<table width=100%>
			<tr>
			  <td valign="top">
			    <p class=small><hr class="forum">
				<B>Section: <%=oFAQ("faqname")%></B>
			    <hr class="forum">
			   <% 'Present the questions...
				Response.Write "<table cellPadding=3 cellSpacing=0 width=""100%"" bgcolor="" & bgcone & "">"
				While Not oFAQ.EOF
					If Trim(sCurrentSubcat) <> Trim( oFAQ("chaptername").Value ) Then
						'Start a new one...
						If sCurrentSubcat <> "" Then
							'Avsluta gamla
							Response.Write "</UL></FONT></TD></TR>"
						End If
						Response.Write "<TR><TD bgColor=#000000 vAlign=top><p class=small><B>" & oFAQ("chaptername").Value  & "</B></p></TD></TR>"
						Response.Write "<TR><TD bgColor=" & bgcone & " vAlign=top><p class=small><UL>"
					End If
					Response.Write "<LI><p class=small><A href=""showquestion.asp?faq=" & nFAQId & "&fldAuto=" & oFAQ("qid") & """>"
					Response.Write Trim(oFAQ("question").Value) & "</a>" & "<br>"
					sCurrentSubcat = oFAQ("chaptername").Value
					oFAQ.MoveNext
				Wend
				If sCurrentSubcat <> "" Then
					Response.Write "</UL></FONT></TD></TR>"
				End If
					
				Response.Write "</table>"
				oFAQ.Close
				Set oFAQ = Nothing
			    %>
			  </td>
			</tr>
		</table>
	  </td>
	</tr>
	</table>
	</td>
</tr>
</table>
</body>
</html>
<%
oConn.Close
Set oConn = Nothing
%>