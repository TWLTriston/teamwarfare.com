<% Option Explicit %>
<% Response.Buffer = true %>
<%
Dim oConn, sql, oRSFAQ, sURL

Dim bgcone, bgctwo, bgc, bgcheader, bgcblack
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
    			<td>&nbsp;<b>TWL Help System</b></td>
    		</tr>
			<tr>
			    <td valign="top">
					<hr class="forum">
					<table border="0" width="100%">
					<tr>
						<td width="50%" bgcolor="#000000"><b>Section</b></td>
						<td width="50%" bgcolor="#000000"><b>Number of Topics</b></td>
					</tr>
					<%
					set oConn = Server.CreateObject("ADODB.Connection")
					oConn.connectionstring=Application("ConnectStr")
					oConn.Open 

					sql = "select max(faq.fldAuto) as faqid, max(faq.name) as name, count(*) as antal "
					sql = sql & "from tbl_faq as faq, tbl_chapter as chapter, tbl_question as question "
					sql = sql & "where question.chapter_fldAuto=chapter.fldAuto AND chapter.faq_fldAuto=faq.fldAuto AND faq.isactive = 1 and chapter.isactive = 1 and question.isactive= 1"
					sql = sql & "group by faq.fldAuto"
					Set oRSFAQ = oconn.Execute(sql )
					While Not  oRSFAQ.EOF
						sURL = "showfaq.asp?fldAuto=" & oRSFAQ("faqid")
					%>        
					<tr>
					  <td width="50%"><a href="<%=sURL%>"><%=oRSFAQ("name")%></a></td>
					  <td width="50%"><%=oRSFAQ("antal")%></td>
					</tr>
					<%
						oRSFAQ.MoveNext
					Wend
					oRSFAQ.Close
					oConn.Close
					Set oCOnn = Nothing
					 %>
	      			</table>
		    	</td>
  			</tr>
		</table>
		</TD>
	</TR>
	</Table>
</TD>
</TR>
</Table>
</body>
</html>
