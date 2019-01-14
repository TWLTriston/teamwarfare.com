<% Option Explicit %>
<% Response.Buffer = true %>
<%
Dim oConn, oFaq, nFaq_fldAuto, nQuestionId, SQL

Dim bgcone, bgctwo, bgc, bgcheader, bgcblack
nQuestionId = Request.QueryString("fldAuto")
If nQuestionId = "" Then
	Response.Redirect "/help/"
End If

set oConn = Server.CreateObject("ADODB.Connection")
set oFAQ = Server.CreateObject("ADODB.RecordSet")
oConn.connectionstring=Application("ConnectStr")
oConn.Open 

sql = "select question, answer from tbl_question where fldAuto=" & nQuestionId 
oFAQ.Open sql, oconn
If Not (oFAQ.EOF) Then 
	nfaq_fldAuto = Request.QueryString("faq")
	
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
	    			<td>
	    				 <p class="small" align=left>&nbsp;<b>TWL Help System</b>
	    			</td>
	    			<TD align=right>
	    				<a href="showfaq.asp?fldAuto=<%=nfaq_fldAuto%>">Back To Topics</a>
					</tD>
	    		</tr>
	    	</TABLE>
			<table border="0" width="100%">
			  <tr>
			    <td valign="top">
					<p class=small>
				  <HR class=forum>
			      <B><%=oFAQ("question").Value%></b>
				  <HR class=forum>
			      <%=Replace(oFAQ("answer").Value,vbCrlf,"<br>")%>
			    </td>
			  </tr>
			</table>
		    </td>
		  </tr>
		</table>
	    </td>
	  </tr>
	</table>
	<%
	oFAQ.Close
	%>
	</body>
	</html>
	<%
Else
	Response.Clear
	Response.Redirect "./"
End If
oConn.Close
Set oConn = Nothing
Set oFAQ = Nothing
%>