<%
For each i in request.servervariables
	response.write i & ":" & Request.ServerVariables(i) & "<br />"
Next
%>