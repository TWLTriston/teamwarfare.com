<% Option Explicit %>
<% 
Response.Buffer = True 

Dim strPageTitle, strMapName, strMapImage
strPageTitle = "Map Close Up"

strMapName = Request.QueryString("mapname")
strMapImage = request.querystring("imgurl")
%>

<HTML>
<HEAD>
<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
<title><%=strPageTitle%></title>
</HEAD>
<body bgcolor="#000000" leftmargin="0" topmargin="00" marginwidth="000" marginheight="0000">
<table border=0 cellspacing=0 cellpadding=0 width=100% height=100%>
<tr valign=center>
<td align=center>
<img src="<%=strMapImage%>">
</td></tr></table>
</BODY>
</html>