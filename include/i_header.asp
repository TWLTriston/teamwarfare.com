<%
Call SetStyle()
Dim strStyle
Select Case Session("StyleID")
	Case 8, 11
		strStyle = "blue"
	Case 9, 12
		strStyle = "green"
	Case Else
		'' Everyone gets red as a default
		strStyle = "red"
End Select
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
	<meta http-equiv="content-type" content="text/html; charset=iso-8859-1" />
	<title><%=strPageTitle%></title>
	
	<link rel="stylesheet" href="/css/all<% If Session("StyleID") > 9 Then Response.Write "-bigger" End If %>.css" media="all" type="text/css" />
	<link rel="stylesheet" href="/css/screen.css" media="screen" type="text/css" />
	<link rel="<% If strStyle <> "red" Then Response.Write "alternate " End If %>stylesheet" href="/css/web-red.css" media="screen" type="text/css" title="Red" />
	<link rel="<% If strStyle <> "blue" Then Response.Write "alternate " End If %>stylesheet" href="/css/web-blue.css" media="screen" type="text/css" title="Blue" />
	<link rel="<% If strStyle <> "green" Then Response.Write "alternate " End If %>stylesheet" href="/css/web-green.css" media="screen" type="text/css" title="Green" />
	<script type="text/javascript" src="/js/twl.js">	</script>

</head>

<body>
<% Call ShowBanner() %>
<!-- #include file="menus/i_menus.asp" -->
<% Call ShowAbsTop() %>