<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title><%=strPageTitle%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<link rel="shortcut icon" href="/favicon.ico" >
	<%
	Call SetStyle()
	Select Case Session("StyleID")
		Case 2, 3, 4, 5, 6
			%>
			<link rel=stylesheet href="/core/style<%=Session("StyleID")%>.css" type="text/css">
			<script language="javascript" type="text/javascript" src="/include/TWLMenus<%=Session("StyleID")%>.js"></script>
			<%
		Case Else
			%>
			<link rel=stylesheet href="/core/style.css" type="text/css">
			<script language="javascript" type="text/javascript" src="/include/TWLMenus.js"></script>
			<%
	End Select
	%>

	<script language="javascript" type="text/javascript" src="/include/TWLDrawMenus.js"></script>
	<script language="javascript" type="text/javascript">
		var mi = new Array();
		
		var arrParents = new Array();
		var arrMyTWL = new Array();
		var arrForums = new Array();
		var arrComp = new Array();
		var arrRules = new Array();
		var arrOperations = new Array();
		var arrHelp = new Array();
	</script>
</head>

<body bgcolor="#000000" text="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" onload="preload();" onresize="fCenterMenus();">
<%
Call ShowBanner()
Call ShowAbsTop()
%>

