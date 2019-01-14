<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
	<link REL=STYLESHEET HREF="/core/twlmenu.css" TYPE="text/css">
	<title>teamwarfare.com</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language="javascript" type="text/javascript" src="/include/TWLMenus.js"></script>
	<script language="javascript" type="text/javascript" src="/include/TWLDrawMenus.js"></script>
	<script language="javscript" type="text/javascript">

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

<div name="divContainer" id="divContainer" style="position: absolute; top: 0px; left: 0px; visibility: hidden; ">
<script language="javascript" type="text/javascript">

	fMyTWL();

fForumsAll();

fCompetition();
fRules();
fOperations();

	fMakeMenu(arrParents, 'help','','help', '', 'fPopHelp();')
fDrawMenus(arrParents, '', 0, 0);
fDrawMenus(arrMyTWL, 'mytwl', 0, 0);
fDrawMenus(arrForums, 'forums', 130);
fDrawMenus(arrComp, 'comp', 260);
fDrawMenus(arrRules, 'rules', 390);
fDrawMenus(arrOperations, 'operations', 520);
fDrawMenus(arrHelp, 'help', 650);
fCenterMenus();
</script>
</div>
</body>

</html>
