
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
	<title><%=strPageTitle%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script language="javascript" type="text/javascript">
	<!--
	var tri
	tri = new Image();
	function preload() {
		tri.src = "/images/tri.gif"
	}

	function popup(url, name, height, width, scrollbars)
	{
		var popwin;
		var opts = "toolbar=no,status=no,location=no,menubar=no,resizable=no";
		opts += ",height=" + height + ",width=" + width + ",scrollbars=" + scrollbars;
		
		popwin = window.open(url, name, opts);
		
		popwin.focus();
		
//		popwin.location = url;
		
	}
	//-->
	</script>
	<script language="JavaScript1.2" src="/core/coolmenus4.js"></script>
	<STYLE>
		/* CoolMenus 4 - default styles - do not edit */
		.clCMEvent{position:absolute; width:99%; height:99%; clip:rect(0,100%,100%,0); left:0; top:0; visibility:visible}
		.clCMAbs{position:absolute; visibility:hidden; left:0; top:0}
		/* CoolMenus 4 - default styles - end */
		  
		 /*Styles for level 0*/

		.clLevel0,.clLevel0over{position:absolute; padding:0px; font-family:arial,helvetica; font-size:12px; font-weight:bold; text-align:center;}
		.clLevel0nonIE,.clLevel0nonIEover{position:absolute; padding:2px; font-family:arial,helvetica; font-size:12px; font-weight:bold;}
		.clLevel0{background-color:#3C0000; layer-background-color:#3C0000; color:#FFFFFF;}
		.clLevel0over{background-color:#550000; layer-background-color:#550000; color:FFFD142; cursor:pointer; cursor:hand; }
		.clLevel0border{position:absolute; visibility:hidden; background-color:#550000; layer-background-color:#550000}

		/*Styles for level 1*/
		.clLevel1, .clLevel1over{position:absolute; padding:2px; font-family:arial,helvetica; font-size:11px; font-weight:bold;}
		.clLevel1{background-color:#3C0000; layer-background-color:#3C0000; color:#FFFFFF;}
		.clLevel1over{background-color:#550000; layer-background-color:#550000; color:#FFD1442; cursor:pointer; cursor:hand; }
		.clLevel1border{position:absolute; visibility:hidden; background-color:#550000; layer-background-color:#550000}

		/*Styles for level 2*/
		.clLevel2, .clLevel2over{position:absolute; padding:2px; font-family:arial,helvetica; font-size:10px; font-weight:bold;}
		.clLevel2{background-color:#3C0000; layer-background-color:#3C0000; color:#FFFFFF;}
		.clLevel2over{background-color:#550000; layer-background-color:#0099cc; color:#FFD142; cursor:pointer; cursor:hand; }
		.clLevel2border{position:absolute; visibility:hidden; background-color:#550000; layer-background-color:#550000}

		/*Styles for level 3*/
		.clLevel3, .clLevel3over{position:absolute; padding:2px; font-family:arial,helvetica; font-size:10px; font-weight:bold;}
		.clLevel3{background-color:#3C0000; layer-background-color:#3C0000; color:#FFFFFF;}
		.clLevel3over{background-color:#550000; layer-background-color:#0099cc; color:#FFD142; cursor:pointer; cursor:hand; }
		.clLevel3border{position:absolute; visibility:hidden; background-color:#550000; layer-background-color:#550000}
	</style>
</head>

<body bgcolor="#000000" text="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" ONLOAD="javascript:preload()">
<!-- #include virtual="/include/i_newmenu.asp" -->
<table width="780" border="0" cellspacing="0" cellpadding="0" height="85" ALIGN=CENTER>
	<tr height=25><td><img src="/images/spacer.gif" height="25" width="1" ALT=""></td></tr>
	<tr valign=top> 
	<%
	Dim objContentRotator, strHeader
	Set objContentRotator = Server.CreateObject( "MSWC.ContentRotator" )
	strHeader = objContentRotator.ChooseContent("/include/headers.txt")
	Set objContentRotator = Nothing
	If strHeader = "" then 
	'	strHeader = "<td background=""/images/header.jpg""><img src=""/images/spacer.gif"" alt=""Graphic compiled by: Triston"" height=50 width=780></td>"
	End If
	Response.Write strHeader
	%>
	</tr>
	<tr><TD BACKGROUND=""><img src="/images/spacer.gif" height="13" width="1"></td></tr>
</TABLE>

<TABLE width="780" border="0" cellspacing="0" cellpadding="0" ALIGN="CENTER">
	<tr><td><img src="/images/abstop.gif" WIDTH="780" HEIGHT="5" BORDER="0"></td></tr>