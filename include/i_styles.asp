<%
''''''''''''''''''''''''''''''''''''''''''
' HTML Display Functions
''''''''''''''''''''''''''''''''''''''''''
Sub SetStyle()
	On Error Resume Next
	Select Case Session("StyleID")
		Case 2, 5, 8, 11
			bgcone = "#00163D"
			bgctwo = "#00102B"
			bgcheader = "#002C73"
			bgcblack = "#000000"
		Case 4, 6, 9, 12
			bgcone = "#123616"
			bgctwo = "#132C15"
			bgcheader = "#226328"
			bgcblack = "#000000"			
		Case Else
			bgcone = "#3C0000"
			bgctwo = "#2B0000"
			bgcheader = "#2B0000"
			bgcblack = "#000000"
	End Select
	On Error Goto 0
End Sub


Sub ShowBanner()
%>
<div id="divHeader">
	<ul>
		<li><a href="/" title="TeamWarfare League &trade; - Community Based Gaming">TeamWarfare League &trade; - Community Based Gaming</a></li>
		<li><a href="http://artofwarcentral.com/" target="_new" title="TeamWarfare League &trade; - Community Based Gaming">TeamWarfare League &trade; - Community Based Gaming</a></li>
	</ul>
</div>
<%
End Sub

Sub ShowAbsTop() 
%>
<div id="divMainWrapper">
	<div id="divMainSubWrapper">
		<div class="start"><hr /></div>
			<div id="divContentWrapper">

<%
End Sub

Sub ShowFooter()
	%>
		</div>
		
		<div id="divAdWrapper">
			<%
			If False Then ' Second(Now()) Mod 2 Or LCase(Request.ServerVariables("PATH_INFO")) = "/template.asp" Then 
				' TF Banner
				%>
				<strong>advertisement</strong>
				<!-- TF 160x600 JScript VAR code -->
				<script language=javascript><!--
				document.write ('<scr'+'ipt language=javascript src="http://a.tribalfusion.com/j.ad?site=TeamWarfare&adSpace=ROS&size=160x600&type=var&requestID='+((new Date()).getTime() % 2147483648) + Math.random()+'"></scr'+'ipt>');
				//-->
				</script>
				<noscript>
				   <a href=" http://a.tribalfusion.com/i.click?site=TeamWarfare&adSpace=ROS&size=160x600&requestID=1978787203" target=_blank>
				   <img src=" http://a.tribalfusion.com/i.ad?site=TeamWarfare&adSpace=ROS&size=160x600&requestID=1978787203"
				                  width=160 height=600 border=0 alt="Click Here"></a>
				</noscript>
				<br /><br />
				<!--Additional Sponsorship From
				<br /><br />
				<a href="http://www.trinitygames.com" target="_new"><img src="http://www.teamwarfare.com/img/twlhosting2.jpg" style="border: solid 1px #000;" /></a>
<!--				<br /><br />
				<a href="http://www.hypertechservers.com/" target="_new"><img src="http://www.teamwarfare.com/img/twlhosting1.jpg" style="border: solid 1px #000;" /></a>-->
				<!-- TF 160x600 JScript VAR code -->
				<% 
			Else
				%>				
				<strong>advertisement</strong>
				<% 
				If False Then ' Disabled 2/3/2012 Angell
					%>
					<script type="text/javascript"><!--
					google_ad_client = "pub-0367881922983400";
					google_ad_width = 160;
					google_ad_height = 600;
					google_ad_format = "160x600_as";
					google_ad_type = "text_image";
					//2007-02-23: Run of site
					google_ad_channel = "5741807764";
					google_color_border = "000000";
					google_color_bg = "000000";
					google_color_link = "ffd142";
					google_color_text = "ffffff";
					google_color_url = "ffd142";
					//--></script>
					<script type="text/javascript" src="http://pagead2.googlesyndication.com/pagead/show_ads.js">
					</script>	
					<%
				Else 
					%>
					<script type="text/javascript"><!--
					google_ad_client = "ca-pub-0367881922983400";
					/* Run of Site 160x600 */
					google_ad_slot = "5857193036";
					google_ad_width = 160;
					google_ad_height = 600;
					//-->
					</script>
					<script type="text/javascript" src="http://pagead2.googlesyndication.com/pagead/show_ads.js"></script>
				<% End If %>
				<br /><br />
								Additional Sponsorship From
<br /><br />
				<a href="http://artofwarcentral.com/" target="_new"><img src="http://www.teamwarfare.com/img/AOWC-TWL-2.gif" style="border: solid 1px #000;" /></a> 
				<br /><br />
				<%
			End If
			%>
			<div class="stop"><hr /></div>
			
		</div>
		<br clear="all" />
	</div>
</div>

	<%
	Dim TimeModifier
  TimeModifier = 6
	%>
	<div id="divFooter">
		TWL&reg; NA Time: <B><%=Now()%></b><br />
		TWL&reg; EU Time: <B><%=DateAdd("h",TimeModifier,Now())%></b><br />
		<div class="stop"><hr /></div>
	</div>

	<div id="divCopyright" class="noAd">
		All content &copy; TeamWarfare.com 2000-<%=Year(Now())%><br />
		TeamWarfare League&trade;<br />
		<a href="/privacy.asp">Privacy Policy</a> | <a href="/terms.asp">Terms and Conditions</a><br />
		<br />
<a target='_blank' title='Facebook' href='http://www.facebook.com/pages/TeamWarfare-League/376863472336763'><img src='http://www.teamwarfare.com/content/news/news/smfb.png' border='0'/></a>&nbsp;&nbsp;
 <a target='_blank' title='Twitter' href='http://www.twitter.com/TWLgaming'><img src='http://www.teamwarfare.com/content/news/news/smtw.png' border='0'/></a>&nbsp;&nbsp; 
<a target='_blank' title='Steam' href='http://steamcommunity.com/groups/teamwarfare'><img src='http://www.teamwarfare.com/content/news/news/smst.png' border='0'/></a>&nbsp;&nbsp; 
<a target='_blank' title='Youtube' href='http://www.youtube.com/user/TWLmedia'><img src='http://www.teamwarfare.com/content/news/news/smyt.png' border='0'/></a>&nbsp;&nbsp; 
<a target='_blank' title='Google+' href='http://plus.google.com/u/0/109783422799649386826/ '><img src='http://www.teamwarfare.com/content/news/news/smgl.png' border='0'/></a>&nbsp;&nbsp; 
<a target='_blank' title='Xfire' href='http://www.xfire.com/communities/teamwarfareleague/ '><img src='http://www.teamwarfare.com/content/news/news/smxf.png' border='0'/></a> 


	</div>
<%
End Sub

Sub ContentEnd()
%>
			</div>
		<div class="stop"><hr /></div>
	</div>
<%
End Sub

Sub ContentStart(byRef strHeaderText)
	If Len (strHeaderText) > 0 Then Response.Write "<h1>" & Server.HTMLEncode(strHeaderText) & "</h1>" & vbCrLf End If %>
		<div class="cssContent">
			<div class="start"><hr /></div>
			<div class="middlebox">
<%
End Sub

Sub ContentNewsStart(byRef strHeaderText)
	ContentStart(strHeaderText)
End Sub

Sub ContentNewsEnd()
	ContentEnd()
End Sub

Sub Content2BoxStart(byRef strHeaderText)
	If Len (strHeaderText) > 0 Then Response.Write "<h1>" & Server.HTMLEncode(strHeaderText) & "</h1>" & vbCrLf End If %>
		<div class="cssContent2Box">
			<div class="start"><hr /></div>
			<div class="leftbox">	
<%
End Sub

Sub Content2BoxMiddle()	%>
			</div>
			<div class="rightbox">	
<% 
End Sub

Sub Content2BoxEnd() %>
			</div>
			<div class="stop"><hr /></div>
		</div>	
<%
End Sub

Sub Content3BoxStart(byRef strHeaderText)
	If Len (strHeaderText) > 0 Then Response.Write "<h1>" & Server.HTMLEncode(strHeaderText) & "</h1>" & vbCrLf End If %>
			<div class="cssContent3Box">
				<div class="start"><hr /></div>
				<div class="leftbox">
<%
End Sub

Sub Content3BoxMiddle1()
%>
				</div>
				<div class="middlebox">
<%
End Sub

Sub Content3BoxMiddle2()
%>
				</div>
				<div class="rightbox">
<%
End Sub

Sub Content3BoxEnd()
%>
				</div>
				<div class="stop"><hr /></div>
			</div>
<%
End Sub

Sub Content33BoxStart(byRef strHeaderText)
	If Len (strHeaderText) > 0 Then Response.Write "<h1>" & Server.HTMLEncode(strHeaderText) & "</h1>" & vbCrLf End If %>
			<div class="cssContent33Box">
				<div class="start"><hr /></div>
				<div class="leftbox">
<%
End Sub

Sub Content33BoxMiddle()
%>
				</div>
				<div class="rightbox">
<%
End Sub

Sub Content33BoxEnd()
%>
				</div>
				<div class="stop"><hr /></div>
			</div>
<%
End Sub

Sub Content66BoxStart(byRef strHeaderText)
	If Len (strHeaderText) > 0 Then Response.Write "<h1>" & Server.HTMLEncode(strHeaderText) & "</h1>" & vbCrLf End If %>
			<div class="cssContent66Box">
				<div class="start"><hr /></div>
				<div class="leftbox">
<%
End Sub

Sub Content66BoxMiddle()
%>
				</div>
				<div class="rightbox">
<%
End Sub	

Sub Content66BoxEnd()
%>
				</div>
				<div class="stop"><hr /></div>
			</div>
<%
End Sub

Sub ForumAds()
%>
	<center>
		<script type="text/javascript"><!--
		google_ad_client = "ca-pub-0367881922983400";
		/* Forum Header */
		google_ad_slot = "7521912818";
		google_ad_width = 728;
		google_ad_height = 90;
		//-->
		</script>
		<script type="text/javascript" src="http://pagead2.googlesyndication.com/pagead/show_ads.js"></script>
	</center>
		<%
End Sub
%>