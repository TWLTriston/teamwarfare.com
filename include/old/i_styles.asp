<%
''''''''''''''''''''''''''''''''''''''''''
' HTML Display Functions
''''''''''''''''''''''''''''''''''''''''''
Sub SetStyle()
	On Error Resume Next
	Select Case Session("StyleID")
		Case 2, 5, 8
			bgcone = "#00163D"
			bgctwo = "#00102B"
			bgcheader = "#002C73"
			bgcblack = "#000000"
		Case 4, 6, 9
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
	Select Case Session("StyleID")
		Case 3, 5, 6, 7, 8, 9
			%>
			<map name="mapBanner" id="mapBanner">
				<area shape="poly" alt="TWL Hosting - Game Server Hosting" coords="549,25, 799,25, 812,4, 990,4, 990,77, 517,77, 531,66, 551,41" HREF="http://twlhosting.teamwarfare.com" target="_blank">
				<area shape="poly" alt="TeamWarfare.com" coords="493,10, 472,16, 447,36, 440,51, 445,70, 464,81, 495,82, 530,64, 547,41, 541,21, 523,11, 508,9" href="/">
				<area shape="poly" alt="TeamWarfare.com - Community Based Gaming" coords="0,4, 213,4, 221,25, 455,24, 441,42, 436,55, 441,72, 450,78, 0,78" href="/">
			</map>
			<table width="990" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr><td><img src="/images/spacer.gif" height="5" width="1" alt=""></td></tr>
				<tr> 
					<td align="center"><img src="/images/style<%=Session("StyleID")%>/header.jpg" width="990" height="87" border="0" alt="" usemap="#mapBanner" /></td>
				</tr>
				<tr><td><img src="/images/spacer.gif" height="37" width="1" alt=""></td></tr>
			</table>
			<%
		Case Else 
			%>
			<map name="mapBanner" id="mapBanner">
				<area shape="poly" alt="TeamWarfare.com - Community Based Gaming" coords="2,4, 214,4, 225,25, 347,25, 334,44, 332,58, 333,67, 344,77, 2,77, 2,4, 2,4" href="/default.asp" />
				<area shape="poly" alt="TWL Hosting - Game Server Hosting" coords="411,77, 776,76, 776,4, 597,4, 585,25, 442,25, 442,45, 434,59, 422,69" href="http://twlhosting.teamwarfare.com" target="TWLHosting" />
				<area shape="poly" alt="TeamWarfare.com" coords="345,75, 362,84, 384,84, 405,77, 420,67, 431,57, 439,44, 440,31, 433,19, 420,9, 404,6, 384,7, 366,14, 355,20, 346,30, 336,44, 334,58, 335,66, 345,75, 345,75" href="/default.asp" />
			</map>
			<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr><td><img src="/images/spacer.gif" height="5" width="1" alt=""></td></tr>
				<tr> 
					<td align="center"><img src="/images/style<%=Session("StyleID")%>/header.jpg" width="778" height="87" border="0" alt="" usemap="#mapBanner" /></td>
				</tr>
				<tr><td><img src="/images/spacer.gif" height="37" width="1" alt=""></td></tr>
			</table>
			<%
	End Select
End Sub

Sub ShowAbsTop() 
	Select Case Session("StyleID")
		Case 7, 8, 9
			%>
			<table width="990" border="0" cellspacing="0" cellpadding="0" align="center">
			<tr><td colspan="2"><img src="/images/style<%=Session("StyleID")%>/abstop.gif" width="990" height="6" border="0"></td></tr>
			<tr>
				<td rowspan="2" valign="top" width="809" background="/images/style<%=Session("StyleID")%>/contentheadlinefiller.gif">
					<!-- Main Content Here //-->
					<table border="0" cellspacing="0" cellpadding="0" width="809">
			<%
		Case 3, 5, 6
			If Session("HomePage") = "tv" Then
			%>
			<table border="0" cellspacing="0" cellpadding="0" align="center" ID="Table3"><tr><td><img src="/images/style<%=Session("StyleID")%>/twltvlogo.gif" align="center" BORDER="0"></td></tr></table>
			<table width="990" border="0" cellspacing="0" cellpadding="0" align="center" ID="Table1">
			<tr><td><img src="/images/style<%=Session("StyleID")%>/abstop.gif" WIDTH="990" HEIGHT="6" BORDER="0"></td></tr>
			<%Else%>
			<table width="990" border="0" cellspacing="0" cellpadding="0" align="center">
			<tr><td><img src="/images/style<%=Session("StyleID")%>/abstop.gif" WIDTH="990" HEIGHT="6" BORDER="0"></td></tr>
			<%End If
		Case Else 
			If Session("HomePage") = "tv" Then
			%>
			<table border="0" cellspacing="0" cellpadding="0" align="center" ID="Table4"><tr><td><img src="/images/style<%=Session("StyleID")%>/twltvlogo.gif" align="center" BORDER="0"></td></tr></table>
			<table width="780" border="0" cellspacing="0" cellpadding="0" align="center" ID="Table2">
			<tr><td><img src="/images/style<%=Session("StyleID")%>/abstop.gif" WIDTH="780" HEIGHT="6" BORDER="0"></td></tr>
			<%Else%>
			<table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
			<tr><td><img src="/images/style<%=Session("StyleID")%>/abstop.gif" WIDTH="780" HEIGHT="6" BORDER="0"></td></tr>
			<%End If
	End Select
End Sub

Sub ShowFooter()
	Dim TimeModifier
	If (Now() < DateValue("3-Apr-05 2:00:00 AM")) Then
	  TimeModifier = 7
	Else
	  TimeModifier = 6
	End If
 	Select Case Session("StyleID")
		Case 7, 8, 9
			%>
					</table>
				</td>
				<td valign="top" width="181" background="/images/style<%=Session("StyleID")%>/ad_space_filler.gif">
					<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
					<tr><td background="/images/style<%=Session("StyleID")%>/ad_empty_filler.gif"><img src="/images/spacer.gif" height="3"></td></tr>
					<tr><td background="/images/style<%=Session("StyleID")%>/ad_space_top.gif"><img src="/images/spacer.gif" height="6"></td></tr>
					<tr>
						<td align="center">
							<div id="divAdSpace2"></div>
							<br />
							<a href="/advertise.asp">Advertise with us</a>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr><td valign="bottom" background="/images/style<%=Session("StyleID")%>/ad_space_filler.gif"><img src="/images/style<%=Session("StyleID")%>/ad_space_bottom.gif" height="6"></td></tr>
			
			<tr>
			<td colspan="2">
				<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%">
				<tr>
					<td height="16" valign="middle" align="center" background="/images/style<%=Session("StyleID")%>/bottomfiller.gif" CLASS="written"> 
						TWL&reg; NA Time: <B><%=Now()%></b><br />
						TWL&reg; EU Time: <B><%=DateAdd("h",TimeModifier,Now())%></b><br />
					</td>
				</tr>
				<tr><td background="/images/style<%=Session("StyleID")%>/footer.gif"><img src="/images/spacer.gif" height="12" width="990" alt="" border="0" /></td><tr>
				<tr>
					<td align="center" valign="middle" bgcolor="#000000" CLASS="written">
						All content &copy; TeamWarfare.com 2000-<%=Year(Now())%><br />
						TeamWarfare League&trade;<br /><a href="privacy.asp">Privacy Policy</a><br /><br /> 
						<% 
						Dim strPathInfo 
						strPathInfo = LCase(Request.ServerVariables("PATH_INFO"))
						Response.Write "<!-- " & strPathInfo & " //-->"
						If strPathInfo = "/default.asp" Then 
							%>
							We support &amp; encourage you to use WiredRed <a href="http://www.wiredred.com/" title="WiredRed Web Conferencing Software">Web Conferencing</a>. 
							<% 
						End If 
						If strPathInfo = "/forums/default.asp" Then 
							%>
							We support &amp; encourage you to use WiredRed <a href="http://www.wiredred.com/" title="WiredRed Video Conferencing Software">Video Conferencing</a>. 
							<% 
						End If 
						%>
						
						<br /><br />
					</td> 
				</tr>
				</table>
			</td>
		</tr>
		</table>			
			<%
		Case 3, 5, 6
			%>
				<tr>
					<td height="16" valign="middle" align="center" background="/images/style<%=Session("StyleID")%>/contentheadlinefiller.gif" CLASS="written"> 
						TWL NA Time: <B><%=Now()%></b><br />
						TWL EU Time: <B><%=DateAdd("h",TimeModifier,Now())%></b><br />
					</td>
				</tr>
				<tr><td background="/images/style<%=Session("StyleID")%>/footer.gif"><img src="/images/spacer.gif" height="14" width="990" alt="" border="0" /></td><tr>
				<tr>
					<td align="center" valign="middle" bgcolor="#000000" CLASS="written">All content &copy; TeamWarfare.com 2000-<%=Year(Now())%><br />TeamWarfare League<br /><a href="/privacy.asp">Privacy Policy</a><br /><br /> We support &amp; encourage you to use WiredRed <a href="http://www.wiredred.com/" title="WiredRed Web Conferencing Software">Web Conferencing</a>. <br /><br /></td> 
				</tr>
			</table>
			<%
		Case Else
			%>
				<tr>
					<td height="16" valign="middle" align="center" background="/images/style<%=Session("StyleID")%>/contentheadlinefiller.gif" CLASS="written"> 
						TWL NA Time: <B><%=Now()%></b><br />
						TWL EU Time: <B><%=DateAdd("h",TimeModifier,Now())%></b><br />
					</td>
				</tr>
				<tr><td background="/images/style<%=Session("StyleID")%>/footer.gif"><img src="/images/spacer.gif" height="14" width="780" alt="" border="0" /></td><tr>
				<tr>
					<td align="center" valign="middle" bgcolor="#000000" CLASS="written">All content &copy; TeamWarfare.com 2000-<%=Year(Now())%><br />TeamWarfare League<br /><a href="/privacy.asp">Privacy Policy</a><br /><br /> We support &amp; encourage you to use WiredRed <a href="http://www.wiredred.com/" title="WiredRed Web Conferencing Software">Web Conferencing</a>. <br /><br /></td>
				</tr>
			</table>
			<%
	End Select
End Sub

Sub ContentEnd()
	Select Case Session("StyleID")
		Case 7,8,9
			%>
				</td>
			</tr>
			<tr valign="middle"><td background="/images/style<%=Session("StyleID")%>/contentend.gif"><img src="/images/spacer.gif" height="6" width="1" alt="" border="0" /></td></tr>
			<%
		Case Else
			Response.Write "</td></tr><tr>"
			Response.Write "<td background=""/images/style" & Session("StyleID") & "/contentend.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=""16""></td></tr>"
	End Select
End Sub

Sub ContentStart(byRef strHeaderText)
	Select Case Session("StyleID")
		Case 7,8,9
			If Len(strHeaderText) = 0 Then
				%>
				<tr valign="middle"><td background="/images/style<%=Session("StyleID")%>/contentheadlinefiller.gif" class="headline"><img src="/images/spacer.gif" height="3" /></td></tr>
				<tr valign="middle"><td background="/images/style<%=Session("StyleID")%>/contentstart.gif"><img src="/images/spacer.gif" height="6" width="1" alt="" border="0" /></td></tr>
				<tr>
					<td background="/images/style<%=Session("StyleID")%>/contentfiller.gif" align="center" style="padding: 0 1px 0 5px">
				<%
			Else 
				%>
				<tr valign="middle"><td background="/images/style<%=Session("StyleID")%>/contentheadlinefiller.gif" class="headline">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=strHeaderText%></td></tr>
				<tr valign="middle"><td background="/images/style<%=Session("StyleID")%>/contentstart.gif"><img src="/images/spacer.gif" height="6" width="1" alt="" border="0" /></td></tr>
				<tr>
					<td background="/images/style<%=Session("StyleID")%>/contentfiller.gif" align="center" style="padding: 0 1px 0 5px">
				<%
			End If
		Case Else
			If Len(strHeaderText) = 0 Then
				Response.Write "<TR valign=middle>"
				Response.Write "<td background=""/images/style" & Session("StyleID") & "/contentheadlinefiller.gif"" CLASS=""headline"">"
				Response.Write "<IMG SRC=""/images/spacer.gif"" BORDER=0 WIDTH=1 HEIGHT=10></td></tr>"
				Response.Write "<tr height=16 valign=middle><td background=""/images/style" & Session("StyleID") & "/contentstart.gif"">"
				Response.Write "<img src=""/images/spacer.gif"" height=16></td></tr>"
				Response.Write "<tr><td align=""center"" background=""/images/style" & Session("StyleID") & "/contentfiller.gif"">"
			Else
				Response.Write "<TR height=23 valign=middle>"
				Response.Write "<td background=""/images/style" & Session("StyleID") & "/contentheadlinefiller.gif"" CLASS=""headline"">"
				Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & strHeaderText & "</td></tr>"
				Response.Write "<tr height=16 valign=middle><td background=""/images/style" & Session("StyleID") & "/contentstart.gif"">"
				Response.Write "<img src=""/images/spacer.gif"" height=16></td></tr>"
				Response.Write "<tr><td align=""center"" background=""/images/style" & Session("StyleID") & "/contentfiller.gif"">"
		End If
	End Select
End Sub

Sub ContentNewsStart(byRef strHeaderText)
	Select Case Session("StyleID")
		Case 7,8,9
			%>
			<tr valign="middle"><td background="/images/style<%=Session("StyleID")%>/contentheadlinefiller.gif" class="headline" style="padding: 3px 0 3px 25px;"><%=strHeaderText%></td></tr>
			<tr valign="middle"><td background="/images/style<%=Session("StyleID")%>/contentstart.gif"><img src="/images/spacer.gif" height="6" width="1" alt="" border="0" /></td></tr>
			<tr>
				<td background="/images/style<%=Session("StyleID")%>/contentfiller.gif" align="center">
					<table width="97%" border="0" cellspacing="0" cellpadding="2" BACKGROUND="">
			<%
		Case Else
	    Response.Write "<TR height=23 valign=middle>"
	    Response.Write "<td background=""/images/style" & Session("StyleID") & "/contentheadlinefiller.gif"" CLASS=""headline"" align=right>"
	    Response.Write strHeaderText & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td></tr>"
			Response.Write "<tr height=16 valign=middle><td background=""/images/style" & Session("StyleID") & "/contentstart.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=16></td></tr>"
			Response.Write "<tr><td align=""center"" background=""/images/style" & Session("StyleID") & "/contentfiller.gif"">"
			Response.Write "<table width=""97%"" border=""0"" cellspacing=""0"" cellpadding=""2"" BACKGROUND="""">"
	End Select
End Sub

Sub ContentNewsEnd()
	Select Case Session("StyleID")
		Case 7,8,9
			%>
					</table>
				</td>
			</tr>
			<tr valign="middle"><td background="/images/style<%=Session("StyleID")%>/contentend.gif"><img src="/images/spacer.gif" height="6" width="1" alt="" border="0" /></td></tr>
			<%			
		Case Else	
			Response.Write "</table>"
			Response.Write "</td></tr><tr>"
			Response.Write "<td background=""/images/style" & Session("StyleID") & "/contentend.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=""16""></td></tr>"
	End Select
End Sub

Sub Content2BoxStart(byRef strHeaderText)
	If Len(strHeaderText) > 0 Then
	  Response.Write "<tr height=""23"" valign=""middle"">"
	  Response.Write "<td background=""/images/style" & Session("StyleID") & "/contentheadlinefiller.gif"" CLASS=""headline"">"
	  Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & strHeaderText & "</td></tr>"
	Else
	  Response.Write "<TR height=3 valign=middle><td background=""/images/style" & Session("StyleID") & "/contentheadlinefiller.gif""><img src=""/images/spacer.gif"" height=""3"" border=""0"" /></td></tr>"
	End If
	Select Case Session("StyleID")
		Case 7, 8, 9
			%>
			<tr valign="middle"><td background="images/style<%=Session("StyleID")%>/content2boxtop.gif"><img src="images/spacer.gif" height="6" width="1" alt="" border="0" /></td></tr>
			<tr>
				<td align="center" background="/images/style<%=Session("StyleID")%>/content2boxfiller.gif">
					<table width=100% border="0" cellspacing="0" cellpadding="0" background="">
					<tr>
						<td><img src="/images/spacer.gif" width="5" height="1"></td>
						<td width="396" valign="top">
			<%			
		Case Else
			Response.Write "<tr height=16 valign=middle><td background=""/images/style" & Session("StyleID") & "/content2boxtop.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=16></td></tr>"
			Response.Write "<tr><td align=""center"" background=""/images/style" & Session("StyleID") & "/content2boxfiller.gif"">"
			Response.Write "<table width=100% border=""0"" cellspacing=""0"" cellpadding=""0"" BACKGROUND="""">"
			Response.Write "<tr><td><img src=""/images/spacer.gif"" width=""5"" height=""1""></td>"
			
			Select Case Session("StyleID")
				Case 3, 5, 6
					Response.Write "<td width=484 valign=""top"">"
				Case Else
					Response.Write "<td width=380 valign=""top"">"
			End Select
	End Select
End Sub

Sub Content2BoxMiddle()
	Select Case Session("StyleID")
		Case 7, 8, 9
			%>
			</td>
			<td><img src="/images/spacer.gif" width="11" height="1"></td>
			<td width="396" valign="top">
			<%
		Case 3, 5, 6
			Response.Write "</td><td><img src=""/images/spacer.gif"" width=""10"" height=""1""></td>"
			Response.Write "<td width=485 valign=""top"">"
		Case Else
			Response.Write "</td><td><img src=""/images/spacer.gif"" width=""10"" height=""1""></td>"
			Response.Write "<td width=379 valign=""top"">"
	End Select
	
End Sub

Sub Content2BoxEnd()
	Select Case Session("StyleID")
		Case 7, 8, 9
			%>
						</td>
						<td><img src="/images/spacer.gif" width="1" height="1"></td>
					</tr>
					</table>
				</td>
			</tr>
			<tr><td background="/images/style<%=Session("StyleID")%>/content2boxbottom.gif"><img src="/images/spacer.gif" height="6" width="1" alt="" border="0" /></td></tr>
			<%
		Case Else
			Response.Write "</td><td><img src=""/images/spacer.gif"" width=""5"" height=""1""></td></tr></table>"
			Response.Write "</td></tr><tr>"
			Response.Write "<td background=""/images/style" & Session("StyleID") & "/content2boxbottom.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=""16""></td></tr>"
	End Select
End Sub

Sub Content3BoxStart(byRef strHeaderText)
  Response.Write "<TR height=23 valign=middle>"
  Response.Write "<td background=""/images/style" & Session("StyleID") & "/contentheadlinefiller.gif"" CLASS=""headline"">"
  Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & strHeaderText & "</td></tr>"

	Select Case Session("StyleID")
		Case 7, 8, 9
%>
		<tr valign="middle"><td background="/images/style<%=Session("StyleID")%>/content3boxtop.gif"><img src="/images/spacer.gif" height=6></td></tr>
		<tr>
			<td align="center" background="/images/style<%=Session("StyleID")%>/content3boxfiller.gif">
				<table width="100%" border="0" cellspacing="0" cellpadding="0" background="">
				<tr valign=top>
					<td><img src="/images/spacer.gif" width="5" height="1"></td>
					<td width="260" valign="top">
<%
		Case Else 
			Response.Write "<tr height=16 valign=middle><td background=""/images/style" & Session("StyleID") & "/content3boxtop.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=16></td></tr>"
			Response.Write "<tr><td align=""center"" background=""/images/style" & Session("StyleID") & "/content3boxfiller.gif"">"
			Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" background=""""><tr valign=top><td><img src=""/images/spacer.gif"" width=""5"" height=""1""></td>"
		
			Select Case Session("StyleID")
				Case 3, 5, 6
					Response.Write "<td width=320 valign=""top"">"
				Case Else
					Response.Write "<td width=250 valign=""top"">"
			End Select
	End Select
End Sub

Sub Content3BoxMiddle1()
	Response.Write "</td><td><img src=""/images/spacer.gif"" width=""10"" height=""1""></td>"

	Select Case Session("StyleID")
		Case 7, 8, 9
			Response.Write "<td width=""262"" valign=""top"">"
		Case 3, 5, 6
			Response.Write "<td width=320 valign=""top"">"
		Case Else
			Response.Write "<td width=250 valign=""top"">"
	End Select
End Sub

Sub Content3BoxMiddle2()
	Response.Write "</td><td><img src=""/images/spacer.gif"" width=""10"" height=""1""></td>"

	Select Case Session("StyleID")
		Case 7, 8, 9
			Response.Write "<td width=""261"" valign=""top"">"
		Case 3, 5, 6
			Response.Write "<td width=320 valign=""top"">"
		Case Else
			Response.Write "<td width=250 valign=""top"">"
	End Select
End Sub

Sub Content3BoxEnd()
	Select Case Session("StyleID")
		Case 7, 8, 9
%>
					</td>
					<td><img src="images/spacer.gif" width="1" height="1"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr><td background="/images/style<%=Session("StyleID")%>/content3boxbottom.gif"><img src="/images/spacer.gif" height="6" width="1" alt="" border="0" /></td></tr>
<%
		Case Else
			Response.Write "</td><td><img src=""/images/spacer.gif"" width=""5"" height=""1""></td></tr></table>"
			Response.Write "</td></tr><tr>"
			Response.Write "<td background=""/images/style" & Session("StyleID") & "/content3boxbottom.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=""16""></td></tr>"
	End SElect
End Sub

Sub Content33BoxStart(byRef strHeaderText)
		if Len(strHeaderText) > 0 Then
	    Response.Write "<TR height=23 valign=middle>"
	  Else 
	    Response.Write "<TR valign=middle>"
	  End If
    Response.Write "<td background=""/images/style" & Session("StyleID") & "/contentheadlinefiller.gif"" CLASS=""headline"">"
		if Len(strHeaderText) > 0 Then
    	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & strHeaderText
    End If
    Response.Write "</td></tr>"
	Select Case Session("StyleID")
		Case 7, 8, 9
%>
		<tr valign=middle><td background="/images/style<%=Session("StyleID")%>/content33boxtop.gif"><img src="/images/spacer.gif" height=6></td></tr>
		<tr>
			<td align="center" background="/images/style<%=Session("StyleID")%>/content33boxfiller.gif">
				<table width="100%" border="0" cellspacing="0" cellpadding="0" background="">
				<tr valign="top">
					<td width="5"><img src="/images/spacer.gif" width="5" height="1"></td>
					<td width="260" valign="top">
<%
		Case Else		    
			Response.Write "<tr height=16 valign=middle><td background=""/images/style" & Session("StyleID") & "/content33boxtop.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=16></td></tr>"
			Response.Write "<tr><td align=""center"" background=""/images/style" & Session("StyleID") & "/content33boxfiller.gif"">"
			Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" background="""">"
			Response.Write "<tr valign=top>"
			Response.Write "<td><img src=""/images/spacer.gif"" width=""5"" height=""1""></td>"
		
			Select Case Session("StyleID")
				Case 3, 5, 6
					Response.Write "<td width=320 valign=""top"">"
				Case Else
					Response.Write "<td width=250 valign=""top"">"
			End Select
	End Select
End Sub

Sub Content33BoxMiddle()
	Response.Write "</td><td><img src=""/images/spacer.gif"" width=""10"" height=""1""></td>"
	
	Select Case Session("StyleID")
		Case 7, 8, 9
			Response.Write "<td width=""533"" valign=""top"">"
		Case 3, 5, 6
			Response.Write "<td width=650 valign=""top"">"
		Case Else
			Response.Write "<td width=510 valign=""top"">"
	End Select
End Sub

Sub Content33BoxEnd()
	Select Case Session("StyleID")
		Case 7, 8, 9
		%>
					</td>
					<td width="1"><img src="/images/spacer.gif" width="1" height="1"></td>
				</tr>
				</table>
			</td>
		</tr>
		<tr><td background="/images/style<%=Session("StyleID")%>/content33boxbottom.gif"><img src="/images/spacer.gif" height="6"></td></tr>
		<%
		Case Else		
			Response.Write "</td><td width=5><img src=""/images/spacer.gif"" width=""5"" height=""1""></td></tr></table>"
			Response.Write "</td></tr><tr>"
			Response.Write "<td background=""/images/style" & Session("StyleID") & "/content33boxbottom.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=""15""></td></tr>"
	End Select
End Sub

Sub Content66BoxStart(byRef strHeaderText)
    Response.Write "<TR height=23 valign=middle>"
    Response.Write "<td background=""/images/style" & Session("StyleID") & "/contentheadlinefiller.gif"" CLASS=""headline"">"
    Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & strHeaderText & "</td></tr>"
   
	Select Case Session("StyleID")
		Case 7, 8, 9
		%>
   		<tr valign=middle><td background="/images/style<%=Session("StyleID")%>/content66boxtop.gif"><img src="/images/spacer.gif" height=6></td></tr>
		<tr>
			<td align="center" background="/images/style<%=Session("StyleID")%>/content66boxfiller.gif">
				<table width="100%" border="0" cellspacing="0" cellpadding="0" background="">
				<tr valign="top">
					<td width="5"><img src="/images/spacer.gif" width="5" height="1"></td>
					<td width="532" valign="top">
		<%
		Case Else
		    
			Response.Write "<tr height=16 valign=middle><td background=""/images/style" & Session("StyleID") & "/content66boxtop.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=16></td></tr>"
			Response.Write "<tr><td align=""center"" background=""/images/style" & Session("StyleID") & "/content66boxfiller.gif"">"
			Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" background="""">"
			Response.Write "<tr valign=top><td width=5><img src=""/images/spacer.gif"" width=""5"" height=""1""></td>"
			
			Select Case Session("StyleID")
				Case 3, 5, 6
					Response.Write "<td width=650 valign=""top"">"
				Case Else
					Response.Write "<td width=510 valign=""top"">"
			End Select
	End Select
End Sub

Sub Content66BoxMiddle()
	Response.Write "</td><td width=""10""><img src=""/images/spacer.gif"" width=""10"" height=""1""></td>"

	Select Case Session("StyleID")
		Case 7, 8, 9
			Response.Write "<td width=""261"" valign=""top"">"
		Case 3, 5, 6
			Response.Write "<td width=320 valign=""top"">"
		Case Else
			Response.Write "<td width=250 valign=""top"">"
	End Select
End Sub	

Sub Content66BoxEnd()
	Select Case Session("StyleID")
		Case 7, 8, 9
		%>
						</td>
						<td><img src="/images/spacer.gif" width="1" height="1"></td>
					</tr>
					</table>
				</td>
			</tr>
			<tr><td background="/images/style<%=Session("StyleID")%>/content66boxbottom.gif"><img src="/images/spacer.gif" height="6"></td></tr>
		<%
		Case Else
			Response.Write "</td><td width=5><img src=""/images/spacer.gif"" width=""5"" height=""1""></td></tr></table>"
			Response.Write "</td></tr><tr>"
			Response.Write "<td background=""/images/style" & Session("StyleID") & "/content66boxbottom.gif"">"
			Response.Write "<img src=""/images/spacer.gif"" height=""15""></td></tr>"
	End Select
End Sub
%>