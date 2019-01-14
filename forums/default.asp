<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Category List"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()
Call ForumCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim blnLoggedIn, strPlayerName
blnLoggedIn = Session("LoggedIn")
strPlayerName = Session("uName")

Dim strOldCategory, strThisCategory
Dim dtmDate, intTimeZoneDifference, strDate, strTime
intTimeZoneDifference = 0

Dim strDateMask, bln24HourTime, blnVerticalBars, strColumnColor1, strColumnColor2
strDateMask = "MM-DD-YYYY"
bln24HourTime = False

If blnLoggedIn Then
	Call UpdateForumVisit()
End If

Dim strCurrentTime, strCurrentDate
Dim strVisitTime, strVisitDate
Dim intPostsCount

Call FixDate(Now(), intTimeZoneDifference, strCurrentDate, strCurrentTime, strDateMask, bln24HourTime)
Call FixDate(Session("CookieTime"), intTimeZoneDifference, strVisitDate, strVisitTime, strDateMask, bln24HourTime)
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Dim strHeaderColor, strHighlight1, strHighlight2
Dim strBGC
strHeaderColor	= bgcheader
strHighlight1	= bgcone
strHighlight2	= bgctwo

blnVerticalBars = False
If blnVerticalBars Then
	strColumnColor1 = ""
	strColumnColor2 = ""
Else
	strColumnColor1 = strHighlight1
	strColumnColor2 = strHighlight2
End If

%>
<% Call ContentStart("") %>
<% Call ForumAds() %>
<table width="100%">
<% Call ShowErrors("<TR><TD>", "</TD></TR>") %>
<tr>
	<td>
		<table BORDER="0" cellspacing="0" cellpadding="0" width="100%">
		<tr>
			<td CLASS="pageheader"><a href="default.asp">TWL Forums</A></td>
		</tr>
		</table>
	</td>
</tr>
<tr>
	<td>
		<table BORDER="0" cellspacing="0" cellpadding="0" width="100%">
		<% If Len(strVisitDate) > 0 Then %>
		<tr>
				<td ALIGN="right" CLASS="note">You last visited: <%=strVisitDate & " " & strVisitTime%></td>
		</tr>
		<% End If %>
		<tr>
			<td>&nbsp;</td>
		</tr>
		</table>	
	</td>
</tr>
<tr>
  <td>
		<table border="0" cellspacing="0" width="100%" cellpadding="0" class="cssBordered">
		<tr bgcolor="<%=strHeaderColor%>">
			<th CLASS="columnheader">&nbsp;</th>
			<th width=80% ALIGN="left" CLASS="columnheader">Forum</th>
			<th CLASS="columnheader">Posts</th>
			<th CLASS="columnheader">Threads</th>
			<th nowrap CLASS="columnheader">Last Post</th>
		</tr>
		<%
		If IsSysAdminLevel2() Then
			strSQL = "SELECT tbl_category.CategoryName, tbl_category.CategoryID, tbl_forums.ForumID, tbl_forums.ForumName, tbl_forums.ForumDescription, "
			strSQL = strSQL & " tbl_forums.ForumThreadCount, tbl_forums.ForumPostCount, tbl_forums.ForumLocked, tbl_forums.ForumLastPostTime, tbl_forums.ForumLastPosterName "
			strSQL = strSQL & " FROM tbl_category, tbl_forums "
			strSQL = strSQL & " WHERE tbl_forums.CategoryID = tbl_category.CategoryID AND tbl_category.CategoryID <> 5  AND tbl_category.CategoryID <> 11 "
			strSQL = strSQL & " ORDER BY tbl_category.CategoryOrder ASC, tbl_forums.ForumOrder ASC, tbl_forums.ForumName ASC "
		ElseIf bAnyLadderAdmin Then
			strSQL = "SELECT tbl_category.CategoryName, tbl_category.CategoryID, tbl_forums.ForumID, tbl_forums.ForumName, tbl_forums.ForumDescription, "
			strSQL = strSQL & " tbl_forums.ForumThreadCount, tbl_forums.ForumPostCount, tbl_forums.ForumLocked, tbl_forums.ForumLastPostTime, tbl_forums.ForumLastPosterName "
			strSQL = strSQL & " FROM tbl_category, tbl_forums "
			strSQL = strSQL & " WHERE tbl_forums.CategoryID = tbl_category.CategoryID "
			strSQL = strSQL & " AND tbl_category.CategoryOrder >= 0 AND tbl_category.CategoryID <> 11  "
			strSQL = strSQL & " ORDER BY tbl_category.CategoryOrder ASC, tbl_forums.ForumOrder ASC, tbl_forums.ForumName ASC  "
		ElseIf blnLoggedIn Then
			strSQL = "SELECT tbl_category.CategoryName, tbl_category.CategoryID, tbl_forums.ForumID, tbl_forums.ForumName, tbl_forums.ForumDescription, "
			strSQL = strSQL & " tbl_forums.ForumThreadCount, tbl_forums.ForumPostCount, tbl_forums.ForumLocked, tbl_forums.ForumLastPostTime, tbl_forums.ForumLastPosterName "
			strSQL = strSQL & " FROM tbl_category, tbl_forums "
			strSQL = strSQL & " WHERE tbl_forums.CategoryID = tbl_category.CategoryID "
			strSQL = strSQL & " AND tbl_category.CategoryOrder >= 0 AND tbl_category.CategoryID <> 7 AND tbl_category.CategoryID <> 11  "
			strSQL = strSQL & " ORDER BY tbl_category.CategoryOrder ASC, tbl_forums.ForumOrder ASC, tbl_forums.ForumName ASC  "
		Else
			'' Not logged in 
			strSQL = "SELECT tbl_category.CategoryName, tbl_category.CategoryID, tbl_forums.ForumID, tbl_forums.ForumName, tbl_forums.ForumDescription, "
			strSQL = strSQL & " tbl_forums.ForumThreadCount, tbl_forums.ForumPostCount, tbl_forums.ForumLocked, tbl_forums.ForumLastPostTime, tbl_forums.ForumLastPosterName "
			strSQL = strSQL & " FROM tbl_category, tbl_forums "
			strSQL = strSQL & " WHERE tbl_forums.CategoryID = tbl_category.CategoryID "
			strSQL = strSQL & " AND tbl_category.CategoryOrder >= 0 AND tbl_category.CategoryID <> 7 AND tbl_category.CategoryID <> 2 AND tbl_category.CategoryID <> 11  "
			strSQL = strSQL & " ORDER BY tbl_category.CategoryOrder ASC, tbl_forums.ForumOrder ASC, tbl_forums.ForumName ASC  "
		End If
		' Response.Write "<!-- " & strSQL & " -->"
		oRS.Open strSQL, oConn
		strOldCategory = ""
		If Not(oRS.EOF and oRS.BOF) Then
			Do While Not(oRS.EOF) 
				If (oRs.Fields("ForumID").Value = "33" AND (strPlayerName = "Triston" OR strPlayerName = "Polaris" OR strPlayerName = "rilke" OR strPlayerName = "Qing" OR strPlayerName = "Tankd" OR strPlayerName = "=db=" OR strPlayerName = "Durinia")) OR oRs.Fields("ForumID").Value <> "33" Then 
					strThisCategory = ors.fields("CategoryName").value
					If strThisCategory <> strOldCategory Then
						Response.Write	 "<TR BGCOLOR=" & bgcblack & " style=""cursor:default""><TD colspan=5><a CLASS=""category"" name=""Category" & oRS.Fields("CategoryID").Value & """ href=""default.asp#Category" & oRS.Fields("CategoryID").Value & """>"
						Response.Write server.HTMLEncode(ors.fields("CategoryName").value & "") & "</A>"
						Response.Write "</td></tr>" & vbCrLf
						strOldCategory = strThisCategory
						strBGC = strHighlight1
					End If
					Response.Write "<TR BGCOLOR=" & strBGC & " style=""cursor:default""><TD VALIGN=""top"" BGCOLOR=""" & strColumnColor1 & """ WIDTH=10 ALIGN=""center"">"
					If oRS.Fields("ForumLocked").Value Then
						Response.Write "<img src=""/images/locked.gif"" border=0 vspace=""3"" alt=""locked"" />"
					ElseIf IsNull(Session("CookieTime")) Then
						Response.Write "<img src=""/images/lighton.gif"" border=0 vspace=""3"" alt=""new posts"" />"
					Else
						If cDate(Session("CookieTime")) < ors.fields("ForumLastPostTime").value Then
							Response.Write "<img src=""/images/lighton.gif"" border=0 vspace=""3"" alt=""new posts"" />"
						Else
							Response.Write "<img src=""/images/lightoff.gif"" border=0 vspace=""3"" alt=""no new post"" />"
						End If
					End If
					Response.Write "</TD><TD BGCOLOR=""" & strColumnColor2 & """ align=left><B><a href=""forumdisplay.asp?forumid=" & ors.fields("ForumID").value & """>" & Server.HTMLEncode(ors.fields("ForumName").value) & "</a></b><br />"
					If Len(ors.fields("ForumDescription").value ) > 0 Then 
						Response.Write "<span class=""forumdescription"">" & Server.HTMLEncode (ors.fields("ForumDescription").value & "") & "</span></td>"
					End If
					Response.write "<TD BGCOLOR=""" & strColumnColor1 & """ align=center>" & ors.fields("ForumPostCount").value & "</td>"
					Response.write "<TD BGCOLOR=""" & strColumnColor2 & """ align=center>" & ors.fields("ForumThreadCount").value & "</td>" 
					If Not(IsNull(ors.fields("ForumLastPostTime").value)) Then 
						Call FixDate(ors.fields("ForumLastPostTime").value, intTimeZoneDifference, strDate, strTime, strDateMask, bln24HourTime)
						Response.Write "<TD BGCOLOR=""" & strColumnColor1 & """ align=right NOWRAP><span class=""smalldate"">" & strDate & "</span>"
						Response.Write "&nbsp;<span class=""smalltime"">" & strTime & "</span><br /><span class=""note"">by <b>" & Server.HTMLEncode (ors.fields("ForumLastPosterName").value & "") & "</b></span></td>"
					Else
						Response.Write "<TD BGCOLOR=""" & strColumnColor1 & """ class=""note"" align=""center"">Never</TD>"
					End If
					Response.Write "</TR>"
				
					If strBGC = strHighlight1 Then
						strBGC = strHighlight2
					Else
						strBGC = strHighlight1
					End If
				End If
				oRS.MoveNext 
			Loop
		End If
		oRS.Close
		%>
		</table>
     </td>
</tr>
<tr>
	<td>&nbsp;</td>
</tr>
<% Call DisplayForumLegend() %>

</table>

<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>

