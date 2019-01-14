<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Rules"

Dim strSQL, oConn, oRS, oRS2, oRs3, oRs4
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")
Set oRS3 = Server.CreateObject("ADODB.RecordSet")
Set oRS4 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim intRuleParentID, strRuleSet
intRuleParentID = Request.QueryString("ruleset")
strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE MenuID = " & intRuleParentID & " ORDER BY SortOrder, ShowMenuName "
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	If oRs.Fields("ParentMenuID").Value <> 2 AND intRuleParentID <> 2 Then
		intRuleParentID = 2
		strRuleSet = "rules"
	Else
		strRuleSet = oRs.Fields("ShowMenuName").Value
	End If
End If
oRs.NextRecordSet
If strRuleSet = "rules" Then
	strRuleSet = "TeamWarfare "
Else
	strRuleSet = strRuleSet & " "
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart(strRuleSet & "Rules") %>
<table border="0" cellspacing="0" cellpadding="0" width="70%" bgcolor="#444444">
<tr>
	<td>
		<table border="0" cellspacing="1" cellpadding="4" width="100%">
		<tr>
			<th bgcolor="#000000" align="center">Choose a ruleset</th>
		</tr>
			<%
			strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & intRuleParentID & " ORDER BY SortOrder, ShowMenuName "
			oRs2.Open strSQL, oConn
			If Not(oRs2.EOF AND oRs2.BOF) Then
				Do While Not(oRs2.EOF)
					Response.Write DisplayMenu(oRs2.Fields("ShowMenuName").Value, oRs2.Fields("MenuID").Value, oRs2.Fields("ParentMenuID").Value, oRs2.Fields("LinkURL").Value, bgcone, "")
					strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & oRs2.Fields("MenuID").Value & " ORDER BY SortOrder, ShowMenuName "
					oRs3.Open strSQL, oConn
					If Not(oRs3.EOF AND oRs3.BOF) Then
						Do While Not(oRs3.EOF)
							Response.Write DisplayMenu(ors3.Fields("ShowMenuName").Value, ors3.Fields("MenuID").Value, ors3.Fields("ParentMenuID").Value, ors3.Fields("LinkURL").Value, bgctwo, "&nbsp;&nbsp;")
							strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & oRs3.Fields("MenuID").Value & " ORDER BY SortOrder, ShowMenuName "
							oRs4.Open strSQL, oConn
							If Not(oRs4.EOF AND oRs4.BOF) Then
								Do While Not(oRs4.EOF)
									Response.Write DisplayMenu(ors4.Fields("ShowMenuName").Value, ors4.Fields("MenuID").Value, ors4.Fields("ParentMenuID").Value, ors4.Fields("LinkURL").Value, "#000000", "&nbsp;&nbsp;&nbsp;&nbsp;")
									oRs4.MoveNext
								Loop
							End If
							oRs4.NextRecordSet
							oRs3.MoveNext
						Loop
					End If
					oRs3.NextRecordSet
					If strRuleSet = "TeamWarfare " Then
						Response.Write "<tr><td bgcolor=""#000000""><img src=""/images/spacer.gif"" height=""4"" width=""1"" alt="""" border=""0"" /></td></tr>"
					End If
					oRs2.MoveNext
				Loop
			End If
			oRs2.NextRecordSet
			%>		
		</table>
	</td>
</tr>	
</table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing

Function DisplayMenu(strMenuName, intMenuID, intParentMenuID, byVal linkURL, spec, spec2)
	Dim strText
	strText = "<tr><td bgcolor=""" & spec & """>"
	if Not(IsNull(linkURL)) AND Len(Trim(linkURL)) > 1 Then
		linkURL = Replace(linkURL, "http://www.teamwarfare.com/", "")
		If Left(linkURL, 1) = "/" Then
			on error resume next
			linkUrl = Right(linkURL, Len(linkURL) - 1)
			if err <> 0 Then
				response.write linkUrl
				Response.end
			end if
			on error goto 0
		End If
		strText = strText & "<a href=""/" & linkURL & """>" & spec2 & Server.HTMLEncode(strMenuName & "") & "</a></td>" & VbCrLf
	Else
		strText = strText & spec2 & Server.HTMLEncode(strMenuName & "") & VbCrLf
	end if
	strText = strTExt & "</td>"
	strText = strText & "</tr>"
	DisplayMenu = strText
End Function
%>

