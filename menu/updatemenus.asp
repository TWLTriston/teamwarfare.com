<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Update TWL Menus"

Dim strSQL, oConn, oRS, oRs2, oRs3, oRs4
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

if not(IsSysAdminLevel2()) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<script language="JavaScript">
<!--
 function fDeleteItem(PMenuID, MenuID) {
 	if (confirm("Are you CERTAIN you want to delete this menu and all it's children?")) {
 		window.location = "SaveMenu.asp?PMenuID=" + PMenuID + "&SaveType=DeleteMenu&MenuID=" + MenuID 
 	}
 }
//-->
</script>
<%
Call ContentStart("Update TWL Menus - " & Request.QuerySTring("Menu"))
If Request.Querystring("s") = "1" Then
	Response.Write "-----------------------<br /><b>Menus are updated.</b><br />-----------------------<br /><br />"
End If
%>
Ladders should be created as: <br />
<b>Region</b> (Blank if US) <b>Mod</b> (Blank if base) <b>Type</b> (Never Blank) <b>Size</b> (1v1 NOT duel) <br /><br />
Ex: Europe Classic Capture the Flag 12v12<br />
Ex: Team Combat 2v2<br />

<a href="AddMenuItem.asp?PMenuID=<%=Request.QuerySTring("MenuID")%>">Add Menu Item</a> | <a href="makemenus.asp">Publish Menus</a>
<br /><br />
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#000000" width="97%">
<tr>
<%
strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE MenuID = '" & Request.QuerySTring("MenuID") & "' ORDER BY SortOrder, ShowMenuName "
oRs.Open strSQL, oConn
If Not(oRS.EOF AND oRs.BOF) Then
	Do While Not(oRs.EOF)
		%>
		<td valign="top" align="center">
		<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" width="97%">
		<tr>
			<td>
			<table border="0" cellspacing="1" cellpadding="4" width="100%">
			
			<%
			Response.Write DisplayMenu(oRs.Fields("ShowMenuName").Value, oRs.Fields("MenuID").Value, oRs.Fields("ParentMenuID").Value, oRs.Fields("LinkURL").Value, "#000000", "")
			strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & oRs.Fields("MenuID").Value & " ORDER BY SortOrder, ShowMenuName "
			oRs2.Open strSQL, oConn
			If Not(oRs2.EOF AND oRs2.BOF) Then
				Do While Not(oRs2.EOF)
					Response.Write DisplayMenu(oRs2.Fields("ShowMenuName").Value, oRs2.Fields("MenuID").Value, oRs2.Fields("ParentMenuID").Value, oRs2.Fields("LinkURL").Value, bgcone, "&nbsp;&nbsp;")
					strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & oRs2.Fields("MenuID").Value & " ORDER BY SortOrder, ShowMenuName "
					oRs3.Open strSQL, oConn
					If Not(oRs3.EOF AND oRs3.BOF) Then
						Do While Not(oRs3.EOF)
							Response.Write DisplayMenu(ors3.Fields("ShowMenuName").Value, ors3.Fields("MenuID").Value, ors3.Fields("ParentMenuID").Value, ors3.Fields("LinkURL").Value, bgctwo, "&nbsp;&nbsp;&nbsp;&nbsp;")
							strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & oRs3.Fields("MenuID").Value & " ORDER BY SortOrder, ShowMenuName "
							oRs4.Open strSQL, oConn
							If Not(oRs4.EOF AND oRs4.BOF) Then
								Do While Not(oRs4.EOF)
									Response.Write DisplayMenu(ors4.Fields("ShowMenuName").Value, ors4.Fields("MenuID").Value, ors4.Fields("ParentMenuID").Value, ors4.Fields("LinkURL").Value, "#000000", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
									oRs4.MoveNext
								Loop
							End If
							oRs4.NextRecordSet
							oRs3.MoveNext
						Loop
					End If
					oRs3.NextRecordSet
					oRs2.MoveNext
				Loop
			End If
			oRs2.NextRecordSet
			%>
			</table>
			</td>
		</tr>
		</table>
		</td>
		<%
		oRs.MoveNext
	Loop
End If
oRs.NextRecordSet
%>
</tr>
</table>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing

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
	strText = strText & "<td bgcolor=""" & spec & """ align=""center""><a href=""AddMenuItem.asp?PMenuID=" & Request.QuerySTring("MenUID") & "&MenuID=" & intMenuID & """>edit</a></td>" & VbCrLf
	strText = strText & "<td bgcolor=""" & spec & """ align=""center""><a href=""javascript:fDeleteItem(" & Request.QuerySTring("MenUID") & ", " & intMenuID & ");"">delete</a></td>" & VbCrLf
	strText = strText & "</tr>"
	DisplayMenu = strText
End Function
%>