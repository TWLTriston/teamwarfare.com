<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add Menu Item"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

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

Dim intMenuID, intParentMenuID, strShowMenuName, strLinkURL
Dim strSaveMode
strSaveMode = "Add"
intMenuID = Request.QueryString("MenuID")
If Len(intMenuID) > 0 AND IsNumeric(intMenuID) Then
	strSQL = "SELECT * FROM tbl_menus WHERE MenuID = '" & intMenuID & "'"
	oRs.OPen strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		intParentMenuID = oRs.Fields("ParentMenuID").Value
		strShowMenuName = oRs.Fields("ShowMenuName").Value
		strLinkURL = oRs.Fields("LinkURL").Value
	End If
	oRs.NextRecordSet
	strSaveMode = "Edit"
End IF
		
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("")
%>
<form name="frmMenus" id="frmMenus" action="SaveMenu.asp" method="post">
<input type="hidden" name="SaveType" id="SaveType" value="AddMenu" />
<input type="hidden" name="SaveMode" id="SaveMode" value="<%=strSaveMode%>" />
<input type="hidden" name="PMenuID" id="PMenuID" value="<%=Request.QuerySTring("PMenUID")%>" />
<input type="hidden" name="MenuID" id="MenuID" value="<%=intMenuID%>" />
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" width="75%">
<tr>
	<td>
<table border="0" cellspacing="1" cellpadding="4" width="100%">
<tr>
	<td align="right" bgcolor="<%=bgcone%>">Parent Menu Item:</td>
	<td bgcolor="<%=bgcone%>"><select name="ParentMenuID" id="ParentMenuID" size="12">
		<option value="0" selected="selected">none</option>
			<%
			strSQL = "SELECT m.MenuID, m.ShowMenuName, m2.ShowMenuName AS ParentMenuName, m3.ShowMenuName AS ParentParentMenuName "
			strSQL = strSQL & " FROM tbl_menus m "
			strSQL = strSQL & " LEFT OUTER JOIN tbl_menus m2 ON m.ParentMenuID = m2.MenuID "
			strSQL = strSQL & " LEFT OUTER JOIN tbl_menus m3 ON m2.ParentMenuID = m3.MenuID "
			strSQL = strSQL & " WHERE m.MenuID = '" & Request.QueryString("PMenuID") & "' OR m3.ParentMenuID = '" & Request.QueryString("PMenuID") & "' OR m2.ParentMenuID = '" & Request.QueryString("PMenuID") & "' OR m.ParentMenuID = '" & Request.QueryString("PMenuID") & "'"
			strSQL = strSQL & " ORDER BY m3.ShowMenuName, m2.ShowMenuName, m.ShowMenuName "
			oRs.Open strSQL, OConn
			If Not (oRs.EOF ANd oRs.BOF) Then
				Do While Not(oRs.EOF)
					Response.Write "<option value=""" & oRs.FieldS("MenuID").Value & """ "
					If oRs.FieldS("MenuID").Value  = intParentMenuID THen
						Response.Write " SELECTED "
					End If
					Response.Write ">" & oRs.Fields("ParentParentMenuName").Value & " - " & oRs.FIelds("ParentMenuName").Value & " - " & oRS.FIelds("ShowMenuName").Value
					Response.write "</option>" & vbCrLF
					
					oRs.MoveNext
				Loop
			End If
			%>
		</select>
	</td>
</tr>
<tr>
	<td align="right" bgcolor="<%=bgctwo%>">Name:</td>
	<td bgcolor="<%=bgctwo%>"><input type="text" id="ShowMenuName" size="60" name="ShowMenuName" value="<%=Server.HTMLENcode(strShowMenuName & "")%>" /></td>
</tr>
<tr>
	<td align="right" bgcolor="<%=bgctwo%>">Link (URLEncoded):</td>
	<td bgcolor="<%=bgcone%>">Style: viewladder.asp?ladder=Tribes+2+CTF<br /><input type="text" id="LinkURL" size="60" name="LinkURL" value="<%=Server.HTMLENcode(strLinkURL & "")%>" />
	<br />
	(dont include leading "/" or leading "http://www.teamwarfare.com")
	</td>
</tr>
<tr>
	<td align="center" colspan="2" bgcolor="#000000"><input type="submit" value="Save it" /></td>
</tr></table>
</td>
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
%>