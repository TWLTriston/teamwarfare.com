<%
 	response.buffer=true

	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	Set oRS2 = Server.CreateObject("ADODB.RecordSet")
	
	oConn.ConnectionString = Application("ConnectStr")
	oConn.Open 	
%>
<!-- #include file="../include/i_funclib.asp" -->
<% 	
Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(IsSysAdminLevel2()) Then
	Response.Redirect "/default.asp"
End If

 	
 	if request.QueryString("saveType")="DeleteMenu" then
 		strSQL = "DELETE FROM tbl_menus WHERE MenuID = '" & CheckString(Request.QueryString("MenuID")) & "'"
 		oConn.Execute(strSQL)
 		Response.Clear
 		Response.Redirect "UpdateMenus.asp?MenuID=" & Request.QueryString("PMenuID")
 	End If

 	if request.form("saveType")="AddMenu" then
 		strShowMenuName = request.form("ShowMenuName")
 		intParentMenuID = request.form("ParentMenuID")
 		intMenuID = Request.Form("MenuID")
 		strLinkURL = Request.Form("LinkURL")
 		If request.Form("SaveMode") = "Edit" Then
 			'' Edit!
 			strSQL = "UPDATE tbl_menus SET "
 			strSQL = strSQL & " ShowMenuName = '" & CheckString(strShowMenuName) & "', "
 			strSQL = strSQL & " ParentMenuID = " & CheckString(intParentMenuID) & ", "
 			strSQL = strSQL & " LinkURL = '" & CheckString(strLinkURL) & "' "
 			strSQL = strSQL & " WHERE MenuID ='" & intMenuID & "'"
 			oConn.Execute(strSQL)
 			oConn.CLose
 			Set oConn = nothing
 			Response.Clear
 			Response.Redirect "UpdateMenus.asp?MenuID=" & Request.Form("PMenuID")
 		Else
	 		strSQL="insert into tbl_menus (ShowMenuName, ParentMenuID, LinkURL ) VALUES ("
	 		strSQL = strSQL & "'" & CheckString(strShowMenuName) & "',"
	 		strSQL = strSQL & "'" & CheckString(intParentMenuID) & "',"
	 		strSQL = strSQL & "'" & CheckString(strLinkURL) & "') "
	 		oConn.execute strSQL
	 		response.clear
	 		response.redirect "UpdateMenus.asp?MenuID=" & Request.Form("PMenuID")
	 	End If
 	end if

%>
 		
 		