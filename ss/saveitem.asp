<% Option Explicit %>
<%
Response.Buffer = True
Server.ScriptTimeout = 10000
%>
<!-- #include VIRTUAL="/include/i_funclib.asp" -->
<%

Dim strSQL, oConn, oRS
Dim strFilePath, objFSO, strFileName, strDirectory, intFileID
Dim intCategoryID, strDescription
Dim strCategoryName

Set oConn = Server.CreateObject("ADODB.Connection")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()
strFilePath = "E:\TWLFiles\UPLOAD\"
If Request.QueryString("SaveType") = "DeleteFile" Then
	If Not(IsSysAdmin()) Then
		oConn.Close
		Set oConn = Nothing
		Response.Clear
		Response.Redirect "/errorpage.asp?error=2"	
	End If
	
	strFileName = Request.QueryString("filename")
	strDirectory = Request.QueryString("category")
	intFileID = Request.QueryString("fileid")
	strSQL = "EXECUTE DeleteFile '" & CheckString(intFileID) & "'"
	oConn.Execute (strSQL)
	Set objFSO = Server.CreateObject ("Scripting.FileSystemObject")
	If objFSO.FileExists(strFilePath & "\" & strDirectory & "\" & strFileName) Then
		objFSO.DeleteFile(strFilePath & "\" & strDirectory & "\" & strFileName)
	End If
	Set objFSO = Nothing
End If


If Request.Form("SaveType") = "EditFile" Then
	If Not(IsSysAdmin() OR IsAnyLadderAdmin()) Then
		oConn.Close
		Set oConn = Nothing
		Response.Clear
		Response.Redirect "/errorpage.asp?error=2"	
	End If
	
	intCategoryID = Request.Form("categoryID")
	strDescription = Request.Form("description")
	intFileID = Request.Form("fileid")
	strSQL = "UPDATE tbl_files SET FileCategoryID = '" & intCategoryID & "', Description='" & CheckString(strDescription) & "' WHERE FileID = '" & intFileID &  "'"
	oConn.Execute (strSQL)
	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/files/?game=" & intCategoryID
End If

If Request.Form("SaveType") = "AddCategory" Then
	strCategoryName = Request.Form("CategoryName")
	strDirectory = Request.Form("Directory")
	strSQL = "INSERT INTO tbl_file_category (CategoryName, Directory) VALUES ('" & CheckString(strCategoryName) & "', '" & CheckString(strDirectory) & "')"
	oConn.Execute(strSQL)
End If

oConn.Close
Set oConn = Nothing
Response.Redirect "/files/"
%>