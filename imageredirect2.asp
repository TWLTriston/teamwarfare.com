<% Option Explicit %>
<!-- #Include virtual="/include/i_funclib.asp" -->
<%
Dim strImageLocation
strImageLocation = Request.QueryString("l")

If Len(strImageLocation) > 0 Then
	Dim strSQL, oConn, oRs
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionString = Application("ConnectStr")
	oConn.Open
	Dim oFs
	Set oFs = Server.CreateObject("Scripting.FileSystemObject")
	If oFs.FileExists(Server.MapPath("/") & strImageLocation) Then
		strSQL = "INSERT INTO tbl_track_image (TrackImageURL, TrackImageDateTime, TrackImagePlayer, TrackImageIP) VALUES ("
		strSQL = strSQL & "'" & CheckString(strImageLocation) & "', "
		strSQL = strSQL & " GetDate(), "
		strSQL = strSQL & "'" & CheckString(Session("uName")) & "', "
		strSQL = strSQL & "'" & CheckString(Request.ServerVariables("REMOTE_ADDR")) & "') "
		oConn.Execute(strSQL)
		On Error Resume Next
		Server.Transfer strImageLocation
		On Error Goto 0
		oConn.Close
		Set oConn = Nothing
	End If
	Set oFs = Nothing
Else
	Response.Write strImageLocation
End If

Response.End%>