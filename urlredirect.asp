<% Option Explicit %>
<!-- #Include virtual="/include/i_funclib.asp" -->
<%
Dim strURL
strURL = Request.QueryString("l")

If Len(strURL) > 0 Then
	Dim strSQL, oConn, oRs
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.ConnectionString = Application("ConnectStr")
	oConn.Open
	Dim oFs
		strSQL = "INSERT INTO tbl_track_url (TrackURLName, TrackURLDateTime, TrackURLPlayer, TrackURLIP) VALUES ("
		strSQL = strSQL & "'" & CheckString(strURL) & "', "
		strSQL = strSQL & " GetDate(), "
		strSQL = strSQL & "'" & CheckString(Session("uName")) & "', "
		strSQL = strSQL & "'" & CheckString(Request.ServerVariables("REMOTE_ADDR")) & "') "
		oConn.Execute(strSQL)
		oConn.Close
		Set oConn = Nothing

	Response.Redirect strURL
Else
	Response.Write strURL
End If

Response.End%>