<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Accept Match"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<%
Dim Comment
If Request.Form("Submit") = " Reply! " Then
	If Request.Form("comment") = "" Then Response.Redirect("viewDemo.asp?DemoID=" & Request.Form("DemoID") & "&err=1")
	
	If Request.Form("sig") = "on" Then
		strSQL = "SELECT * FROM tbl_Players WHERE PlayerID=" & Request.Form("PlayerID")
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRS.BOF) Then
			comment = Checkstring(Request.Form("comment")) & "<br><br>" & oRS("PlayerSignature")
		End If
		oRS.Close
	Else
		comment = Checkstring(Request.Form("comment"))
	End If
	
	strSQL = "INSERT INTO lnk_comment_demo (DemoID,PlayerID,Comment,CommentTime,IPAddress) VALUES("
	strSQL = strSQL & Request.Form("DemoID") & ","
	strSQL = strSQL & Request.Form("PlayerID") & ","
	strSQL = strSQL & "'" & Replace(comment, "'", "''") & "',"
	strSQL = strSQL & "'" & now & "',"
	strSQL = strSQL & "'" & Request.ServerVariables("REMOTE_ADDR") & "')"
	oConn.Execute(strSQL)
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Redirect "viewDemo.asp?DemoID=" & Request.Form("DemoID")
Else
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Redirect("default.asp")
End If
%>
