<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Save Flags"

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
Dim AdminFlag, PublicFlag
If Not(Session("LoggedIn")) then 
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Redirect("/errorpage.asp?error=2")
End If

If Request.Form("submitFlags") = "  Update  " Then
	AdminFlag = Request.Form("AdminFlag")
	PublicFlag = Request.Form("PublicFlag")
	
	If AdminFlag = "on" Then 
		AdminFlag = 1
	Else
		AdminFlag = 0
	End If
	
	If PublicFlag = "on" Then
		PublicFlag = 1
	Else
		PublicFlag = 0
	End If
	
	strSQL = "UPDATE tbl_Demos SET "
	strSQL = strSQL & "AdminFlag=" & AdminFlag & ","
	strSQL = strSQL & "PublicFlag=" & PublicFlag
	strSQl = strSQL & " WHERE DemoID=" & Request.Form("DemoID")
	
	oConn.Execute(strSQL)
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Redirect("viewDemo.asp?DemoID=" & Request.Form("DemoID"))
Else
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Clear
	Response.Redirect("default.asp")
End If
%>