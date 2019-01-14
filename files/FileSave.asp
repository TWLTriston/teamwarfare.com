<% Option Explicit %>
<%
Response.Buffer = True
Server.ScriptTimeout = 10000
%>
<!-- #include VIRTUAL="/include/i_funclib.asp" -->
<%

Dim strSQL, oConn, oRS
Dim path
Dim upl
Dim NewFileName

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=2"	
End If

Set upl = Server.CreateObject("SoftArtisans.FileUp")
path = "D:\TWLFiles\UPLOAD\" & Right(upl.Form("Category"), Len(upl.Form("Category")) - InStr(1, upl.Form("Category"), "|"))

upl.Path = path
		
If Not IsObject(upl.Form("FILE1")) Then		
	'this error shouldnt happen at all, unless the name of the form is changed
	oConn.Close
	set oConn = Nothing
	Set oRs = Nothing
	Set UPL = Nothing
	Response.Clear 
	Response.Redirect("/errorpage.asp?error=19")
ElseIf upl.Form("FILE1").IsEmpty Then 
	'if file is empty
	oConn.Close
	set oConn = Nothing
	Set oRs = Nothing
	Set UPL = Nothing
	Response.Clear 
	Response.Redirect("/errorpage.asp?error=17")
ElseIf upl.Form("FILE1").ContentDisposition <> "form-data" Then 
	'if bad browser
	oConn.Close
	set oConn = Nothing
	Set oRs = Nothing
	Set UPL = Nothing
	Response.Clear 
	Response.Redirect("/errorpage.asp?error=16")
Else
	If Err <> 0 Then
		'error occured in upload/save
		oConn.Close
		set oConn = Nothing
		Set oRs = Nothing
		Set UPL = Nothing
		Response.Clear 
		Response.Redirect("/errorpage.asp?error=18")
	Else
		NewFileName = Mid(upl.form("FILE1").UserFilename, InstrRev(upl.form("FILE1").UserFilename, "\") + 1)
		upl.form("FILE1").SaveAs path & "\" & NewFileName
		strSQL = "EXECUTE AddFile '" & CheckString(NewFileName) & "','" & upl.Form("PlayerID") & "','" & CheckString(upl.Form("Description")) & "', '" & CheckString(Left(upl.Form("Category"), Instr(1, upl.Form("Category"), "|") - 1)) & "'"
		oConn.Execute(strSQL)

		oConn.Close
		set oConn = Nothing
		Set oRs = Nothing
		Set UPL = Nothing
		Response.Clear 
		Response.Redirect("/files/")
	End If
End If
Response.Redirect("/files/")
%>