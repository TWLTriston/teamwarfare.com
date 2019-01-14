<% Option Explicit %>
<%
Response.Buffer = True
Server.ScriptTimeout = 10000
Dim strPageTitle

strPageTitle = "TWL: Save Demo File"

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

Dim fname_chars, fname_length, c, rnum, file_dir, fso		
Dim publicFlag, AdminFlag
Dim Path, UPL, Msg , fname

%>
<!-- #include virtual="/include/i_funclib.asp" -->
<%
If Not(Session("LoggedIn")) then 
	oConn.Close
	set oConn = Nothing
	Set oRs = Nothing
	Response.Clear 
	Response.Redirect("/errorpage.asp?error=2")
End If
path = Application("DemoUploadDir")

Set upl = Server.CreateObject("SoftArtisans.FileUp")

If upl.Form("Submit") = " Add Demo " Then
	' check for blank entries
	If upl.Form("map") = "" Then Msg = "<li>Map(s) Played"
	If upl.Form("PositionPlayed") = "" Then Msg = Msg & "<li>Position Played"
	
	If Msg = "" Then
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
		Elseif lCase(Right(upl.Form("FILE1").UserFilename, 4)) <> ".zip" and lCase(Right(upl.Form("FILE1").UserFilename, 4)) <> ".rec" Then
			'if not zip file
			oConn.Close
			set oConn = Nothing
			Set oRs = Nothing
			Set UPL = Nothing
			Response.Clear 
			Response.Redirect("/errorpage.asp?error=21")
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
				'successful, add stuff to demos table
				'generate file name
				fname_chars = "abcdefghijklmnopqrstuvxyzABCDEFGHIJKLMNOPQRSTUVXYZ0123456789" 'chars in pw
				fname_length = 12   'length of filename
				
				Randomize
				For c = 1 To fname_length
					rnum = Int((len(fname_chars) * Rnd) + 1)
					fname = fname & Mid(fname_chars, rnum, 1 )
				Next				
				
				fname = fname & lCase(Right(upl.Form("FILE1").UserFilename, 4))
				file_dir = Application("DemoUploadDir") & upl.Form("MatchID")
				
				' see if <MatchID> directory already exists, if not.. create it
				set fso = Server.CreateObject("Scripting.FileSystemObject")
				If Not fso.FolderExists(file_dir) then
					fso.CreateFolder(file_dir)
				End If
				Set fso = Nothing
				upl.Form("FILE1").SaveAs file_dir & "\" & fname
				
				PublicFlag = 0
				AdminFlag = 0
				
				If upl.Form("public") = "on" Then PublicFlag = 1
				If upl.Form("admin") = "on" Then AdminFlag = 1
				
				strSQL = "INSERT INTO tbl_Demos (FileName, FilePath, PlayerID, HistoryID,"
				strSQL = strSQL & "TLLinkID, FileSize, upload_dtim, MapPlayed, PositionPlayed,"
				strSQL = strSQL & "DownloadCount, PublicFlag, AdminFlag, IPAddress, Comments) VALUES("
				strSQL = strSQL & "'" & CheckString(fname) & "',"
				strSQL = strSQL & "'" & CheckString(file_dir) & "',"
				strSQL = strSQL & upl.Form("PlayerID") & ","
				strSQL = strSQL & upl.Form("HistoryID") & ","
				strSQL = strSQL & upl.Form("TLLinkID") & ","
				strSQL = strSQL & upl.Form("FILE1").TotalBytes & ","
				strSQL = strSQL & "'" & Now & "',"
				strSQL = strSQL & "'" & CheckString(upl.Form("map")) & "',"
				strSQL = strSQL & "'" & CheckString(upl.Form("PositionPlayed")) & "',"
				strSQL = strSQL & "0,"
				strSQL = strSQL & PublicFlag & ","
				strSQL = strSQL & AdminFlag & ","
				strSQL = strSQL & "'" & upl.Form("IPAddress") & "',"
				strSQL = strSQL & "'" & CheckString(upl.Form("Comments")) & "')"
				oConn.Execute(strSQL)

				oConn.Close
				set oConn = Nothing
				Set oRs = Nothing
				Set UPL = Nothing
				Response.Clear 
				Response.Redirect("default.asp")
			End If
		End If
	
	End If
Else
	oConn.Close
	set oConn = Nothing
	Set oRs = Nothing
	Set UPL = Nothing
	Response.Clear 
	Response.Redirect("/demos/addDemo.asp")
End If
Response.redirect "default.asp"
%>