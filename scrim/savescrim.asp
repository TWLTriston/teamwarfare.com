<%
Option Explicit
Response.Buffer = True
%>
<!-- #include file="../include/i_funclib.asp" -->
<% 	
Dim strSaveType
strSaveType = Trim(Request.Form("SaveType"))
If Len(strSaveType) = 0 Then
	strSaveType = Trim(Request.QueryString("SaveType"))
End If
If Len(strSaveType) = 0 Then
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
	
Dim oConn, oRs, strSQL

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open 	

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If' Elo Ladder Variables
Dim intGameID, intLadderID
Dim strLadderName, strAbbr, strRules, intMinRoster, intMaxRatingDiff
Dim intIdentifierID, intActive, intLocked
Dim strMethod

If strSaveType = "EloLadder" Then
	strMethod = Request.Form("SaveMethod")
	
	intLadderID = Request.Form("EloLadderID")
	strLadderName = Request.Form("LadderName")
	intGameID = Request.Form("GameID")
	strAbbr = Request.Form("Abbreviation")
	strRules = Request.Form("Rules")
	intMinRoster = Request.Form("MinPlayer")
	intMaxRatingDiff = Request.Form("MaxRatingDiff")
	intIdentifierID = Request.Form("selIdentifierID")
	intActive = Request.Form("Active")
	intLocked = Request.Form("Locked")

	Select Case strMethod
		Case "New"
			strSQL = "SELECT EloLadderID FROM tbl_elo_ladders WHERE EloLadderName = '" & CheckString(strLadderName) & "'"
			oRs.Open strSQL, oConn
			If Not(oRs.EOF AND oRs.BOF) Then
				Response.Clear
				Response.Redirect "LadderAdd.asp?e=1"
			End If
			oRs.NextRecordSet
			strSQL = "INSERT INTO tbl_elo_ladders ( "
			strSQL = strSQL & " EloLadderName, EloRulesName, EloGameID, EloActive, EloLocked, EloAbbreviation, EloMinPlayer, EloMaxRatingDiff, EloIdentifierID"
			strSQL = strSQL & " ) VALUES ( "
			strSQL = strSQL & " '" & CheckString(strLadderName) & "', "
			strSQL = strSQL & " '" & CheckString(strRules) & "', "
			strSQL = strSQL & " '" & CheckString(intGameID) & "', "
			strSQL = strSQL & " '" & CheckString(intActive) & "', "
			strSQL = strSQL & " '" & CheckString(intLocked) & "', "
			strSQL = strSQL & " '" & CheckString(strAbbr) & "', "
			strSQL = strSQL & " '" & CheckString(intMinRoster) & "', "
			strSQL = strSQL & " '" & CheckString(intMaxRatingDiff) & "', "
			strSQL = strSQL & " '" & CheckString(intIdentifierID) & "'"
			strSQL = strSQL & ") "
			Response.Write strSQL
			oConn.Execute(strSQL)
			
			oConn.Close
			Set oConn = Nothing
			Set oRs = Nothing
			
			Response.Clear
			Response.Redirect "GeneralAdmin.asp"
		Case "Edit"
			strSQL = "UPDATE tbl_elo_ladders SET "
			strSQL = strSQL & " EloLadderName = '" & CheckString(strLadderName) & "', "
			strSQL = strSQL & " EloRulesName = '" & CheckString(strRules) & "', "
			strSQL = strSQL & " EloGameID = '" & CheckString(intGameID) & "', "
			strSQL = strSQL & " EloActive = '" & CheckString(intActive) & "', "
			strSQL = strSQL & " EloLocked = '" & CheckString(intLocked) & "', "
			strSQL = strSQL & " EloAbbreviation = '" & CheckString(strAbbr) & "', "
			strSQL = strSQL & " EloMinPlayer = '" & CheckString(intMinRoster) & "', "
			strSQL = strSQL & " EloMaxRatingDiff = '" & CheckString(intMaxRatingDiff) & "', "
			strSQL = strSQL & " EloIdentifierID = '" & CheckString(intIdentifierID) & "' "
			strSQL = strSQL & " WHERE EloLadderID = '" & CheckString(intLadderID) & "' "
			oConn.Execute(strSQL)
			oConn.Close
			Set oConn = Nothing
			Set oRs = Nothing
			
			Response.Clear
			Response.Redirect "GeneralAdmin.asp"
			
	End Select
End If
%>
 		
 		