<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: SaveItem"

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
<!-- #INCLUDE virtual="/include/i_funclib.asp" -->
<%
if not(Session("LoggedIn")) Then
	' Require login to perform action.
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=2"
End If

'Response.Write "<font color=#ffffff>Querystring: " & Request.QueryString
'Response.Write "<br>Form data: " & Request.Form & "</font></br>"

Dim intLadderID, SaveType, strBuiltURL
Dim intMatchID, strLadderName, strTeamName
Dim formItem

intLadderID = Request("LadderID")
SaveType = Request("SaveType")
strBuiltURL = "ladderoptions.asp?ladderid=" & intLadderID

Dim strOptionName, chrSide, chrSelectedBy, intMapNumber, intOptionID
Dim intCounter
Select Case SaveType
	Case "AddOption"
		strOptionName	= Request("OptionName")
		chrSide			= Request("SideChoice")
		chrSelectedBy	= Request("SelectedBy")
		intMapNumber	= Request("MapNumber")
		intOptionID		= Request("OptionID")
		If Request("Method") = "New" Then
			strSQL = "INSERT INTO tbl_ladder_options (OptionName, SideChoice, SelectedBy, MapNumber, LadderID, Active) VALUES ("
			strSQL = strSQL & "'" & CheckString(strOptionName) & "','" & chrSide & "','" & chrSelectedBy & "','" & intMapNumber & "','" & intLadderID & "', 'Y')"
			oConn.Execute (strSQL)
		ElseIf Request("Method") = "Edit" Then
			strSQL = "UPDATE tbl_ladder_options SET "
			strSQL = strSQL & " OptionName = '" & CheckString(strOptionName) & "', "
			strSQL = strSQL & " SideChoice = '" & CheckString(chrSide) & "', "
			strSQL = strSQL & " SelectedBy = '" & CheckString(chrSelectedBy) & "', "
			strSQL = strSQL & " MapNumber = '" & CheckString(intMapNumber) & "' "
			strSQL = strSQL & " WHERE OptionID = '" & intOptionID & "'"
			oConn.Execute (strSQL)
		End If	
	Case "OptionValue"
		strSQL = ""
		intOptionID		= Request("OptionID")
		For intCounter = 1 To Request.Form("OptionValueID").Count
			If Request.Form("OptionValueID")(intCounter) = "NEW" And Len(Trim(Request.Form("ValueName")(intCounter))) > 0 Then
				strSQL = strSQL & "INSERT INTO tbl_ladder_option_value (ValueName, OptionID, Active) VALUES ('" & CheckString(Request.Form("ValueName")(intCounter)) & "', '" & intOptionID & "', '" & Request.Form("Active")(intCounter) & "');"
			ElseIf Len(Trim(Request.Form("ValueName")(intCounter))) > 0 Then
				strSQL = strSQL & "UPDATE tbl_ladder_option_value SET "
				strSQL = strSQL & " ValueName = '" & CheckString(Request.Form("ValueName")(intCounter)) & "', "
				strSQL = strSQL & " Active = '" & Request.Form("Active")(intCounter) & "' "
				strSQL = strSQL & " WHERE OptionValueID = '" & Request.Form("OptionValueID")(intCounter) & "'; "
			
			End If
		Next
		oConn.Execute (strSQL)
		strBuiltURL = "addoption.asp?ladderid=" & intLadderID & "&optionid=" & intOptionID & "&isEdit=true"
	Case "deleteoption"
		intOptionID		= Request("OptionID")
		strSQL = "UPDATE tbl_ladder_options SET Active='N' WHERE OptionID='" & intOptionID & "'"
		oConn.Execute  (strSQL)
	Case "restoreoption"
		intOptionID		= Request("OptionID")
		strSQL = "UPDATE tbl_ladder_options SET Active='Y' WHERE OptionID='" & intOptionID & "'"
		oConn.Execute  (strSQL)
	Case "SaveMapOptions"
		strLadderName	= Request("LadderName")
		strTeamName		= Request("TeamName")
		strBuiltURL		= "/teamladderadmin.asp?team=" & Server.URLEncode(strTeamName) & "&ladder=" & strLadderName
		intMatchID		= Request("MatchID")
		strSQL = ""
		For Each formItem In Request.Form
			If Left(formItem, 3) = "MO_" Then
				' This is a match option
				strSQL = strSQL & "EXEC usp_AddMatchOption '" & intMatchID & "', '" &  Right(formItem, Len(formItem) - 3)  & "', '" &  Request.Form(formItem) & "';"
			End If
		Next
		If Len(strSQL) > 0 Then
			oConn.Execute strSQL
		End If
End Select
'-----------------------------------------------
' Housekeeping
'-----------------------------------------------
oConn.Close 
set ors = nothing
set oConn = nothing	
set ors2 = nothing	
Response.Clear
Response.Redirect strBuiltURL
Response.End
%>	
