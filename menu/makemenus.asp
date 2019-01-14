<% Option Explicit %>
<%
Response.Buffer = True
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<%
Dim strPageTitle

strPageTitle = "TWL: Update TWL Menus"

Dim strSQL, oConn, oRS, oRs2, oRs3, oRs4
Dim oFS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")
Set oRS3 = Server.CreateObject("ADODB.RecordSet")
Set oRS4 = Server.CreateObject("ADODB.RecordSet")
Set oFS = Server.CreateObject("Scripting.FileSystemObject")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

if not(IsSysAdminLevel2()) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If

Dim oFile, oReadFile
Response.Write Server.MapPath("/include/menus")
Set oFile = oFS.CreateTextFile(Server.MapPath("/include/menus") & "\i_forums_public.asp", true)
' oFile.WriteLine ("	function fForumsAll() { ")
oFile.WriteLine ("<" & "%")
oFile.WriteLine ("'--------------------------------------")
oFile.WriteLine ("' This file is auto generated by the /menu/makemenus.asp ")
oFile.WriteLine ("'--------------------------------------")
oFile.WriteLine ("%" & ">")
Call fWriteMenu(2574)
oFile.Close

Set oFile = oFS.CreateTextFile(Server.MapPath("/include/menus") & "\i_forums_staff_1.asp", true)
' oFile.WriteLine ("	function fForumsAll() { ")
oFile.WriteLine ("<" & "%")
oFile.WriteLine ("'--------------------------------------")
oFile.WriteLine ("' This file is auto generated by the /menu/makemenus.asp ")
oFile.WriteLine ("'--------------------------------------")
oFile.WriteLine ("%" & ">")
Call fWriteMenu(2573) 
oFile.Close

Set oFile = oFS.CreateTextFile(Server.MapPath("/include/menus") & "\i_forums_staff_2.asp", true)
' oFile.WriteLine ("	function fForumsAll() { ")
oFile.WriteLine ("<" & "%")
oFile.WriteLine ("'--------------------------------------")
oFile.WriteLine ("' This file is auto generated by the /menu/makemenus.asp ")
oFile.WriteLine ("'--------------------------------------")
oFile.WriteLine ("%" & ">")
Call fWriteMenu(2572) 
oFile.Close

oConn.Close
Set oConn = Nothing
Set oRS = Nothing

Response.Clear
Response.Redirect "default.asp?s=1"

Function fWriteMenuLine(byVal strMenuName, byVal linkURL, byVal bHasChild)
	Dim strText
	Response.Write strMenuName & " - [" & linkURL & "] - " & bHasChild & "<br />"
	if Not(IsNull(linkURL)) AND Len(Trim(linkURL)) > 0 Then
		linkURL = Replace(linkURL, "http://www.teamwarfare.com/", "")
		If Left(linkURL, 1) = "/" Then
			linkURL = Right(linkURL, Len(linkURL) - 1)
		End If
	end if
	If Len(linkURL) = 0 Then
		strText = "<a href=""#"""
	Else 
		strText = "<a href=""/" & linkURL & """"
	End If
	If (bHasChild) Then
		strText = strText & " class=""haschild"""
	End If
	strText = strText & ">" & Server.HTMLEncode(strMenuName) & "</a>"
	fWriteMenuLine = strText
End Function

Function fWriteMenu(iParentMenuID) 
	strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & iParentMenuID & " ORDER BY SortOrder, ShowMenuName "
	oRs2.Open strSQL, oConn
	If Not(oRs2.EOF AND oRs2.BOF) Then
		Do While Not(oRs2.EOF)
			oFile.Write ("<li>")
	
			strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & oRs2.Fields("MenuID").Value & " ORDER BY SortOrder, ShowMenuName "
			oRs3.Open strSQL, oConn
			If Not(oRs3.EOF AND oRs3.BOF) Then
				oFile.Write fWriteMenuLine(oRs2.Fields("ShowMenuName").Value, oRs2.Fields("LinkURL").Value, True)
				oFile.Write "<ul>"
				Do While Not(oRs3.EOF)
					oFile.Write ("<li>")
					
					strSQL = "SELECT MenuID, ParentMenuID, ShowMenuName, LinkURL, SortOrder FROM tbl_menus WHERE ParentMenuID = " & oRs3.Fields("MenuID").Value & " ORDER BY SortOrder, ShowMenuName "
					oRs4.Open strSQL, oConn
					If Not(oRs4.EOF AND oRs4.BOF) Then
						oFile.Write fWriteMenuLine(oRs3.Fields("ShowMenuName").Value, oRs3.Fields("LinkURL").Value, True)
						oFile.Write "<ul>"
						Do While Not(oRs4.EOF)
							oFile.Write 	"<li>"
							oFile.Write fWriteMenuLine(oRs4.Fields("ShowMenuName").Value, oRs4.Fields("LinkURL").Value, False)
							oFile.Write "</li>" 
							oRs4.MoveNext
						Loop
						oFile.Write "</ul>"
					Else 
						oFile.Write fWriteMenuLine(oRs3.Fields("ShowMenuName").Value, oRs3.Fields("LinkURL").Value, False)
					End If
					oRs4.NextRecordSet
					oFile.Write ("</li>")
					oRs3.MoveNext
				Loop
	
				oFile.Write "</ul>"
			Else 
				oFile.Write fWriteMenuLine(oRs2.Fields("ShowMenuName").Value, oRs2.Fields("LinkURL").Value, False)
			End If
			oRs3.NextRecordSet
	
			oFile.Write ("</li>")
			oRs2.MoveNext
		Loop
	End If
	oRs2.NextRecordSet
End Function
%>