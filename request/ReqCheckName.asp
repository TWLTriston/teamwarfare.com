<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = ""

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<%
strSQL="SELECT count(playerHandle) cnt from tbl_players where playerHandle='" & CheckString(Request("txtNewName")) & "'"
oRS.Open strSQL, oConn
if oRS.Fields("cnt") > 0 then
	Response.clear
	Response.Redirect "ReqNameChange.asp?e=1&player=" & Server.URLEncode(Session("uName"))
end if
oRS.Close
strSQL = "INSERT INTO tbl_PlayerNameChange(PlayerID, OldName, NewName, RequestDate, Approved, ApprovedByID, ApprovedDate, Notes) VALUES "
strSQL = strSQL & "('" & Request("hdnPlayerID") & "',"
strSQL = strSQL & "'" & CheckString(Session("uName")) & "',"
strSQL = strSQL & "'" & CheckString(Request("txtNewName")) & "',"
strSQL = strSQL & "GetDate(),0,0,'','')"  
oConn.Execute strSQL
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Response.clear 
Response.Redirect "ReqThx.asp"
Response.End 
%>	
