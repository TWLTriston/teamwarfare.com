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
if IsSysAdmin then 
	strSQL = "UPDATE tbl_PlayerNameChange SET "
	strSQL = strSQL & "Approved=" & Request("status") & ", ApprovedDate=GetDate() where RequestID = " & Request("req")
			'Response.Write strSQL & "<br />"
	oConn.Execute strSQL
	if Request("status") = 1 then
		strSQL = "SELECT PlayerID, NewName from tbl_PlayerNameChange where requestID=" & Request("req")
			'Response.Write strSQL  & "<br />"
		oRS.Open strSQL, oConn
		dim intPlayerID, strNewName
		if not (oRS.EOF and oRS.BOF) then
			intPlayerID = oRS.Fields("PlayerID")
			strNewName = oRS.Fields("NewName")
			strSQL = "Update tbl_players set playerhandle = '" & CheckString(strNewName) & "' where playerID=" & intPlayerID
			'Response.Write strSQL  & "<br />"
			oConn.Execute strSQL
		end if
	end if
end if
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Response.clear 
Response.Redirect "AdmNames.asp"
Response.End 
%>	
