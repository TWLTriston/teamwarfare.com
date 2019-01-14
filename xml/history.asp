<!-- #include virtual="/include/xml.asp" -->
<%
Const adParamInput = &H0001
Const adCmdStoredProc = &H0004
Const adVarChar = 200
Const adParamReturnValue = &H0004
Const adParamOutput = &H0002
Const adInteger = 3
Const adLongVarChar = 201

Dim LadderName
LadderName = Trim(Request.QueryString("LadderName"))
If Len(LadderName) = 0 Then
	Response.Write "You must specify a laddername in the querystring."
	Response.End 
End If

dim conn
set conn = server.CreateObject("adodb.connection")
conn.ConnectionString = Application("ConnectStr")
conn.Open

dim cmd, rs,  rsXML
set rs = server.CreateObject("adodb.recordset")
set cmd = server.CreateObject("adodb.command")
set cmd.ActiveConnection = conn
	cmd.CommandText = "usp_get_RecentLadderHistory"
	cmd.CommandType = adcmdstoredproc
	cmd.Parameters.Append cmd.CreateParameter("LadderName", adVarChar, adParamInput, 50, LadderName)
	set rs = cmd.Execute 

if not rs.bof or not rs.eof then
	chk = true
	rsXML = rs.getrows
end if

Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf

Dim strString

	strString = strString & "<MatchHistory>"
		if chk = true then 
			For b = 0 to ubound(rsXML,2)
				strString = strString & "<LadderMatchHistory>"
				strString = strString & "<LadderName>" & XMLEncode(rsXML(0,b)) & "</LadderName>"
				strString = strString & "<Rank>" & XMLEncode(rsXML(1,b)) & "</Rank>"
				strString = strString & "<WinnerName>" & XMLEncode(rsXML(2,b)) & "</WinnerName>"
				strString = strString & "<LoserName>" & XMLEncode(rsXML(3,b)) & "</LoserName>"
				strString = strString & "<MatchDate>" & XMLEncode(rsXML(4,b)) & "</MatchDate>"
				strString = strString & "</LadderMatchHistory>"
			Next
		end if
	strString = strString & "</MatchHistory>"
Response.Write strString

rs.close
set rs = nothing
set cmd = nothing
conn.Close 
set conn = nothing
%>