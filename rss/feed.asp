<!-- #include virtual="/include/xml.asp" -->
<%
Function CheckString(byVal strData)
	CheckString = strData
	CheckString = Replace(strData, "'", "''")
End Function

Const adParamInput = &H0001
Const adCmdStoredProc = &H0004
Const adVarChar = 200
Const adParamReturnValue = &H0004
Const adParamOutput = &H0002
Const adInteger = 3
Const adLongVarChar = 201

Dim oConn, strSQL, oRs, oRs2
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

Dim strGameName, strTitle, intGameID, intNewsType
strGameName = Request.QueryString("Game")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

intGameID = -1
strSQL = "SELECT GameID, GameName FROM tbl_Games WHERE GameName = '" & CheckString(strGameName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.BOF AND oRs.EOF) Then
	Do While Not(oRs.EOF)
		intGameID = oRs.Fields("GameID").Value
		strGameName = oRs.Fields("GameName").Value
		oRs.MoveNext
	Loop
End If
oRs.NextRecordSet

If intGameID = -1 Then
	strTitle = "TeamWarfare: News"
	intNewsType = 0
Else
	strTitle = "TeamWarfare: " & Server.HTMLEncode(strGameName) & " News"
	intNewsType = intGameID
End If

Dim description
Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf
%>
<rss version="2.0">
<channel>
<title><%=strTitle%></title>
<link>http://www.teamwarfare.com/</link>
<description>TeamWarfare League: Community based gaming</description>
<language>en-us</language>
<% 
strSQL = "SELECT TOP 10 NewsID, NewsDate, NewsHeadline, NewsAuthor, NewsContent FROM tbl_news WHERE NewsType = '" & CheckString(intNewsType) & "' ORDER BY NewsID DESC "
oRs.Open strSQL, oConn
If Not(oRs.BOF AND oRs.EOF) Then
	Do While Not(oRs.EOF)
		Response.Write vbTab & vbTab & "<item>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<title>" & Server.HTMLEncode(oRs.Fields("NewsHeadline").Value & "") & "</title>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<link>http://www.teamwarfare.com/#news" & oRs.Fields("NewsID").Value & "</link>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<pubDate>" & Server.HTMLEncode(oRs.Fields("NewsDate").Value & "") & "</pubDate>" & vbCrLf
		
		description = Replace(oRs.Fields("NewsContent").Value, vbCrLf, "<br />") & ""
		description = XMLEncode(description)
		
		Response.Write vbTab & vbTab & vbTab & "<description>" & description & "</description>" & vbCrLf
		Response.Write vbTab & vbTab & "</item>" & vbCrLf
		oRs.MoveNext
	Loop
End If

oRS.NextRecordSet

oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing
%>
</channel>
</rss>