<!-- #include virtual="/include/xml.asp" -->
<%
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

oConn.ConnectionString = Application("ConnectStr")
oConn.Open
Dim description
Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf
%>
<rss version="2.0">
<channel>
<title>TeamWarfare: News</title>
<link>http://www.teamwarfare.com/</link>
<description>Team Warfare League: Community-based gaming</description>
<language>en-us</language>
<% 
strSQL = "SELECT TOP 10 NewsID, NewsHeadline, NewsContent FROM tbl_news WHERE NewsType = 0 ORDER BY NewsID DESC "
oRs.Open strSQL, oConn
If Not(oRs.BOF AND oRs.EOF) Then
	Do While Not(oRs.EOF)
		Response.Write vbTab & vbTab & "<item>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<title>" & Server.HTMLEncode(oRs.Fields("NewsHeadline").Value & "") & "</title>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<guid>http://www.teamwarfare.com/#" & oRs.Fields("NewsID").Value & "</guid>" & vbCrLf
		
		description = Server.HTMLEncode(Replace(oRs.Fields("NewsContent").Value, vbCrLf, " ") & "")
		description = XMLEncode(description)
		
		Response.Write vbTab & vbTab & vbTab & "<description>" & description & "</description>" & vbCrLf
		Response.Write vbTab & vbTab & "</item>" & vbCrLf
		oRs.MoveNext
	Loop
End If

oRS.NextRecordSet

strSQL = "SELECT GameID, GameName FROM tbl_Games WHERE GameID > 0 ORDER BY GameName "
oRs.Open strSQL, oConn
If Not(oRs.BOF AND oRs.EOF) Then
	Do While Not(oRs.EOF)
		strSQL = "SELECT TOP 5 NewsID, NewsHeadline, NewsDate, NewsAuthor, NewsContent FROM tbl_news WHERE NewsType = '" & oRs.Fields("GameID").Value & "' ORDER BY NewsID DESC "
		oRs2.Open strSQL, oConn
		If Not(oRs2.BOF AND oRs2.EOF) Then
			Do While Not(oRs2.EOF)
				Response.Write vbTab & vbTab & "<item>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<title>" & Server.HTMLEncode(oRS.Fields("GameName").Value & "") & ":" & Server.HTMLEncode(oRs2.Fields("NewsHeadline").Value & "") & "</title>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<guid>http://www.teamwarfare.com/#" & oRs2.Fields("NewsID").Value & "</guid>" & vbCrLf
				
		description = Server.HTMLEncode(Replace(oRs2.Fields("NewsContent").Value, vbCrLf, " ") & "")
		description = XMLEncode(description)
		
		Response.Write vbTab & vbTab & vbTab & "<description>" & description & "</description>" & vbCrLf
				Response.Write vbTab & vbTab & "</item>" & vbCrLf
				oRs2.MoveNext
			Loop
		End If
		oRs2.NextRecordSet
		
		oRs.MoveNext
	Loop
End If
oRs.NextRecordSet

oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing
%>
</channel>
</rss>