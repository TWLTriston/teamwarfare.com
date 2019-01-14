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

Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf
Response.Write "<news>" & vbCrLf
Response.Write vbTab & "<announcements>" & vbCrLf
strSQL = "SELECT TOP 10 NewsHeadline, NewsDate, NewsAuthor, NewsContent FROM tbl_news WHERE NewsType = 0 ORDER BY NewsID DESC "
oRs.Open strSQL, oConn
If Not(oRs.BOF AND oRs.EOF) Then
	Do While Not(oRs.EOF)
		Response.Write vbTab & vbTab & "<announcement>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<headline>" & XMLEncode(Server.HTMLEncode(oRs.Fields("NewsHeadline").Value & "")) & "</headline>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<newsdatetime>" & XMLEncode(Server.HTMLEncode(oRs.Fields("NewsDate").Value & "")) & "</newsdatetime>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<author>" & XMLEncode(Server.HTMLEncode(oRs.Fields("NewsAuthor").Value & "")) & "</author>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<content>" & XMLEncode(Server.HTMLEncode(Replace(oRs.Fields("NewsContent").Value, vbCrLf, " ") & "")) & "</content>" & vbCrLf
		Response.Write vbTab & vbTab & "</announcement>" & vbCrLf
		oRs.MoveNext
	Loop
End If
oRS.NextRecordSet
Response.Write vbTab & "</announcements>" & vbCrLf

strSQL = "SELECT GameID, GameName FROM tbl_Games WHERE GameID > 0 ORDER BY GameName "
oRs.Open strSQL, oConn
If Not(oRs.BOF AND oRs.EOF) Then
	Do While Not(oRs.EOF)
		strSQL = "SELECT TOP 5 NewsHeadline, NewsDate, NewsAuthor, NewsContent FROM tbl_news WHERE NewsType = '" & oRs.Fields("GameID").Value & "' ORDER BY NewsID DESC "
		oRs2.Open strSQL, oConn
		If Not(oRs2.BOF AND oRs2.EOF) Then
			Response.Write vbTab & "<gamespecificnews game=""" & XMLEncode(Server.HTMLEncode(oRS.Fields("GameName").Value & "")) & """>" & vbCrLf
			Do While Not(oRs2.EOF)
				Response.Write vbTab & vbTab & "<gamenews>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<headline>" & XMLEncode(Server.HTMLEncode(oRs2.Fields("NewsHeadline").Value & "")) & "</headline>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<newsdatetime>" & XMLEncode(Server.HTMLEncode(oRs2.Fields("NewsDate").Value & "")) & "</newsdatetime>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<author>" & XMLEncode(Server.HTMLEncode(oRs2.Fields("NewsAuthor").Value & "")) & "</author>" & vbCrLf
				Response.Write vbTab & vbTab & vbTab & "<content>" & XMLEncode(Server.HTMLEncode(Replace(oRs2.Fields("NewsContent").Value, vbCrLf, " ") & ""))	 & "</content>" & vbCrLf
				Response.Write vbTab & vbTab & "</gamenews>" & vbCrLf
				oRs2.MoveNext
			Loop
			Response.Write vbTab & "</gamespecificnews>" & vbCrLf
		End If
		oRs2.NextRecordSet
		
		oRs.MoveNext
	Loop
End If
oRs.NextRecordSet
Response.Write "</news>"

oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing
%>