<%
Option Explicit
%>
<!-- #include virtual="/include/xml.asp" -->
<%
Dim oRS, strSQL, oConn

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Set oRS = Server.CreateObject("ADODB.RecordSet")

Dim blnFirstTime, strThisGroup
strThisGroup = -1

Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf
	Response.Write "<TeamWarfareStaff>" & vbCrLf
	blnFirstTime = True
	strSQL = "SELECT s.*, 'Group' = sg.Description FROM tbl_staff s, tbl_staff_group sg WHERE sg.staffgroupid = s.staffgroupid ORDER BY sg.SeqNum, sg.Description, s.SeqNum, s.displayname "
	oRS.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		Do While Not(oRS.EOF)
			If strThisGroup <> ors("group").value Then
				If Not(blnFirstTime) Then
					Response.Write vbTab & "</group>" & vbCrLf
				End If
				blnFirstTime = False
				strThisGroup = oRS("group").value
				Response.Write vbTab & "<group name=""" & XMLEncode(Server.HTMLEncode("" & strThisGroup)) & """>" & vbCrLf
			End If
			Response.Write vbTab & vbTab & "<staff name=""" & XMLEncode(Server.HTMLEncode("" & oRS.Fields("DisplayName").Value)) & """>" & vbCrlF
			'Response.Write vbTab & vbTab & vbTab & "<name>" & oRS("displayname").value & "</name>" & vbCrlF
			Response.Write vbTab & vbTab & vbTab & "<position>" & XMLEncode(ors("title").value) & "</position>" & vbCrlF
			Response.Write vbTab & vbTab & vbTab & "<email>" & XMLEncode(ors("email").value) & "</email>" & vbCrlF
			Response.Write vbTab & vbTab & vbTab & "<description>" & XMLEncode(ors("description").value) & "</description>" & vbCrlF
			Response.Write vbTab & vbTab & "</staff>" & vbCrlF
			oRS.MoveNext
		Loop
		Response.Write vbTab & "</group>" & vbCrlF
	End If
	oRS.Close
	Response.Write "</TeamWarfareStaff>" & vbCrLf
oConn.Close
Set oConn = Nothing
Set oRs = Nothing

%>