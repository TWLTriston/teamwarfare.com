<% Option Explicit %>
<!-- #include file="../include/adovbs.inc" -->
<%
Dim oRs, strSQL, oConn, oCmd

Dim intIdentifierID, strIdentiferValue, strPlayerHandle
Dim strIdentifierName

strIdentifierName = Request.QueryString("i")
If Len(strIdentifierName) = 0 Then
	strIdentifierName = Request.Form("i")
End If

strIdentiferValue = Request.QueryString("v")
If Len(strIdentiferValue) = 0 Then
	strIdentiferValue = Request.Form("v")
End If

strPlayerHandle = Request.QueryString("p")
If Len(strPlayerHandle) = 0 Then
	strPlayerHandle = Request.Form("p")
End If

Dim blnIdentifierMatch, blnPlayerValid, blnIdentifierValueValid, blnIdentifierValid

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Set oCmd = Server.CreateObject("ADODB.Command")
With oCmd
  .CommandText = "IdentifierCheck"
  .ActiveConnection = oConn
  .CommandType = adCmdStoredProc
  .Parameters.Append .CreateParameter("@IdentifierName", adVarchar, adParamInput, 50, strIdentifierName)
 	.Parameters.Append .CreateParameter("@PlayerHandle", adVarchar, adParamInput, 50, strPlayerHandle)
 	.Parameters.Append .CreateParameter("@IdentifierValue", adVarchar, adParamInput, 50, strIdentiferValue)
	.Parameters.Append .CreateParameter("@IdentifierMatch", adInteger, adParamOutput, 50)
	.Parameters.Append .CreateParameter("@PlayerValid", adInteger, adParamOutput, 50)
	.Parameters.Append .CreateParameter("@IdentifierValueValid", adInteger, adParamOutput, 50)
	.Parameters.Append .CreateParameter("@IdentifierValid", adInteger, adParamOutput, 50)
End With

oCmd.Execute

blnIdentifierMatch = oCmd.Parameters("@IdentifierMatch").Value
blnPlayerValid = oCmd.Parameters("@PlayerValid").Value
blnIdentifierValueValid = oCmd.Parameters("@IdentifierValueValid").Value
blnIdentifierValid = oCmd.Parameters("@IdentifierValid").Value


Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf
Response.Write "<idlookup identifier=""" & Server.HTMLEncode(strIdentifierName & "") & """ loginname=""" & Server.HTMLEncode(strPlayerHandle & "") & """ idvalue=""" & Server.HTMLEncode(strIdentiferValue & "") & """>"
Response.Write vbTab & "<returncodes idvalid=""" & blnIdentifierValid & """ "
Response.Write " loginvalid=""" & blnPlayerValid & """ "
Response.Write " idvaluevalid=""" & blnIdentifierValueValid & """ "
Response.Write " idmatch=""" & blnIdentifierMatch & """ />"

If Not(CBool(blnPlayerValid)) Then
	Response.Write vbTab & "<error number=""1"">No match on specified login name. (Not in DB)</error>"
End If
If Not(CBool(blnIdentifierValid)) Then
	Response.Write vbTab & "<error number=""2"">Invalid identifier specified. (Not in DB)</error>"
End If
If Not(CBool(blnIdentifierMatch)) Then
	Response.Write vbTab & "<error number=""3"">ID does not belong to login name</error>"
End If
If Not(CBool(blnIdentifierValueValid)) Then
	Response.Write vbTab & "<error number=""4"">No match of specified identifier value. (Not in DB)</error>"
End If
Response.Write "</idlookup>"


Set oCmd = Nothing
Set oConn = Nothing

%>