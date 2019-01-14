<%@ Language=VBScript %>
<% 
   If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If
 Response.Buffer = true 

Dim strPageTitle

strPageTitle = "TWL: Error"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

On Error Resume Next
oConn.ConnectionString = Application("ConnectStr")
oConn.Open
If Err <> 0 Then
	Response.Clear
	Response.Redirect "/errors/offline.asp"
End If

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
  Const lngMaxFormBytes = 200

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL

    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0

  Set objASPError = Server.GetLastError
Call Contentstart("Error Occurred...")
%>
	<table width=760 border="0" cellspacing="0" cellpadding="0">

		<tr bgcolor=<%=bgcone%> height=25>
			<td align=center><p class=text><b>There has been an error... click back on your browser and insure all form fields are filled out (if applicable).<BR>
			The developers have been contacted and will look into this shortly.</b></p></td>
		</tr>
<%
  Dim bakCodepage
  on error resume next
	  bakCodepage = Session.Codepage
	  Session.Codepage = 1252
  on error goto 0
	'  Response.Write Server.HTMLEncode(objASPError.Category)
  text = objASPError.Category & vcrlf
  If objASPError.ASPCode > "" Then 
	'	Response.Write Server.HTMLEncode(", " & objASPError.ASPCode)
		text = text & objASPError.ASPCode & vbcrlf
  end if

  'Response.Write Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & "<br>"
  text = text & "(0x" & Hex(objASPError.Number) & ")" & vbcrlf

  If objASPError.ASPDescription > "" Then 
		'Response.Write Server.HTMLEncode(objASPError.ASPDescription)
		text = text & objASPError.ASPDescription & vbcrlf
  elseIf (objASPError.Description > "") Then 
		 'Response.Write Server.HTMLEncode(objASPError.Description)
		 text = text & objASPError.Description & vbcrlf
  end if
  blnErrorWritten = False

  ' Only show the Source if it is available and the request is from the same machine as IIS
  If objASPError.Source > "" Then
    strServername = LCase(Request.ServerVariables("SERVER_NAME"))
    strServerIP = Request.ServerVariables("LOCAL_ADDR")
    strRemoteIP =  Request.ServerVariables("REMOTE_ADDR")
    If (strServername = "localhost" Or strServerIP = strRemoteIP) And objASPError.File <> "?" Then
      'Response.Write Server.HTMLEncode(objASPError.File)
	  	text = text & objASPError.File & vbcrlf
      If objASPError.Line > 0 Then 
			'Response.Write ", line " & objASPError.Line
		 	text = text & objASPError.line & vbcrlf
	  	end if
      If objASPError.Column > 0 Then 
				'Response.Write ", column " & objASPError.Column
				text = text & objASPError.Column & vbcrlf
	  	end if
      'Response.Write "<br>"
     	'Response.Write "<font style=""COLOR:FFFFFF; FONT: 8pt/11pt courier new""><b>"
      'Response.Write Server.HTMLEncode(objASPError.Source) & "<br>"
      text = text & objASPError.Source & vbCRLF
      If objASPError.Column > 0 Then 
				'Response.Write String((objASPError.Column - 1), "-") & "^<br>"
				text = text & String((objASPError.Column - 1), "-") & vbcrlf
	  	end if
      'Response.Write "</b></font>"
      blnErrorWritten = True
    End If
  End If

  If Not blnErrorWritten And objASPError.File <> "?" Then
    'Response.Write "<br><b>" & Server.HTMLEncode(  objASPError.File)
    If objASPError.Line > 0 Then 
			'Response.Write Server.HTMLEncode(", line " & objASPError.Line)
			text = text & "Line #" & objASPError.Line & vbCRLF
			end if
   	If objASPError.Column > 0 Then 
      'Response.Write ", column " & objASPError.Column
      text = text & "Column: " & objASPError.column & vbCRLF
    end if
    'Response.Write "</b><br>"
  End If
	page = Request.ServerVariables("SCRIPT_NAME") 
	line = objASPError.Line

	text = text & "Browser: " & Request.ServerVariables("HTTP_USER_AGENT") & vbcrlf 
  strMethod = Request.ServerVariables("REQUEST_METHOD")

'  Response.Write strMethod & " "
  text = text & "Page: " & strmethod & vbcrlf

  If strMethod = "POST" Then
    'Response.Write Request.TotalBytes & " bytes to "
    text = text & Request.TotalBytes & " bytes to "
  End If
  text = text & Request.ServerVariables("SCRIPT_NAME") & vbCRLF
  'Response.Write Request.ServerVariables("SCRIPT_NAME")
  If strMethod = "GET" Then
  	text = text & Request.QueryString
  End If

  lngPos = InStr(Request.QueryString, "|")

  If lngPos > 1 Then
   ' Response.Write "?" & Left(Request.QueryString, (lngPos - 1))
    text = text & "?" & Left(Request.QueryString, (lngPos - 1)) & vbcrlf
  End If 
  If strMethod = "POST" Then
    text = text & "POST Data: " & vbCRLF
    If Request.TotalBytes > lngMaxFormBytes Then
'       Response.Write Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes & "")) & " . . ."
       ' text = text & Left(Request.Form, lngMaxFormBytes & "") & " . . ." & vbcrLF
    Else
'      Response.Write Server.HTMLEncode(Request.Form)
      text = text & Request.Form & vbcrlf
    End If
  End if
  If Session("uName") = "Triston" Then
  	Response.Write "<tr><td bgcolor=" & bgctwo & ">Line: " & line & "<br />" & Replace(text, vbCrLf, "<br />") & "</td></tr>"
  End If
  	
  strSQL = "INSERT INTO tbl_errors ( ErrorMessage, Player, page, line) VALUES ('" & CheckString(Text) & "', '" & CheckString(Session("uName")) & "', '" & CheckString(page) & "','" & CheckString(line) & "')"
  oconn.execute (strSQL)
  
  
'	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
'	smtp = "127.0.0.1"
'	Mailer.RemoteHost  = smtp
'	Mailer.FromName    = "TWL: SQL Error Gotten By: " & Session("uName")
'	mailer.FromAddress = "triston@teamwarfare.com"
'	subject = "ASP Error 500.100 Recieved"
'	Mailer.Subject     = subject
'	Mailer.BodyText    = text
'	Mailer.AddRecipient "Triston", "automailer@teamwarfare.com"	
'	'on error resume next
'	If Not(Mailer.SendMail) Then
	'	Response.Write "Mail send failure. Error was " & Mailer.Response & smtp
'	End If

'	set mailer = nothing
			
%>
</TABLE>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>