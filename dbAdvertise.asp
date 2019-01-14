<%' Option Explicit %>
<%
Server.ScriptTimeout = 1000
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Advertise"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<%
If InStr(Request.Form("txtName"), "@") <= 0 AND Len(Request.Form("txtName")) > 0 Then

			strBody = "Name: " & Request.Form("txtName") & vbCrLf
			strBody = strBody & "Email: " & Request.Form("txtEmail") & vbCrLf
			strBody = strBody & "Phone: " & Request.Form("txtPhone") & vbCrLf
			strBody = strBody & "Company: " & Request.Form("txtCompany") & vbCrLf
			strBody = strBody & "Website: " & Request.Form("txtCompanyURL") & vbCrLf
			strBody = strBody & "Details: " & vbCrLf & Request.Form("txtDetails") & vbCrLf
			strBody = strBody & "IP: " & vbCrLf & Request.ServerVariables("REMOTE_ADDR") & vbCrLf


			Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			Mailer.RemoteHost  = "127.0.0.1"
			Mailer.FromName    = "TeamWarfare Automailer"
			mailer.FromAddress = "automailer@teamwarfare.com"
			Mailer.AddRecipient "Triston", "triston@gmail.com"
			Mailer.AddRecipient "Polaris", "Polaris@teamwarfare.com"
			Mailer.Subject     = "TWL: Advertiser Interest"
			Mailer.BodyText    = strBody
		  If Not(Mailer.SendMail) Then
		    if Mailer.Response <> "" then
		      strError = Mailer.Response
		    else
		      strError = "Unknown"
		    end if
		    Response.Write "Mail failure occured. Reason: " & strError
			End If
			set mailer = nothing
' 		Response.Write "Sent mail."
End if
	
	Response.Clear
	Response.Redirect "advertise_thankyou.asp"
%>