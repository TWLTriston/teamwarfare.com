<%
Dim Mailer
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.FromName = "Auto Mailer"
Mailer.FromAddress = "automailer@web.teamwarfare.com"
Mailer.Remotehost = "127.0.0.1"
Mailer.AddRecipient "Triston", "triston@gmail.com"

Mailer.Subject = "Test"
Mailer.BodyText = "Testing"
if not Mailer.SendMail then
  if Mailer.Response <> "" then
    strError = Mailer.Response
  else
    strError = "Unknown"
  end if
  Response.Write "Mail failure occured. Reason: " & strError
end if
Response.Write "It worked!"
%>