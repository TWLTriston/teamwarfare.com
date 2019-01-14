<%
Option Explicit

Server.ScriptTimeout = 10000
Response.Buffer = True


Dim strBody
strBody = ""
strBody = strBody & "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" & vbCrLf
strBody = strBody & "" & vbCrLf
strBody = strBody & "<html>" & vbCrLf
strBody = strBody & "<head>" & vbCrLf
strBody = strBody & "	<title>Panel Invitation</title>" & vbCrLf
strBody = strBody & "</head>" & vbCrLf
strBody = strBody & "<body bgcolor=""#000000"">" & vbCrLf
strBody = strBody & "" & vbCrLf
strBody = strBody & "<table width=""600"" cellpadding=""4"" cellspacing=""0"" align=""center"" bgcolor=""#000000"" style=""border: 1px solid #444444;"">" & vbCrLf
strBody = strBody & "<tr>" & vbCrLf
strBody = strBody & "	<td><img src=""http://www.teamwarfare.com/email/logo_panel.jpg"" width=""299"" height=""80"" alt="""" border=""0""></td>" & vbCrLf
strBody = strBody & "</tr>" & vbCrLf
strBody = strBody & "<tr>" & vbCrLf
strBody = strBody & "  <td>" & vbCrLf
strBody = strBody & "  	<font face=""Verdana"" size=""3"" color=""#ffd142"" style=""font-size: 14px;"">" & vbCrLf
strBody = strBody & "  		<b>What do you think of the Xbox, Playstation 3, PSP...?</b><br />" & vbCrLf
strBody = strBody & "  	</font>" & vbCrLf
strBody = strBody & "  	<font face=""Verdana"" size=""2"" color=""#ffffff"" style=""font-size: 10px;"">" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  		We are building  a team of members to help us understand the gamer's opinion on new products, technology, games, consoles, gadgets, computers and other things you use.<br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  		The team is called the <a href=""http://survey.sotech.com/557001/start.asp?s=3"" target=""_blank""><font color=""#FFD142"">Team Warfare League Panel</font></a>.<br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  		All we ask you to do is participate in an online survey about once a month. When you do, we may reward you with cash, gift certificates, other prizes and information. Please know that all of the information you provide is always kept confidential.<br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  		<a href=""http://survey.sotech.com/557001/start.asp?s=3"" target=""_blank""><font color=""#FFD142"">Join the Team today!</font></a><br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  		<a href=""http://survey.sotech.com/557001/start.asp?s=3"" target=""_blank""><img src=""http://www.teamwarfare.com/email/pic_joinnow.jpg"" width=""260"" height=""332"" alt="""" border=""0"" align=""right""></a>" & vbCrLf
strBody = strBody & "  		" & vbCrLf
strBody = strBody & "  		If you have any further questions, please visit the F.A.Q. and Privacy Policy Section on our new panel website - <a href=""http://www.teamwarfareleaguepanel.com"" target=""_blank""><font color=""#FFD142"">www.teamwarfareleaguepanel.com</font></a>.<br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "" & vbCrLf
strBody = strBody & "  		<font face=""Verdana"" size=""3"" color=""#ffd142"" style=""font-size: 14px;""><b>Privacy And Confidentiality</b></font><br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  		Team Warfare has partnered with Socratic  Technologies, Inc. a full-service market research firm that values your privacy. All of your responses will be kept confidential and only reported in the aggregate.<br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  		Your personal information has not, and will not be sold or traded to any companies. For more information about Socratic's privacy policy,  please visit <a href=""http://www.sotech.com/main/eval.asp?PID=106"" target=""_blank""><font color=""#FFD142"">www.sotech.com/main/eval.asp?PID=106</font></a>.<br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  		If you have questions about the survey, please contact the Member Services Manager, Ryan Hill, at <a href=""mailto:Ryan.Hill@TeamWarfareLeaguePanel.com?subject=Project%20557-001%20Team%20Warfare"" target=""_blank""><font color=""#FFD142"">Ryan.Hill@TeamWarfareLeaguePanel.com</font></a> and reference project 557-001. Ryan can also be reached at 1-800-5-SOCRATIC (1-800-576-2728), or 001-415-430-2200 outside USA.<br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  		If you would like to be removed from our contact list, please reply to this email and type ""Remove - Team Warfare"" in the subject line, or call 1-800-5-SOCRATIC (1-800-576-2728) or 001-415-430-2200 outside USA. You may also request your removal by writing to: Socratic Technologies, Inc., 2505 Mariposa Street, San Francisco, CA 94110 USA.<br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  		Socratic Technologies, Inc. is a member of the Interactive Marketing Research Organization (IMRO) and we subscribe to the privacy policies and code of research ethics published by this group.  You can visit IMRO at <a href=""http://www.imro.org"" target=""_blank""><font color=""#FFD142"">www.imro.org</font></a> for more information.<br />" & vbCrLf
strBody = strBody & "  		<br />" & vbCrLf
strBody = strBody & "  	</font>" & vbCrLf
strBody = strBody & "	</td>" & vbCrLf
strBody = strBody & "</tr>" & vbCrLf
strBody = strBody & "</table>" & vbCrLf
strBody = strBody & "" & vbCrLf
strBody = strBody & "</body>" & vbCrLf
strBody = strBody & "</html>" & vbCrLf

Dim oConn, oRs, strSQL
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Dim objEmail

strsQL = "SELECT TOP 200 * FROM tbl_email WHERE EMailSent = 0"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	Do While Not (oRs.EOF)
		Set objEmail = Server.CreateObject("SMTPsvg.Mailer")
		objEmail.RemoteHost  = "127.0.0.1"
		objEmail.FromName    = "Team Warfare Panel"
		objEmail.FromAddress = "teamwarfarepanel@teamwarfare.com"
		objEmail.AddRecipient oRs.Fields("EmailAddress").Value, oRs.Fields("EmailAddress").Value
		objEmail.Subject     = "Team Warfare League Survey Panel Invitation"
		objEmail.ContentType = "text/html"
		objEmail.BodyText    = strBody
		On Error Resume Next
		objEmail.SendMail
		On Error Goto 0
		Set objEmail = Nothing
		
		strSQL = "UPDATE tbl_email SET EmailSent = 1 WHERE EmailID = '" & oRs.Fields("EmailID").Value & "'"
		oConn.Execute(strSQL)

		Response.Write oRs.Fields("EmailAddress").Value & "<br />" & vbCrLf
		Response.Flush

		oRs.MoveNext
	Loop
End If
oRs.Close
Response.Write "Email Sent, "

strSQL = "SELECT COUNT(*) FROM tbl_email WHERE EmailSent = 0"
oRs.Open strSQL, oConn
Response.Write oRs.Fields(0).Value
oRs.Close

oConn.Close

Set oRs = Nothing
Set oConn = Nothing
%>