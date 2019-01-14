<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Forgot your Password?"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim pw_length, pw_chars, new_pw, rnum, c
Dim encpass, enc2pass, session_id, Text, done
Dim PlayerHandle,PlayerEmail, mailer, i

Dim Msg
If Request.Form("hdnSubmit") = "1" Then	
	If Request.Form("PlayerHandle") = "" Then Msg = Msg & "<li>Player Handle"

	' make sure player exists
	strSQL = "SELECT * FROM tbl_Players WHERE PlayerHandle='" & CheckString(Request.Form("PlayerHandle")) & "'"
	Set oRS = oConn.Execute(strSQL)
		If oRS.EOF Then Response.Redirect("errorpage.asp?error=13")
	oRS.Close
	
	If Msg = "" Then

		' REDACTED
	
		
		' generate session id
		'session_id = now & " " & Request.ServerVariables("REMOTE_ADDR") & " - " & int(10 * rnd + 1)
		Randomize
		For c = 1 To 3
			rnum = Int((len(pw_chars) * Rnd) + 1)
			session_id = session_id & Mid( pw_chars, rnum, 1 )
		Next
	
		' add PlayerNewPassword and PlayerNewPasswordID to tbl_Players
		strSQL = "UPDATE tbl_Players SET "
		strSQL = strSQL & "PlayerNewPassword='" & enc2pass & "',"
		strSQL = strSQL & "PlayerNewPasswordID='" & session_id & "'"
		strSQL = strSQL & " WHERE PlayerHandle='" & CheckString(Request.Form("PlayerHandle")) & "'"
		oConn.Execute(strSQL)

		
		' get player handle/email
		strSQL = "SELECT * FROM tbl_Players WHERE PlayerHandle='" & CheckString(Request.Form("PlayerHandle")) & "'"
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRS.BOF) Then
			PlayerHandle = oRS("PlayerHandle")
			PlayerEmail = oRS("PlayerEmail")
		End If
		oRS.NextRecordset 
		
		' send e-mail w/link to activate.asp?session_id
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.RemoteHost  = "127.0.0.1"
		Mailer.FromName    = "TWL"
		Mailer.FromAddress = "automailer@web.teamwarfare.com"
		Mailer.AddRecipient PlayerHandle, PlayerEmail
		Mailer.Subject     = "TWL: New Password Request"

		Text = PlayerHandle & ", you are receiving this message because a new password request was submitted through "
		Text = Text & "our website.  If you did not submit this request, simply disregard this e-mail.  If you did, "
		Text = Text & "please follow the link below to activate your new password." & vbCrLf & vbCrLf
		Text = Text & "New Password: " & new_pw & vbCrLf
		Text = Text & "Activation Code: " & session_id & vbCrLf
		Text = Text & "http://www.teamwarfare.com/activatePassword.asp"
		Text = Text & vbCrLf & vbCrLf & date & vbCrLf & "- TWL"
			
			
		Mailer.BodyText    = text
		If Not Mailer.SendMail Then
			Response.Write "[" & Mailer.Response & "]"
		End If
		set mailer = nothing
		
		done = 1	
	End If
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%	Call ContentStart("Forgot Your Password?") %>
	<table width=760 border="0" cellspacing="0" cellpadding="2" ALIGN=CETER>
	<tr><td>
		<% If Msg <> "" Then %>
			<table width=300 align=center border=0>
				<tr>
					<td>
					An error has occured; you left the following field(s) blank:<br>
					<%=Msg%>
					</td>
				</tr>
			</table>
		<% End If
		If done = 0 Then
		%>
	                  <p class=small align=center>Please enter your handle:
	                  
					  <form action="forgotPassword.asp" method="post">
					  <input type="hidden" name="hdnSubmit" id="hdnSubmit" value="1">
						<input type=text name=PlayerHandle value="<%=Server.HTMLEncode(Request.form("PlayerHandle"))%>"><br>
						<br>
						<input type=submit name="btnSubmit" id="btnSubmit" value=" Get Password ">
					  </form>
					  
					  </p>
					  <p class=small>This form will send an email to the account which is registered under your playername,
					  this email will contain an activation code for the new password, as well as a new password. Submission of
					  this form will not change your password, you must use the activation code first.
					  </P>
		<% ElseIf done = 1 Then %>
			<p class=small align=center>Your request has been submitted. Please check your email for instructions.
		<% End If %>

	</td></tr>
	</table>            
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>
