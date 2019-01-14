<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Activate Account"

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

Dim intPlayerID, Playerhandle, PlayerEmail

If Request.Form("submit") = " Activate Account " Then
	Dim Msg
	If Request.Form("PlayerHandle") = "" And Request.Form("email") = "" Then Msg = Msg & "<BR>You must enter either player handle or mail"

	' make sure player exists
	If Len(Request.Form("PlayerHandle")) > 0 Then
		strSQL = "SELECT PlayerID, PlayerHandle, PlayerEmail FROM tbl_Players WHERE PlayerHandle='" & Replace(Request.Form("PlayerHandle"), "'", "''") & "'"
	Else
		 strSQL = "SELECT PlayerID, PlayerHandle, PlayerEmail FROM tbl_Players WHERE PlayerEmail = '" & CheckString(Request.Form("Email")) & "'"
	End If
	Set oRS = oConn.Execute(strSQL)
	If oRS.EOF Then 
		msg = msg & " <BR>Unknown player name or email address, please try again."
	Else
		intPlayerID = oRS.Fields("PlayerID").Value 
		PlayerHandle = oRS.Fields ("PlayerHandle").Value 
		PlayerEmail = oRS.Fields ("PlayerEmail").Value 
	End If
	oRS.Close
	If Msg = "" Then

		Dim pw_length, pw_chars, new_pw, rnum, c
		Dim encpass, enc2pass, session_id, Text, done
		Dim  mailer, i, testtest
	
		session_id = Session.SessionID 
		' add PlayerNewPassword and PlayerNewPasswordID to tbl_Players
		strSQL = "UPDATE tbl_Players SET "
		strSQL = strSQL & "ActivationCode='" & session_id & "'"
		strSQL = strSQL & " WHERE PlayerID = " & intPlayerID
'		Response.Write strSQL
		oConn.Execute(strSQL)

		' send e-mail w/link to activate.asp?session_id
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.RemoteHost  = "127.0.0.1"
		Mailer.FromName    = "TWL"
		mailer.FromAddress = "automailer@web-mailer.teamwarfare.com"
		Mailer.AddRecipient PlayerHandle, PlayerEmail
		Mailer.Subject     = "TWL: Account Activation"
		
		Text = PlayerHandle & ", you are receiving this message because an activation code request was submitted through "
		Text = Text & "our website.  If you did not submit this request, simply disregard this e-mail.  If you did, "
		Text = Text & "please follow the link below to activate your account." & vbCrLf & vbCrLf
		Text = Text & "Activation Code: " & session_id & vbCrLf
		Text = Text & "http://www.teamwarfare.com/activateaccount.asp?playername=" & Server.URLEncode(PlayerHandle) & "&actcode=" & session_id
		Text = Text & vbCrLf & vbCrLf & date & vbCrLf & "- TWL"
					
		Mailer.BodyText    = text
		'Keep out the bad guys
		If instr(1, PlayerEmail,"cjb.net") = 0 Then
		  Call Mailer.SendMail
		End If
		set mailer = nothing
		
		done = 1	
	End If
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%	Call ContentStart("Need to activate your account?") %>
	<table width=760 border="0" cellspacing="0" cellpadding="2" ALIGN=CETER>
	<tr><td>
		<% If Request.QueryString ("error") = "1" Then %>
			<table width=700 align=center border=0>
				<tr>
					<td><FONT COLOR="#FF0000">
					An error has occured. <BR>
					Your account has been deactivated, please fill out the form below to have an activation code emailed to you.
					</FONT>
					</td>
				</tr>
			</table>
		<% End If %>
		<% If Msg <> "" Then %>
			<table width=700 align=center border=0>
				<tr>
					<td><FONT COLOR="#FF0000">
					<B>An error has occured:</B>
					<%=Msg%>
					</td>
				</tr>
			</table>
		<% End If
		If done = 0 Then
		%>
		<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER>
		<form action="activate.asp" method=post id=form1 name=form1>
		<TR>
			<TD>
				<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=4 BGCOLOR="#444444">
					<TR BGCOLOR="#000000">
						<TH COLSPAN=2>Resend Activation Email</TH>
					</TR>
					<TR BGCOLOR="<%=bgcone%>">
						<TD ALIGN=RIGHT>Please enter your handle:</TD>
						<TD><input type=text name=PlayerHandle value="<%=Request.form("PlayerHandle")%>"></TD>
					</TR>
					<TR BGCOLOR="<%=bgctwo%>">
						<TD ALIGN=RIGHT>Or enter your email:</TD>
						<TD><input type=text name=Email value="<%=Request.form("EMail")%>"></TD>
					</TR>
					<TR BGCOLOR="#000000">
						<TD COLSPAN=2 ALIGN=CENTER><input type=submit name=submit value=" Activate Account "></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		</form>
		</TABLE>
		<BR>
		<CENTER>
		This form will send an email to the account specified which will contain an activation code for your account.
		</CENTER>
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
