<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Error Mail"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
Dim body, code, message, text, uname, mailer
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Error Mail") %>
    <table width="90%" border="0">
      <tr>
        <td><p class=small>
	<%
	body=request.form("MailBody")
	code = request.form("ErrorCode")
	message = request.form("ErrorMessage")
	uname = session("uName")
	if uname = "" then 
		uname =	"Not logged in/no cookie present"
	end if
	
	if body <> "" AND len(message) < 75 then
		text = body & vbcrlf & vbcrlf
		text = text & "Code Number: " & code & vbcrlf & vbcrlf
		text = text & "Message Associated with the code: " & message & vbcrlf & vbcrlf
		text = text & "Logged username: " & uname  & vbcrlf  & vbcrlf 
		text = text & "Logged IP: " & Request.ServerVariables("REMOTE_ADDR") & vbcrlf 
		text = text & vbcrlf 
		Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
		Mailer.RemoteHost  = "127.0.0.1"
		Mailer.FromName    = "TWL: Error mailer"
		mailer.FromAddress = "automailer@web.teamwarfare.com"
		Mailer.Subject     = code & ": Error message was recieved, follows is their response."
		Mailer.BodyText    = text
		Mailer.AddRecipient "triston@gmail.com", "triston@gmail.com"
		Mailer.AddRecipient "triston@teamwarfare.com", "triston@teamwarfare.com"
		Mailer.AddRecipient "engineering@teamwarfare.com", "engineering@teamwarfare.com"
		if false then 
			strsql = "Select PlayerHandle, PlayerEmail from sysadmins s, tbl_players p WHERE p.playerID = s.AdminID and s.SendEmail = 1"
			oRS.open strsql, oConn
			if not (oRS.eof and oRS.bof) then
				do while not(oRS.eof)
					Mailer.AddRecipient oRS.fields(0).value, oRS.fields(1).value
					oRS.movenext
				loop
			end if
			oRS.NextRecordset 
		end if
		if Mailer.SendMail then
			response.write "Mail has been sent, an admin should be in touch with you in the next 24 hours, if no one has responded, check #teamwarfare on irc.dynamix.com"
		else
		  Response.Write "Mail send failure. Error was " & Mailer.Response
		end if
		set mailer = nothing
	else
		response.write "Empty body message, no mail sent."
	end if
%>
</td></tr></table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>
