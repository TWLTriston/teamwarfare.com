<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Activate Password"

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

Dim bDone
bDone = False

If Request.Form("submit") = " Activate Password " Then
	Dim Msg
	If Request.Form("PlayerHandle") = "" Then Msg = "<li>Your Handle"
	If Request.Form("actCode") = "" Then Msg = Msg & "<li>Activation code"
	
	If Msg = "" Then
		Dim PlayerHandle,actCode
		
		PlayerHandle = Replace(Request.Form("PlayerHandle"), "'", "''")
		actCode = Replace(Request.Form("actCode"), "'", "''")
		
		strSQL = "SELECT * FROM tbl_Players WHERE PlayerHandle='" & PlayerHandle & "' AND "
		strSQL = strSQL & "PlayerNewPasswordID='" & actCode & "'"
		oRS.Open strSQL, oConn
		If oRS.EOF Then 
			oRs.Close
			oConn.Close
			Set oConn = Nothing
			Set oRs = Nothing
			Response.Clear 			
			Response.Redirect("/errorpage.asp?error=14")
		Else
			strSQL = "UPDATE tbl_Players SET "
			strSQL = strSQL & "PlayerPassword='" & oRS("PlayerNewPassword") & "',"
			strSQL = strSQL & "PlayerNewPassword='',"
			strSQL = strSQL & "PlayerNewPasswordID=''"
			strsQL = strSQL & " WHERE PlayerHandle='" & PlayerHandle & "'"
			oConn.Execute(strSQL)			
			bDone = True
		End If
		oRS.NextRecordSet
	End If
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Activate Password")
If Not(bDone) Then 
	%>
	<table width=500 border="0" cellspacing="0" cellpadding="2" ALIGN=CENTER>
	<tr><td>
		Please fill out the forms below to activate your new password.<br>
		Refer to your e-mail for the activation code, and your new password. If you lost the original email,
		please request a new one <a href="/forgotpassword.asp">here</a>.
		</TD>
	</TR>
	</TABLE>
	<BR>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
	<TR>
		<form action=activatePassword.asp method=post id=form1 name=form1>
		<TD>
		<table width=300 cellspacing=1 cellpadding=2 border=0 align=center>
			<tr BGCOLOR="<%=bgcone%>">
				<td ALIGN=RIGHT>Your Handle:</td>
				<td><Input type=text name=PlayerHandle value="<%=Request.Form("PlayerHandle")%>">
			</tr>
			<tr BGCOLOR="<%=bgctwo%>">
				<td ALIGN=RIGHT>Activation Code:</td>
				<td><Input type=text name=actCode value="<%=Request.Form("actCode")%>">
			</tr>			
			<tr BGCOLOR="<%=bgcone%>">
				<td colspan=2 align=center><input type=submit name=submit value=" Activate Password "></td>
			</tr>
		</table>
		</TD>
		</form>
	</TR>
	</TABLE>
		<%Else %>
		<p align=center class=small>
			Your password has been successfully updated!<br>
			Refer to your email for your new password.
		</p>
		<%End If%>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>