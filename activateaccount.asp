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

Dim strPlayerName, strActivationCode
strPlayerName = Request.QueryString("playername")
strActivationCode = Request.QueryString("actCode")

Dim bDone, bInvalidPlayerName, bInvalidActivation
bDone = False
bInvalidPlayerName = False
bInvalidActivation = False
If len(strPlayerName) > 0 AND Len(strActivationCode) > 0 Then
	strSQL = "SELECT PlayerID, ActivationCode FROM tbl_players WHERE PlayerHandle = '" & CheckString(strPlayerName) & "'"
	oRs.Open strSQL, oConn
	If Not(ors.EOF and ors.BOF) Then
		If cStr(oRs.Fields("ActivationCode").Value) <> cStr(strActivationCode) Then
			bInvalidActivation = True
		Else
			strSQL = "UPDATE tbl_players SET PlayerActive = 'Y' WHERE PlayerID = " & ors.Fields("PlayerID").Value 
			oConn.Execute (strSQL)
			bDone = True
		End If
	Else
		bInvalidPlayerName = True
	End If
	ors.NextRecordset 
ElseIF Len(strActivationCode) = 0 AND Request.QueryString ("submit") <> "" Then
	bInvalidActivation = True
ElseIf Len(strPlayerName) = 0  AND Request.QueryString ("submit") <> "" Then
	bInvalidPlayerName = True
ENd If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Activate Account")
If Not(bDone) Then 
	%>
	<table width=500 border="0" cellspacing="0" cellpadding="2" ALIGN=CENTER>
	<tr><td>Please fill out the forms below to activate your account.</TD></TR>
	</TABLE>
	<BR>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
	<form action=activateaccount.asp method=GET id=form1 name=form1>
	<TR>
		<TD>
		<table width=300 cellspacing=1 cellpadding=2 border=0 align=center>
			<% If bInvalidPlayerName Then %>
			<TR BGCOLOR="<%=bgcone%>">
				<TD COLSPAN=2><FONT COLOR="#FF0000">Invalid player name, please reenter.</FONT></TD>
			</TR>
			<% End If %>
			<% If bInvalidActivation Then %>
			<TR BGCOLOR="<%=bgcone%>">
				<TD COLSPAN=2><FONT COLOR="#FF0000">Invalid activation code, please reenter, or have another code emailed to you.<BR>
				<A href="activate.asp">Get a new code here.</A></FONT></TD>
			</TR>
			<% End If %>
			<tr BGCOLOR="<%=bgcone%>">
				<td ALIGN=RIGHT>Your Player Name:</td>
				<td><Input type=text name=playername value="<%=Request.Querystring("playername")%>">
			</tr>
			<tr BGCOLOR="<%=bgctwo%>">
				<td ALIGN=RIGHT>Activation Code:</td>
				<td><Input type=text name=actCode value="<%=Request.Querystring("actCode")%>">
			</tr>			
			<tr BGCOLOR="<%=bgcone%>">
				<td colspan=2 align=center><input type=submit name=submit value=" Activate Password "></td>
			</tr>
		</table>
		</TD>
	</TR>
	</form>
	</TABLE>
		<%Else %>
		<p align=center class=small>
			Your account has been successfully activated!<br>
			You may now login.
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