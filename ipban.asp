<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: IP Ban"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

if not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("")
%>
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
<TR>
	<TD>
		<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=4>
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2>Ban the fuckers</TH>
		</TR>
		<%
		strSQL = "SELECT IP FROM tbl_ban_ip ORDER BY IP "
		oRS.Open strSQL, oConn
		If Not(oRS.EOF AND oRS.BOF) Then
			Do While Not(oRS.EOF)
				If bgc = bgcone Then
					bgc = bgctwo
				Else
					bgc = bgcone
				End If
				%>
				<TR BGCOLOR="<%=bgc%>">
					<TD><B>Banned:</B> <a href="tracker.asp?frm_ip=<%=oRS.Fields("IP").Value%>"><%=oRS.Fields("IP").Value%></a></TD>
					<TD><A href="saveitem.asp?saveType=DeleteIPBan&IP=<%=Server.URLEncode(oRS.Fields("IP").Value)%>">remove</A></TD>
				</TR>
				<%
				oRS.MoveNext
			Loop
		Else
		%>
		<TR BGCOLOR="<%=bgcone%>">
			<TD COLSPAN=2>No blessed soul is on our list.</TD>
		</TR>
		<%
		End If
		oRS.NextRecordset 
		%>		
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2>Add!</TH>
		</TR>
		<FORM NAME="frm_ipban" ACTION="saveitem.asp" METHOD="get">
		<TR BGCOLOR="#000000">
			<TD ALIGN=CENTER COLSPAN=2><INPUT TYPE=TEXT NAME="ip"></TD>
		</TR>
		<TR BGCOLOR="#000000">
			<TD COLSPAN=2 ALIGN=CENTER><INPUT TYPE="SUBMIT" VALUE="Ban It!"></TD>
		</TR>
		<INPUT TYPE="HIDDEN" NAME="SaveType" VALUE="AddIPBan">
		</FORM>
		</TABLE>
	</TD>
</TR>
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