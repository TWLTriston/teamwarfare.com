<% Option Explicit %>
<%
Response.Buffer = True
Server.ScriptTimeout = 4000

Dim strPageTitle

strPageTitle = "TWL: IP Tracker"

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

Dim blnBrowserTag 
blnBrowserTag = False

if not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If

Dim strFormName, strFormIP
strFormName	= Trim(Request("frm_name"))
strFormIP 	= Trim(Request("frm_ip"))

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("IP Address Tracker")
%>
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
<FORM NAME="frm_tracker" ID="frm_tracker" ACTION="tracker.asp" METHOD="post">
<TR><TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=2>Search Criteria</TH>
	</TR>
	<TR BGCOLOR="<%=bgcOne%>">
		<TD ALIGN=RIGHT>Player Name<BR>
		<FONT STYLE="color:#888888"><b>EXACT PLAYER NAME</b>: (Triston)</FONT></TD>
		<TD><INPUT TYPE=TEXT MAXLENGTH=50 SIZE=30 NAME=frm_name ID=frm_name VALUE="<%=strFormName%>"></TD>
	</TR>
	<TR BGCOLOR="<%=bgctwo%>">
		<TD ALIGN=RIGHT>EXACT IP Address<BR>
		<FONT STYLE="color:#888888">Form Of: (192.1.168.1)</FONT></TD>
		<TD><INPUT TYPE=TEXT MAXLENGTH=50 SIZE=30 NAME=frm_ip ID=frm_ip VALUE="<%=strFormIP%>"></TD>
	</TR>
	<TR BGCOLOR="<%=bgcOne%>">
		<TD COLSPAN=2 ALIGN=CENTER><INPUT TYPE=SUBMIT VALUE="Search Now"></TD>
	</TR>
	</TABLE>
</TD></TR>
</FORM>
</TABLE>
<%
Call ContentEnd()
%>
<%
If Len(strFormIP) > 0 Or Len(strFormName) > 0 Then 

	Call ContentStart("Search Results")
	strSQL = "EXECUTE TrackIPAddress "
	
	If Len(strFormIP) > 0 Then
		strSQL = strSQL & " @IpAddress = '" & CheckString(strFormIP) & "' "
		If Len(strFormName) > 0 Then
			strSQL = strSQL & ", "
		End If
	End If
	If Len(strFormName) > 0 Then
		strSQL = strSQL & " @PlayerHandle = '" & CheckString(strFormName) & "'"
	End If
	oRS.Open strSQL, oConn
	If Not(oRs.eof and oRs.bof) Then
		%>
		<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
		<TR><TD>
			<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 align="center">
			<TR BGCOLOR="#000000">
				<TH COLSPAN=5>Search Results</TH>
			</TR>
			<TR BGCOLOR="#000000">
				<TH>Player Name</TH>
				<TH>IP Address</TH>
				<TH>Last Logged</TH>
				<TH COLSPAN=2>Extend Results</TH>
			</TR>
			<%
			Do While Not(oRS.EOF)
				If bgc = bgcone then
					bgc = bgctwo
				else
					bgc = bgcone
				end if
				%>
				<TR BGCOLOR="<%=bgc%>">
					<TD nowrap="nowrap"><A HREF="viewplayer.asp?player=<%=Server.URLEncode(oRS.Fields("PlayerHandle").Value)%>"><%=Server.HTMLEncode(oRS.Fields("PlayerHandle").Value)%></A></TD>
					<TD nowrap="nowrap"><%=oRS.Fields("REMOTE_ADDR").Value%></TD>
					<TD nowrap="nowrap"><%=FormatDateTime(oRS.Fields("TimeLogged").Value, 0)%></TD>
					<TD nowrap="nowrap"><a href="/tracker.asp?frm_name=<%=oRS.Fields("PlayerHandle").Value%>">Other IPs (<%=oRs.Fields("OtherIps").Value%>)</A></TD>
					<TD nowrap="nowrap"><a href="/tracker.asp?frm_ip=<%=oRS.Fields("REMOTE_ADDR").Value%>">Other Players (<%=oRs.Fields("OtherPlayers").Value%>)</A></TD>
				</TR>
				<%
				oRS.MoveNext
			Loop
			%>
			</TABLE>
		</TD></TR></TABLE>
		<%
	Else
		Response.Write "<I>No matches found.</I>"
	End If
	oRS.Close
	Call ContentEnd()
End If
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>