<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Ladder Activity Report"

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

If Not(bSysAdmin or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim ActiveTeams, AvailTeams, RestTeams, TotalTeams, LadderName
Dim intCounter, intPercentActive
intCounter = 0
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call Content2BoxStart("Ladder Activity Report")
%>
	<table width=780 border="0" cellspacing="0" cellpadding="0" BACKGROUND="">
	<tr>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	<td width=380>
	<%
	strsql = "exec sp_ladderactivity"
	ors.open strsql, oconn
	if not(ors.eof and ors.bof) then
		ActiveTeams = 0
		AvailTeams = 0
		RestTeams = 0
		TotalTeams = 0
		do while not(ors.eof)
			If intCounter Mod 2 = 0 And intCounter <> 0 Then
				%>
				</td>
				<td><img src="/images/spacer.gif" width="5" height="1"></td>
				</tr>
				</table>
				<%
				Call Content2BoxEnd()
				Call Content2BoxStart("")
				%>
				<table width=780 border="0" cellspacing="0" cellpadding="0" BACKGROUND="">
				<tr>
				<td><img src="/images/spacer.gif" width="5" height="1"></td>
				<td width=380>
				<%				
			ElseIf intCounter <> 0 Then
				%>
				</td>
				<td><img src="/images/spacer.gif" width="10" height="1"></td>
				<td width=379>
				<%		
			End if
			
			intCounter = intCounter + 1
			Laddername = ors.fields("laddername").value
			ActiveTeams = ors.fields("ActiveTeams").value
			AvailTeams = ors.fields("AvailTeams").value
			RestTeams = ors.fields("RestTeams").value
			TotalTeams = ors.fields("TotalTeams").value
			if TotalTeams > 0 Then
				intPercentActive = ActiveTeams / TotalTeams
			Else
				intPercentActive = 0
			End If
			%>
			<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 ALIGN=CENTER WIDTH=225 BGCOLOR="#444444">
			<TR><TD>
				<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 ALIGN=CENTER WIDTH=100%>
				<TR BGCOLOR="#000000">
					<TH COLSPAN=2 NOWRAP><%=Laddername%></TH>
				</TR>
				<TR BGCOLOR="<%=bgcone%>">
					<TD ALIGN=RIGHT WIDTH=175>Active Teams:</TD>
					<TD ALIGN=RIGHT WIDTH=50><%=activeteams%></TD>
				</TR>
				<TR BGCOLOR="<%=bgctwo%>">
					<TD ALIGN=RIGHT>Available Teams:</TD>
					<TD ALIGN=RIGHT><%=AvailTeams%></TD>
				</TR>
				<TR BGCOLOR="<%=bgcone%>">
					<TD ALIGN=RIGHT>Resting Teams:</TD>
					<TD ALIGN=RIGHT><%=restTeams%></TD>
				</TR>
				<TR BGCOLOR="<%=bgctwo%>">
					<TD ALIGN=RIGHT>Total Teams:</TD>
					<TD ALIGN=RIGHT><%=TotalTeams%></TD>
				</TR>
				<TR BGCOLOR="<%=bgcone%>">
					<TD ALIGN=RIGHT>Percent Active:</TD>
					<TD ALIGN=RIGHT><%=FormatPercent(intPercentActive, 2)%></TD>
				</TR>
				</TABLE>
			</TD></TR></TABLE>
			<%
			ors.movenext
		loop
	end if
	ors.close
If intCounter Mod 2 <> 0 Then
	%>
	</td>
	<td><img src="/images/spacer.gif" width="10" height="1"></td>
	<td width=379>&nbsp;
	<%
End If
%>
	</td>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	</tr>
	</table>
<%
Call Content2BoxEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>