<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Tournament Administration"

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

Dim X, tournamentid

If Not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Create Tournament")
%>
<table ALIGN=CENTER border=0 cellpadding=0 cellspacing=0 BGCOLOR="#444444">
<TR><TD align="center">
	
<table ALIGN=CENTER border=0 cellpadding=4 cellspacing=1 WIDTH=100%>
	<TR BGCOLOR="#000000"><TH colspan="6"><a href="createtourny.asp">Create a new tournament</a><br /></TH></TR>
	
	<TR BGCOLOR="#000000"><TH colspan="6">Choose a tournament</TH></TR>
	<TR BGCOLOR="#000000">
		<TH>Tournament Name</TH>
		<TH>Edit Divisions</TH>
		<TH>Edit Tournament</TH>
		<TH>Edit Content</TH>
		<TH>Seedings</TH>
	</TR>
	<%
	strSQL = "SELECT TournamentName, TournamentID, Locked, Signup FROM tbl_tournaments WHERE Active = 1 order by TournamentName"
	ors.open strsql, oconn
	if not(ors.eof and ors.bof) then
		Do While Not (oRS.EOF)
		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		%>
		<TR BGCOLOR=<%=bgc%>>
			<TD>
		  		<a href="/tournament/viewBracket.asp?tournament=<%=server.urlencode(ors.fields("TournamentName").value)%>"><%=oRS("TournamentName")%></a><br>
		  	</td>
			<TD>
		  		<a href="editDivisions.asp?tournament=<%=server.urlencode(ors.fields("TournamentName").value)%>">Edit Divisions</a><br>
		  	</td>
				<td><a href="EditTournament.asp?Tournament=<%=server.urlencode(ors.fields("TournamentName").value)%>">Edit Tournament</a></td>
				<td><a href="EditContent.asp?Tournament=<%=server.urlencode(ors.fields("TournamentName").value)%>">Edit Content</a></td>
			<TD align="center">
		  		<a href="editSeedings.asp?tournament=<%=server.urlencode(ors.fields("TournamentName").value)%>">Reseed</a><br>
		  	</td>
		</tr>
		<%
		oRS.MoveNext
		Loop
	end if
	oRS.Close
	%>
</table>		
</TD></TR>
</tABLE>
<br />
<br />
<table ALIGN=CENTER border=0 cellpadding=0 cellspacing=0 BGCOLOR="#444444">
<TR><TD align="center">
	
<table ALIGN=CENTER border=0 cellpadding=4 cellspacing=1 WIDTH=100%>
	<TR BGCOLOR="#000000"><TH colspan="6">Inactive Tournaments</TH></TR>
	<TR BGCOLOR="#000000">
		<TH>Tournament Name</TH>
		<TH>Edit Divisions</TH>
		<TH>Edit Tournament</TH>
		<TH>Edit Content</TH>
		<TH>Seedings</TH>
	</TR>
	<%
	strSQL = "SELECT TournamentName, TournamentID, Locked, Signup FROM tbl_tournaments WHERE Active = 0 order by TournamentName"
	ors.open strsql, oconn
	if not(ors.eof and ors.bof) then
		Do While Not (oRS.EOF)
		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		%>
		<TR BGCOLOR=<%=bgc%>>
			<TD>
		  		<a href="/tournament/viewBracket.asp?tournament=<%=server.urlencode(ors.fields("TournamentName").value)%>"><%=oRS("TournamentName")%></a><br>
		  	</td>
			<TD>
		  		<a href="editDivisions.asp?tournament=<%=server.urlencode(ors.fields("TournamentName").value)%>">Edit Divisions</a><br>
		  	</td>
				<td><a href="EditTournament.asp?Tournament=<%=server.urlencode(ors.fields("TournamentName").value)%>">Edit Tournament</a></td>
				<td><a href="EditContent.asp?Tournament=<%=server.urlencode(ors.fields("TournamentName").value)%>">Edit Content</a></td>
			<TD align="center">
		  		<a href="editSeedings.asp?tournament=<%=server.urlencode(ors.fields("TournamentName").value)%>">Reseed</a><br>
		  	</td>
		</tr>
		<%
		oRS.MoveNext
		Loop
	end if
	oRS.Close
	%>
</table>		
</TD></TR>
</tABLE>
<%
Call ContentEnd() 
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>