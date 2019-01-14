<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Create Tournament"

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

Dim strTournamentName, intTournamentID
strTournamentName = Request.QueryString("tournament")
strSQL = "SELECT TournamentID FROM tbl_tournaments WHERE TournamentName = '" & CheckSTring(strTournamentName) & "'"
oRs.Open strSQL, oConn
If Not(ORs.Eof and Ors.BOF) Then
	intTournamentID = oRs.Fields("TournamentID").Value
End If
oRs.NextRecordSet
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("Edit " & Server.HTMLEncode(strTournamentName) & " Divisions")
%>
<table ALIGN=CENTER border=0 cellpadding=0 cellspacing=0 BGCOLOR="#444444">
<TR><TD align="center">
	
<table ALIGN=CENTER border=0 cellpadding=2 cellspacing=1 WIDTH=100%>
<form method="post" action="savetournament.asp">
<input type="hidden" name="SaveType" id="SaveType" value="DivisionNames" />
<%
Dim intCounter
intCounter = 1
strSQL = "SELECT DivisionName, DivisionID, TDivisionID FROM tbl_tdivisions WHERE TournamentID = '" & intTournamentID & "'"
oRs.Open strSQL, oConn
If Not(oRS.EOF AND oRs.BOF) THen
	Do While Not (oRs.EOF)
		%>
		<tr>
			<td nowrap="nowrap" bgcolor="<%=bgcone%>">Division <%=intCounter%> (<%=oRs.Fields("TDivisionID").Value%>)</td>
			<td bgcolor="<%=bgctwo%>"><input type="hidden" name="hdnDivisionID" id="hdnDivisionID" value="<%=oRs.Fields("TDivisionID").Value%>" />
				<input type="text" name="txtDivisionName" id="txtDivisionName" value="<%=oRs.Fields("DivisionName").Value%>" size="40" />
			</td>
		</tr>
		<%
		intCounter = intCounter + 1
		oRs.MoveNext
	Loop
End If
oRs.NextRecordset
%>
	<tr>
		<td width="100%" align="center" colspan="2" bgcolor="#000000"><input type="Submit" name="submit" value="Save Names"></td>
	</tr>
</form>
</table></td></tr>
</table>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>