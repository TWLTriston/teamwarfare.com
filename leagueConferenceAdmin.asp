<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Conference Setup"

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

dim intGameID

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Conference Setup")

intLeagueID = request.querystring("intLeagueID")
strSQL = "select count(*) from tbl_league_conferences where LeagueID=" & intLeagueID
oRS.open strSQL, oConn
intNumConferences = oRS.fields(0).value
ors.close
%>
<form name="frmConferenceAdmin" id="frmConferenceAdmin" method=post action="saveLeague.asp">
	<table border="0" cellspacing="0" cellpadding="0">
		<%
			for intCount = 1 to intNumConferences
				response.write "<tr><td>Conference " & intCounter & " Name: </td>"
				response.write "<td><input name=""txtConferenceName_" & intCounter & """ /></td></tr>"

				response.write "<tr><td>Divisions for Conference " & intCounter & ": </td>"
				response.write "<td><input name=""txtConferenceDivCount_" & intCounter & """ /></td></tr>"
				
				response.write "<tr><td>&nbsp;</td></tr>"
			next
		%>
		<tr>
			<td colspan="2" align="center"><input type="hidden" name="intLeagueID" value="<%=intLeagueID%>" />
			<input type="hidden" name="SaveType" value="ConferenceSettings" />
			<input type="submit" value="Save Conference Data" /></td></tr>
	</table>
</form>

<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>