<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Create League"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2= Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

dim intCounter
dim intLeagueID
dim intDivCount

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Division Config")
%>
<form name="frmDivisionConfig" id="frmDivisionConfig" method="post" action="saveLeague.asp">
<table border="0" cellspacing="0" cellpadding="0">
<%
	intLeagueID = request.querystring("intLeagueID")
	strSQL = "select ConferenceName, LeagueConferenceID from tbl_league_conferences where LeagueID=" & intLeagueID
	oRS.open strSQL, oConn
	if not (oRS.bof and oRS.eof) then
		do while not oRS.eof
			response.write "<tr><td colspan=""2""><b>Conference: " & oRS.fields("ConferenceName").value & "</b></td></tr>"
			strSQL = "select LeagueDivisionID, DivisionName from tbl_league_divisions where LeagueID=" & intLeagueID & " and LeagueConferenceID=" & oRS.fields("LeagueConferenceID").value
			oRS2.open strSQL, oConn
			if not (oRs2.eof and oRs2.bof) then
				do while not oRs2.eof
					response.write "<tr><td>Division " & oRs2.fields("DivisionName").value & " Name: </td>" & vbCrlF
					response.write "<td><input type=""text"" name=""txtDivName"" value=""" & oRs2.fields("DivisionName").value & """ />" & vbCrlF
					response.write "<input type=""hidden"" name=""hdnDivID"" id=""hdnDivID"" value=""" & oRs2.fields("LeagueDivisionID").value & """ /> </td></tr>" & vbCrlF
					oRs2.movenext
				loop
			end if
			oRs2.nextRecordSet
			response.write "<tr><td colspan=2>&nbsp;</td></tr>"
			oRS.movenext
		loop
	end if
	oRS.close	
%>
	<tr>
		<td><input type="hidden" name="hdnLeagueID" value="<%=intLeagueID%>" />
		<input type="hidden" name="SaveType" id = "saveType" value="Divisions" />
		<input type="submit" value="Save" /></td></tr>
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