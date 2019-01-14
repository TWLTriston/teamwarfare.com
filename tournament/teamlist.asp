<script language="javascript" type="text/javascript">
<!--
function fConfirmRemove(intTMLinkID, strTeamName) {
	if (confirm("Are you absolutely certain you want to kick " + strTeamName + " out of the tournament?")) {
		window.location = "/saveitem.asp?SaveType=TournamentRemove&TMLinkID=" + intTMLinkID + "&Tournament=<%=Server.URLEncode(strTournamentName & "")%>";
	}
}
//-->
</script>
<table border="0" cellspacing="0" cellpadding="4" width="90%" align="center">
<%
strSQL = "SELECT t.TeamName, t.TeamTag, l.TMLinkID FROM lnk_t_m l INNER JOIN tbl_teams t ON t.TeamID = l.TeamID  "
strSQL = strSQL & " WHERE TournamentID = '" & intTournamentID & "'"
strSQL = strSQL & " ORDER BY TeamName ASC"
ors.open strsql, oconn
if not(ors.eof and ors.bof) then
	do while not(ors.eof)
		%>
		<tr>
			<td><a href="/viewteam.asp?team=<%=Server.URLEncode(oRs.fields("TeamName").Value & "")%>"><%=Server.HTMLEncode(oRs.Fields("teamName").Value & " - " & ors.fields("teamtag").value)%></a> <%
			If bSysAdmin THen %>
				-- <a href="javascript:fConfirmRemove(<%=oRS.Fields("TMLinkID").Value%>, '<%=Replace(Server.HTMLEncode(oRs.Fields("TeamName").Value & ""), "'", "\'")%>');">remove from tournament</a>
				<%
			End If
			%>
		</tr>
		<%
		ors.movenext
	loop
else
	%>
	<tr>
		<td>No teams have signed up yet.</td>
	</tr>
	<%	
end if
ors.nextrecordset
%>
</table>
