<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Save Demo File"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

Dim numMatches, totalc, PlayerHandle
Dim c, opp, bgcc, MatchDate, bgcc2
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%

If Not(Session("LoggedIn")) then 
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Set oRS2 = Nothing
	Response.Redirect("/errorpage.asp?error=2")
end if

If Request.QueryString("n") = "" Then 
	numMatches = 4
Else
	numMatches = "0"
End If

bgc=bgcone
	
totalc = 0
	
Call ContentStart("Add Demo")
%>
	<table width=760 border="0" cellspacing="0" cellpadding="2">
	<tr><td>
	<table width=500 align=center border=0 cellpadding=2 cellspacing=0>
		<tr bgcolor=<%=bgc%>>
			<td>Select the match that the demo is of...</td>
			<td align=right>Step: <b>1</b> 2 3</td>
		</tr>
	</table>
	<BR>
	<center>
		<%
		If numMatches = 4 Then 
			Response.Write "<b>Showing Recent Matches</b>"
		Else
			Response.Write "<b>Showing All Matches </b>"
		End If
		%>
		
	</center>
		<%
		PlayerHandle = Session("uName")
		strSQL = "SELECT * FROM vPlayerTeams WHERE PlayerHandle='" & CheckString(PlayerHandle) & "'"
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRS.BOF) Then
			Do While Not oRS.EOF
				%>
				<table width=500 border=0 align=center cellpadding=0 cellspacing=0 BGCOLOR="#444444">
				<TR><TD>
				<table width=100% border=0 align=center cellpadding=2 cellspacing=1 BGCOLOR="#444444">
				
				<tr BGCOLOR="#000000">
					<TH colspan=3><%=oRS("LadderName")%></TH>
				</tr>
				<tr BGCOLOR="#000000">
					<tH>&nbsp;</th>
					<th>Date</th>
					<th>Opponent</th>
				</tr>			
				<%
				bgcc2 = bgcone
				strSQL = "SELECT WinnerName, MatchWinnerID, LoserName, MatchDate, MatchID FROM vHistory WHERE (MatchWinnerID=" & oRS("TLLinkID") & " OR MatchLoserID=" & oRS("TLLinkID") & ") AND MatchLadderID=" & oRS("LadderID") & " ORDER BY MatchDate DESC"
				oRs2.Open strSQL, oConn
				If Not(oRS2.EOF AND oRS2.BOF) Then 
					Do While Not oRS2.EOF
						if totalc = numMatches and numMatches <> "0" Then
							exit do
						end if
						
						If oRS2("MatchWinnerID") = oRS("TLLinkID") Then
							opp = oRS2("LoserName")
						Else
							Opp = oRs2("WinnerName")
						End If
						'get opponent's name
						MatchDate = month(oRS2("MatchDate")) & "/" & day(oRS2("MatchDate")) & "/" & Year(oRS2("MatchDate"))
						totalc = totalc + 1
						%>
							<tr bgcolor=<%=bgcc2%>>
								<td align=center><a href="addDemo2.asp?MatchID=<%=oRS2("MatchID")%>">Select</a></td>
								<td><%=MatchDate%></td>
								<td><%=opp%></td>
							</tr>
						<%
						If bgcc2 = bgcone Then
							bgcc2 = bgctwo
						else
							bgcc2 = bgcone
						End If
						oRS2.MoveNext
					Loop
				End If
				oRS2.NextRecordSet
				totalc = 0
				If numMatches = 4 Then 
					Response.Write "<tr bgcolor=000000><td colspan=3 align=right><a href=""addDemo.asp?n=0"">Show All Matches</a></td></tr>"
				End If
				oRS.MoveNext
				%>
				</TABLE></TD></TR></TABLE><BR><BR>
				<%
			Loop
		End If
		oRS.Close
		%>
	</td></tr>
	</table>            
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing
%>