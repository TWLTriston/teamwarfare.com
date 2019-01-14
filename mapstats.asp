<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Map Statistics"

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

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("")
Dim strLadderName
Dim intMaps, strMapConfiguration
Dim intDef, intAtt, intRan, intTeamTotal, intMapTotal
Dim intDefTotal, intAttTotal, intRanTotal, intTeamTotalTotal, intMapTotalTotal

strLadderName = Request.QueryString("Ladder")
strSQL = "EXECUTE GetMapStats @LadderName = '" & CheckString(strLadderName) & "'"
oRs.Open strSQL, oConn
If (oRs.State <> 1) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
If (oRs.EOF AND oRs.BOF) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
If oRs.Fields("Error").Value = "1" Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=30"
End If
Dim i, intCols
strMapConfiguration = oRs.Fields("MapConfiguration").Value
intMaps = cInt(oRs.Fields("Maps").Value)
strLaddername = oRs.Fields("LadderName").Value
''Response.Write intMaps
Set oRs = oRs.NextRecordSet
intCols = intMaps
If InStr(strMapConfiguration, "A") > 0 OR InStr(strMapConfiguration, "C") > 0 Then
	intCols = intCols + 1
End If
If InStr(strMapConfiguration, "D") > 0 Then
	intCols = intCols + 1
End If
If InStr(strMapConfiguration, "R") > 0 Then
	intCols = intCols + 1
End If

%>
<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" align="center">
<tr>
	<td>
		<table border="0" cellspacing="1" cellpadding="4">
		<tr>
			<th bgcolor="#000000" colspan="<%=intCols + 2%>" valign="bottom"><%=strLaddername%> Map Statistics</th>
		</tr>
		<tr>
			<th bgcolor="#000000" rowspan="2" valign="bottom">Map Name</th>
			<th bgcolor="#000000" colspan="<%=intCols+1%>">Choosen</th>
		</tr>
		<tr>
			<% For i = 1 to intMaps %>
			<th bgcolor="#000000">As Map <%=i%></th>
			<% Next %>
			<% If InStr(strMapConfiguration, "A") > 0 OR InStr(strMapConfiguration, "C") > 0 Then %>
			<th bgcolor="#000000">By Attacker</th>
			<% End If %>
			<% If InStr(strMapConfiguration, "D") > 0 Then %>
			<th bgcolor="#000000">By Defender</th>
			<% End If %>
			<% If InStr(strMapConfiguration, "R") > 0 Then %>
			<th bgcolor="#000000">Randomly</th>
			<% End If %>
			<th bgcolor="#000000">Total</th>
		</tr>
		<%
		Dim chrMapChoosen, arrTotals(13)
		arrTotals(0) = 0
		arrTotals(1) = 0
		arrTotals(2) = 0
		arrTotals(3) = 0
		arrTotals(4) = 0
		arrTotals(5) = 0
		arrTotals(6) = 0
		arrTotals(7) = 0
		arrTotals(8) = 0
		arrTotals(9) = 0
		arrTotals(10) = 0
		arrTotals(11) = 0
		arrTotals(12) = 0
		
		Do While Not(oRS.EOF)
			arrTotals(9) = 0
			arrTotals(10) = 0
			arrTotals(11) = 0
			arrTotals(12) = 0
			%>
			<tr>
				<td bgcolor="<%=bgcone%>"><%=oRs.Fields("MapName").Value%></td>
				<% 
				For i = 1 to intMaps 
					%>
					<td bgcolor="<%=bgctwo%>" align="right"><%=oRs.Fields("AsMap" & i).Value%></td>
					<% 
					chrMapChoosen = mid(strMapConfiguration, i, 1)
					arrTotals(0) = arrTotals(0) + oRs.Fields("AsMap" & i).Value
					arrTotals(9) = arrTotals(9) + oRs.Fields("AsMap" & i).Value
					arrTotals(i) = arrTotals(i) + oRs.Fields("AsMap" & i).Value
					If chrMapChoosen = "R" Then
						arrTotals(6) = arrTotals(6) + oRs.Fields("AsMap" & i).Value
						arrTotals(10) = arrTotals(10) + oRs.Fields("AsMap" & i).Value
					ElseIf chrMapChoosen = "A" OR  chrMapChoosen = "C" Then
						arrTotals(7) = arrTotals(7) + oRs.Fields("AsMap" & i).Value
						arrTotals(11) = arrTotals(11) + oRs.Fields("AsMap" & i).Value
					ElseIf chrMapChoosen = "D" Then
						arrTotals(8) = arrTotals(8) + oRs.Fields("AsMap" & i).Value
						arrTotals(12) = arrTotals(12) + oRs.Fields("AsMap" & i).Value
					End If
				Next 
				%>
				<% If InStr(strMapConfiguration, "A") > 0 OR InStr(strMapConfiguration, "C") > 0 Then %>
				<td bgcolor="<%=bgcone%>" align="right"><%=arrTotals(11)%></td>
				<% End If %>
				<% If InStr(strMapConfiguration, "D") > 0 Then %>
				<td bgcolor="<%=bgcone%>" align="right"><%=arrTotals(12)%></td>
				<% End If %>
				<% If InStr(strMapConfiguration, "R") > 0 Then %>
				<td bgcolor="<%=bgcone%>" align="right"><%=arrTotals(10)%></td>
				<% End If %>
				<td bgcolor="<%=bgcone%>" align="right"><%=arrTotals(9)%></td>
			</tr>
			<%		
			oRs.MoveNext
		Loop
		%>
		<tr>
			<td bgcolor="#000000" colspan="<%=intCols + 1%>" align="right">Total:</td>
			<td bgcolor="#000000" align="right"><%=arrTotals(0)%></td>
		</tr>
		</table>
	</td>
</tr>
</table>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

