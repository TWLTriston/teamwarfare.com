<% Option Explicit %>
<%
Response.Buffer = True
Server.ScriptTimeout = 900
Dim strPageTitle

strPageTitle = "TWL: Active Sessions Graph"

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

If Not(bSysAdmin) then
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Last 7 Days of Session Activity Averages")
%>
<table border="0" cellspacing="0" cellpadding="0" width="700" align="center">
<tr>
	<td width="75"><img src="/images/spacer.gif" alt="" border="0" height="1" width="75" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
</tr>
<tr>
	<td>Time</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>0</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>50</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>100</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>150</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>200</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>250</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>300</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>350</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>400</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>450</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>500</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>550</td>
</tr>

<%
Sub ShowIt()
	Dim strDate
	
	'strSQL = "SELECT TimeOfDay, Visitors, Members FROM tbl_active_sessions_stat ORDER BY TimeOfDay ASC "
	strSQL = "EXECUTE ActiveSessionsReport '" & Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()) & "'"
	'Response.Write strSQL
	oRs.Open strSQL, oConn
	If Not(oRS.EOF AND oRs.BOF) Then
		Do While Minute(oRs.Fields("TimeOfDay").Value) <> "0" AND NOT(oRs.EOF)
			oRs.MoveNext
		Loop
		If Not(oRs.EOF) Then
			Do While Not(oRs.EOF)
				Response.Write "<tr>" & vbCrLf
				If Minute(oRs.Fields("TimeOfDay").Value) = "0" Then
					Response.Write "<td align=""left"" rowspan=""14"">" & FormatDateTime(oRs.Fields("TimeOfDay").Value, 4) & "</td>"
				End if
				Response.Write "<td colspan=""28"" height=""1""><img src=""/ballot/images/bar.gif"" height=""1"" width=""" & (oRs.Fields("Visitors").Value + oRs.Fields("Members").Value) & """ alt=""" & FormatDateTime(oRs.Fields("TimeOfDay").Value, 3) & """></td>"
				Response.Write "</tr>"
				If Minute(oRs.Fields("TimeOfDay").Value) = "30" Or Minute(oRs.Fields("TimeOfDay").Value) = "0" Then
				%>
				<tr height="100%">
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
					<td width="1" bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
					<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
				</tr>
				<%
				End if
				oRs.MoveNext
			Loop
		End If
	End if
	oRs.NextRecordSet
End Sub
Call ShowIt()
Call ShowIt()
		'Response.Write "<TD BGCOLOR=#000000><img src=""/ballot/images/bar.gif"" height=10 width=" & (fix(pct) + 1) * 4 & "> &nbsp;&nbsp;" & formatnumber(pct,2,-1) & "%</TD>"
		%>
</table>

<table border="0" cellspacing="0" cellpadding="0" width="700" align="center">
<tr>
	<td>&nbsp;</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>0</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>50</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>100</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>150</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>200</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>250</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>300</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>350</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>400</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>450</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>500</td>
	<td bgcolor="#ffffff"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td>550</td>
</tr>
<tr>
	<td width="75"><img src="/images/spacer.gif" alt="" border="0" height="1" width="75" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
	<td width="1"><img src="/images/spacer.gif" alt="" border="0" height="1" width="1" /></td>
	<td width="49"><img src="/images/spacer.gif" alt="" border="0" height="1" width="49" /></td>
</tr>
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