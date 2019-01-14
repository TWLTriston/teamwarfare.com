	<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "teamwarfare.com"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Call ContentStart("")
%>
<table border="0" cellspacing="0" cellpadding="0" width="97%" align="center" bgcolor="#444444">
<tr>
	<td>
	<table border="0" cellspacing="1" cellpadding="4" width="100%">
	<tr>
		<th bgcolor="#000000">Response</th>
		<th bgcolor="#000000">Original Results</th>
		<th bgcolor="#000000">Actual Results</th>
		<th bgcolor="#000000">Founders Only</th>
	</tr>
	<tr>
		<th bgcolor="#000000" colspan="4">1. Should creativity be awarded in pub mode only?</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">No, pub should match competitive play. Any changes made should affect both pub and tourny mode.</td>
		<td bgcolor="<%=bgcone%>" align="right">60</td>
		<td bgcolor="<%=bgcone%>" align="right">37</td>
		<td bgcolor="<%=bgcone%>" align="right">5</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>">Yes, encourage passing and style in pubs.</td>
		<td bgcolor="<%=bgctwo%>" align="right">23</td>
		<td bgcolor="<%=bgctwo%>" align="right">7</td>
		<td bgcolor="<%=bgctwo%>" align="right">1</td>
	</tr>
	<tr>
		<th bgcolor="#000000" colspan="4">2. Should the creativity system be changed?</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>">Nope, it's fine the way it is.</td>
		<td bgcolor="<%=bgctwo%>" align="right">20</td>
		<td bgcolor="<%=bgctwo%>" align="right">12</td>
		<td bgcolor="<%=bgctwo%>" align="right">2</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Yes, I feel there is something that should be changed.</td>
		<td bgcolor="<%=bgcone%>" align="right">68</td>
		<td bgcolor="<%=bgcone%>" align="right">35</td>
		<td bgcolor="<%=bgcone%>" align="right">4</td>
	</tr>
	<tr>
		<th bgcolor="#000000" colspan="4">3. How would you like it changed?</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>">Change it, give less reward for the different levels (for example) 5-10-25 points.</td>
		<td bgcolor="<%=bgctwo%>" align="right">15</td>
		<td bgcolor="<%=bgctwo%>" align="right">7</td>
		<td bgcolor="<%=bgctwo%>" align="right">0</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Disable it completely, TR2 is about possession and scoring.	</td>
		<td bgcolor="<%=bgcone%>" align="right">25</td>
		<td bgcolor="<%=bgcone%>" align="right">7</td>
		<td bgcolor="<%=bgctwo%>" align="right">3</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>">Fix the variance bug, and all will be well.</td>
		<td bgcolor="<%=bgctwo%>" align="right">19</td>
		<td bgcolor="<%=bgctwo%>" align="right">11</td>
		<td bgcolor="<%=bgctwo%>" align="right">0</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">Leave it alone.</td>
		<td bgcolor="<%=bgcone%>" align="right">12</td>
		<td bgcolor="<%=bgcone%>" align="right">10</td>
		<td bgcolor="<%=bgctwo%>" align="right">2</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>">Other... should be changed, no good solutions here.	</td>
		<td bgcolor="<%=bgctwo%>" align="right">14</td>
		<td bgcolor="<%=bgctwo%>" align="right">4</td>
		<td bgcolor="<%=bgctwo%>" align="right">1</td>
	</tr>
	<tr>
		<th bgcolor="#000000" colspan="4">4. Increase the maximum jackpot amount?</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">No, it's fine.</td>
		<td bgcolor="<%=bgcone%>" align="right">66</td>
		<td bgcolor="<%=bgcone%>" align="right">44</td>
		<td bgcolor="<%=bgcone%>" align="right">6</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>">Yes, more would be better.</td>
		<td bgcolor="<%=bgctwo%>" align="right">16</td>
		<td bgcolor="<%=bgctwo%>" align="right">3</td>
		<td bgcolor="<%=bgcone%>" align="right">0</td>
	</tr>
	<tr>
		<th bgcolor="#000000" colspan="4">5. How do you like GPA so far?</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">It's a step in the right direction, needs some tweaking.</td>
		<td bgcolor="<%=bgcone%>" align="right">35</td>
		<td bgcolor="<%=bgcone%>" align="right">19</td>
		<td bgcolor="<%=bgcone%>" align="right">3</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>">It's great, requires more skill to shoot.</td>
		<td bgcolor="<%=bgctwo%>" align="right">30</td>
		<td bgcolor="<%=bgctwo%>" align="right">24</td>
		<td bgcolor="<%=bgctwo%>" align="right">3</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>">It's horrible, revert back.</td>
		<td bgcolor="<%=bgcone%>" align="right">5</td>
		<td bgcolor="<%=bgcone%>" align="right">1</td>
		<td bgcolor="<%=bgcone%>" align="right">0</td>
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
Set oRS = Nothing
Set oConn = Nothing
%>