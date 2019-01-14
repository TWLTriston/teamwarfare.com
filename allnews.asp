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

Dim intPendingCnt, intHistoryCnt, strCurrDate
Dim strWinnerTag, strWinnerName
Dim strLoserTag, strLoserName
Dim strLadderName, strLadderAbbr
Dim strDefenderName, strDefenderTag
Dim strAttackerName, strAttackerTag
Dim strMatchTime

Dim intNewsID, intArticles, intNewsCnt
intArticles = 5
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
	strSQL = "select top 20 n.*, p.PlayerEmail from tbl_News n, tbl_players p WHERE p.playerhandle = n.NewsAuthor order by NewsID desc"
	oRS.Open strSQL, oConn
	intNewsID = 0
	intNewsCnt = 1
	If Not(oRS.EOF and oRS.BOF) Then
		Do While Not(oRS.EOF) AND (intNewsCnt <= intArticles)
			If intNewsCnt = 1Then
				strCurrDate = ors.fields("NewsDate").value
				Call ContentNewsStart(weekdayname(weekday(strCurrDate)) & ", " & monthname(month(strCurrDate)) & " " & day(strCurrDate))
			ElseIf (strCurrDate <> oRS.Fields("NewsDate").Value) then
				Call ContentNewsEnd()
				strCurrDate = ors.fields("NewsDate").value
				Call ContentNewsStart(weekdayname(weekday(strCurrDate)) & ", " & monthname(month(strCurrDate)) & " " & day(strCurrDate))
			End If
			%>
			<tr><td>
				<table width="100%" align=center border=0 cellpadding="0">
				   <tr valign="top"> 
				    <td class="newsheader"><%=Server.HTMLEncode (ors.fields("NewsHeadLine").value)%></td>
				    <td ALIGN=RIGHT>Written by <a href="mailto:<%=ors.fields("PlayerEmail").value%>"><%=Server.HTMLEncode(ors.fields("NewsAuthor").value)%><br>
				        </a><%=formatdatetime(strCurrDate, 4)%>
				    </td>
				  </tr>
				  <TR>
					<TD COLSPAN=2><%=ors.fields("NewsContent").value%></td>
				  </tr>
				</table>
			</td></tr>
			<%
			intNewsCnt = intNewsCnt + 1
			oRS.MoveNext 
		Loop
	End If
	oRS.Close
	Call ContentNewsEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oRS = Nothing
Set oConn = Nothing
%>