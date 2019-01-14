<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: News Archive"

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

Dim strHeadline, strContent, intMonth, intYear
intMonth = Request.QueryString("month")
intYear = Request.QueryString("year")

If intMonth = "" Then
	intMonth = Month(Now())
	intYear = Year(Now())
End If
If intYear = "" Then
	intYear = Year(Now())
End If
Dim intLocation, intYearCount, intMonthCount
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("News Archive") %>
<table width=500 border=0 cellspacing=0 cellpadding=0 align=center BGCOLOR="#444444">
<TR><TD>
<table width=100% border=0 cellspacing=1 cellpadding=2 align=center>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=2>Choose Month</TH>
	</TR>
	<TR BGCOLOR=<%=bgcOne%>>
		<TD>
	<%
	intLocation = 0
	For intYearCount = 2001 To Year(Now())
		For intMonthCount = 1 to 12
			If (intYearCount = Year(Now()) AND intMonthCount <= Month(now())) OR (intYearCount < Year(Now())) Then
				If intLocation Mod 2 = 1 Then
					Response.Write "</TD><TD>"
				ElseIf intLocation > 0 Then
					Response.Write "</TD></TR><TR BGCOLOR=" & bgcOne & "><TD>"
				End If
				intLocation = intLocation + 1
				Response.Write "<a href=""/newsarchive.asp?month=" & intMonthCount & "&year=" & intYearCount & """>" & MonthName(intMonthCount) & " " & intYearCount & "</A>"
			End If
		Next
	Next
	If intLocation Mod 2 = 1 Then
		Response.Write "</TD><TD>&nbsp;"
	End If
	Response.Write "</TD></TR>"
	%>
	
</TABLE>
</TD></TR>
</TABLE>
<BR><BR>

<%
strSQL = "select * from tbl_News WHERE NewsDate >= '" & intMonth & "/1/" & intYear & "' AND NewsDate < DateAdd(m, 1, '" & intMonth & "/1/" & intYear & "') ORDER BY NewsID ASC"
'Response.Write strSQL
oRs.Open strSQL, oConn
bgc = bgctwo
if not (ors.eof and ors.bof) then
	do while not ors.EOF 
		%>
		<table width=760 border=0 cellspacing=0 cellpadding=0 align=center BGCOLOR="#444444">
		<TR><TD>
		<table width=100% border=0 cellspacing=1 cellpadding=2 align=center>
			<tr BGCOLOR=#000000><td><b><%=Server.HTMLEncode(ors.Fields("NewsHeadline").Value)%></b> (<%=ors.Fields("NewsDate").Value%> by <%= Server.HTMLEncode(ors.Fields("NewsAuthor").Value)%>)</td></tr>
		<tr bgcolor=<%=bgc%>><td><%=ors.fields("NewsContent").value%></td></tr>
		</TABLE></TD></TR>
		</TABLE>
		<BR>
		<%
		ors.MoveNext 
		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if 
	loop
end if
oRS.NextRecordset 

Call ContentEnd() 
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

