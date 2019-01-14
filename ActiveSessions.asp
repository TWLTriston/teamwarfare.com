<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Active Sessions"

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

If Not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Active Sessions") %>
<table width=760 border=0 cellspacing=0 cellpadding=0 align=center BGCOLOR="#444444">
<TR><TD>
<table width=100% border=0 cellspacing=1 cellpadding=2 align=center>
<tr BGCOLOR="#000000">
	<TH>Player</th>
	<th>IP Address</th>
	<TH>First Page</th>
	<TH>Last Page</th>
	<TH>Online</th>
	<TH>Last Ping</th>
</tr>
	<%
	Dim strMinutes, intMemberCount, intVisitorCount	
	Dim strURL, intForums 
	intForums  = 0
	intMemberCount = 0
	intVisitorCount	= 0
	strSQL = "SELECT * FROM tbl_active_sessions ORDER BY LastPingTime DESC"
	oRS.Open strSQL, oConn
	bgc = bgctwo
	if not (ors.eof and ors.bof) then
		do while not ors.EOF 
			Response.write "<tr bgcolor=" & bgcone & ">"
			If Len(ors.Fields("PlayerHandle").Value) > 0 Then
				Response.write "<td><a href=""/viewplayer.asp?player=" & Server.URLEncode(ors.Fields("PlayerHandle").Value & "") & """>" & Server.HTMLEncode(ors.Fields("PlayerHandle").Value & "") & "</A></TD>"
				intMemberCount = intMemberCount + 1
			Else
				intVisitorCount = intVisitorCount + 1
				Response.Write "<td>Visitor</td>"
			End If
		  Response.Write "<td>" & oRs.Fields("IPAddress").value & "</td>"
			strURL = ors.Fields("FirstPageView").Value
			If Len(strURL) > 40 Then
				strURL = Left(strURL, 20) & "..." & Right(strURL, 10)
			End If
			Response.write "<td><a href=""" & ors.Fields("FirstPageView").Value & """>" & Server.HTMLEncode(strURL & "") & "</a></TD>"
			
			strURL = ors.Fields("LastPageView").Value
			If Len(strURL) > 40 Then
				strURL = Left(strURL, 20) & "..." & Right(strURL, 10)
			End If
			Response.write "<td><a href=""" & ors.Fields("LastPageView").Value & """>" & Server.HTMLEncode(strURL & "") & "</a></TD>"
			If InStr(1, ors.Fields("LastPageView").Value, "forums", 1) > 0 Then
				intForums = intForums + 1
			End If
			strMinutes = DateDiff("n", ors.Fields("FirstPingTime").Value, ors.Fields("LastPingTime").Value) 
			Response.write "<td align=""center"">" & strMinutes & "</td>"
			strMinutes = DateDiff("n", ors.Fields("LastPingTime").Value, Now()) 
			Response.write "<td align=""center"">" & strMinutes & "</td>"
			ors.MoveNext 
			if bgc = bgcone then
				bgc = bgctwo
			else
				bgc = bgcone
			end if 
		loop
	end if
	ors.close
	%>
</table></TD></TR>
</TABLE>
<table align="left" border=0 cellspacing=0 cellpadding=16>
<TR>
	<TD>
	<%
	
	Response.Write "<B>In the Forums:</B> " & intForums  & "<BR>"
	Response.Write "<B>Members:</B> " & intMemberCount & "<BR>"
	Response.Write "<B>Visitors:</B> " & intVisitorCount & "<BR>"
	Response.Write "<B>Total users:</b> " & intVisitorCount + intMemberCount & "<BR>"
	%>
	</TD>
</TR>
</table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

