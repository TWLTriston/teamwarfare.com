<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Referral Tracking"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin) then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim intSpot
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Trackign")

strSQL = "SELECT TrackURLName, DistinctIPs = COUNT(DISTINCT TrackURLIP), DistinctPlayers = COUNT(DISTINCT TrackURLPlayer), Clicks = Count(TrackURLID) FROM tbl_track_url GROUP BY TrackURLName"
oRs.Open strSQL, oConn
If Not(oRS.EOF AND oRS.BOF) Then
	%>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER>
	<TR><TD>
		<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2>
		<TR BGCOLOR="#000000">
			<TH>URL</TH>
			<TH>Clicks</TH>
			<TH>Unique IPs</TH>
			<TH>Unique Accounts</TH>
		</TR>
		<%
		bgc=bgcone
		Do While Not(oRs.EOF)
			Response.Write "<TR BGCOLOR=" & bgc & ">"
			Response.Write "<TD>" & oRs.Fields("TrackURLName").Value & "</TD>"
			Response.Write "<TD ALIGN=RIGHT>" & oRs.Fields("Clicks").Value & "</TD>"
			Response.Write "<TD ALIGN=RIGHT>" & oRs.Fields("DistinctIPs").Value & "</TD>"
			Response.Write "<TD ALIGN=RIGHT>" & oRs.Fields("DistinctPlayers").Value & "</TD>"
			Response.Write "</TR>"
			
			If bgc = bgcOne then
				bgc = bgctwo
			Else
				bgc = bgcone
			End If
			oRs.MoveNext
		Loop
		%>
		</TABLE>
	</TD></TR>
	</TABLE>
	<%
End If
oRs.Close
strSQL = "SELECT TrackImageURL, DistinctIPs = COUNT(DISTINCT TrackImageIP), DistinctPlayers = COUNT(DISTINCT TrackImagePlayer), DisplayCount = COUNT(TrackImageID) FROM tbl_track_image GROUP BY TrackImageURL"
oRs.Open strSQL, oConn
If Not(oRS.EOF AND oRS.BOF) Then
	%>
	<br />
	<br />
	
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER>
	<TR><TD>
		<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2>
		<TR BGCOLOR="#000000">
			<TH>Image Name</TH>
			<TH>Display Count</TH>
			<TH>Unique IPs</TH>
			<TH>Unique Accounts</TH>
		</TR>
		<%
		bgc=bgcone
		Do While Not(oRs.EOF)
			Response.Write "<TR BGCOLOR=" & bgc & ">"
			Response.Write "<TD>" & oRs.Fields("TrackImageURL").Value & "</TD>"
			Response.Write "<TD ALIGN=RIGHT>" & oRs.Fields("DisplayCount").Value & "</TD>"
			Response.Write "<TD ALIGN=RIGHT>" & oRs.Fields("DistinctIPs").Value & "</TD>"
			Response.Write "<TD ALIGN=RIGHT>" & oRs.Fields("DistinctPlayers").Value & "</TD>"
			Response.Write "</TR>"
			
			If bgc = bgcOne then
				bgc = bgctwo
			Else
				bgc = bgcone
			End If
			oRs.MoveNext
		Loop
		%>
		</TABLE>
	</TD></TR>
	</TABLE>
	<%
End If
oRs.Close
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>