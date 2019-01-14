<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Ballot Results"

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

Dim qRS, rRS, vRS, q, BallotName
Dim responses(6)
Dim rvalue(6)
Dim i, pct, totalresponses, intCols

If bSysAdmin Then
	intCols = 7
Else
	intCols = 6
End If
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Voting Booth")
%>
<TABLE BORDER=0 CElLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" width="97%">
<TR><TD>
<TABLE BORDER=0 CElLSPACING=1 CELLPADDING=2 width="100%">
<%
If Request.QueryString("error") = 1 Then
	Response.Write "<TR BGCOLOR=""#000000""><TD ALIGN=CENTER COLSPAN=" & intCols & "><B><FONT COLOR=""#FF0000"">You have already voted in that ballot.</FONT></B></TD></TR>"
End If
If Request.QueryString("error") = 2 Then
	Response.Write "<TR BGCOLOR=""#000000""><TD ALIGN=CENTER COLSPAN=" & intCols & "><B><FONT COLOR=""#FF0000"">The polls are closed on that ballot.</FONT></B></TD></TR>"
End If
If Request.QueryString("error") = 3 Then
	Response.Write "<TR BGCOLOR=""#000000""><TD ALIGN=CENTER COLSPAN=" & intCols & "><B><FONT COLOR=""#FF0000"">Ballot is for founders of ladder participants only.</FONT></B></TD></TR>"
End If
If Request.QueryString("error") = 4 Then
	Response.Write "<TR BGCOLOR=""#000000""><TD ALIGN=CENTER COLSPAN=" & intCols & "><B><FONT COLOR=""#FF0000"">Ballot is for founders of league participants only.</FONT></B></TD></TR>"
End If

%>
<TR BGCOLOR="#000000">
	<TH COLSPAN=<%=intCols%>>Choose a ballot</TH>
</TR>
<TR BGCOLOR="#000000">
	<TH>Ballot Name</TH>
	<TH>Vote Type</TH>
	<th>Ladder</th>
	<TH>Questions</TH>
	<TH>Vote</TH>
	<TH>Results</TH>
	<% If bSysAdmin Then %>
	<TH>Deactivate</TH>
	<% End If %>
</TR>
<%
'strSQL = "SELECT ballotid, BName, type, description, qCount, laddername FROM tbl_ballot inner join tbl_ladders on tbl_ballot.ladderid=tbl_ladders.ladderid WHERE IsActive=1 ORDER BY bName"
strSQL = "EXECUTE GetBallots 1"
oRS.Open strSQL, oConn
If Not(oRs.EOF AND oRS.BOF) Then
	do while not (ors.EOF)
		If bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		Response.Write "<TR BGCOLOR=" & bgc & ">"
		Response.Write "<TD>" & oRS("bName") & "</TD>"
		Response.Write "<TD ALIGN=CENTER>"
		If oRS.Fields("type").Value = 0 Then
			Response.Write "Global"
		Else
			Response.Write "Founder's only"
		End If
		Response.Write "</TD>"
		response.Write "<td align=center>" & Replace(ors.Fields("laddername").Value, " - ", " -<br />") & "</td>"
		Response.Write "<TD ALIGN=RIGHT>" & ors.Fields("qCount").Value & "</TD>"
		Response.Write "<TD align=center><A HREF=""ballot.asp?ballotid=" & oRS("BallotID") & """>vote</A></TD>"
		Response.Write "<TD align=center><A HREF=""results.asp?ballotid=" & oRS("BallotID") & """>results</A></TD>"
		If bSysAdmin Then
			Response.Write "<TD align=center><a href=""activate.asp?ballotid=" & oRS("BallotID") & """>deactivate</A></TD>"
		End If
		Response.Write "</TR>"
		Response.Write "<TR BGCOLOR=#000000><TD COLSPAN=" & intCols & "><FONT COLOR=""#888888"">&nbsp;&nbsp;&nbsp;&nbsp;" & oRs.Fields("Description").Value & "</FONT></TD></TR>"
		ors.movenext
	loop
Else
	%>
	<TR BGCOLOR="#000000"><TD COLSPAN=<%=intCols%>><I>No active ballots at this time.</i></TD></TR>
	<%
End If 
%>
</TABLE>
</TD></TR>
</TABLE>
<%
oRS.NextRecordSet
Call ContentEnd()

If bSysAdmin then
	Call ContentStart("Inactive Ballots")
	%>
	<TABLE BORDER=0 CElLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" width="97%">
	<TR><TD>
	<TABLE BORDER=0 CElLSPACING=1 CELLPADDING=2>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=7>Choose a ballot</TH>
	</TR>
	<TR BGCOLOR="#000000">
		<TH>Ballot Name</TH>
		<TH>Vote Type</TH>
		<TH>Ladder</TH>
		<TH>Questions</TH>
		<TH>Vote</TH>
		<TH>Results</TH>
		<TH>Activate</TH>
	</TR>
	<%
'	strSQL = "SELECT ballotid, BName, type, description, qCount, laddername FROM tbl_ballot "
'	strSQL = strSQL & " inner join tbl_ladders on tbl_ballot.ladderid=tbl_ladders.ladderid "
'	strSQL = strSQL & " WHERE IsActive=0 ORDER BY bName "
	strSQL = "EXECUTE GetBallots 0"
	oRS.Open strSQL, oConn
	If Not(oRs.EOF AND oRS.BOF) Then
		do while not (ors.EOF)
			If bgc = bgcone then
				bgc = bgctwo
			else
				bgc = bgcone
			end if
			Response.Write "<TR BGCOLOR=" & bgc & ">"
			Response.Write "<TD>" & oRS("bName") & "</TD>"
			Response.Write "<TD ALIGN=CENTER>"
			If oRS.Fields("type").Value = 0 Then
				Response.Write "Global"
			Else
				Response.Write "Founder's only"
			End If
			Response.Write "</TD>"
			If oRs.Fields("LadderType").Value = "T" Then 
				response.Write "<td align=center>" & ors.Fields("laddername").Value & " Ladder</td>"		
			Else
				response.Write "<td align=center>" & ors.Fields("laddername").Value & " League</td>"		
			End If
			Response.Write "<TD ALIGN=RIGHT>" & ors.Fields("qCount").Value & "</TD>"
			Response.Write "<TD align=center><A HREF=""ballot.asp?ballotid=" & oRS("BallotID") & """>vote</A></TD>"
			Response.Write "<TD align=center><A HREF=""results.asp?ballotid=" & oRS("BallotID") & """>results</A></TD>"
			Response.Write "<TD align=center><a href=""activate.asp?ballotid=" & oRS("BallotID") & """>activate</A></TD>"
			Response.Write "</TR>"
			Response.Write "<TR BGCOLOR=#000000><TD COLSPAN=7><FONT COLOR=""#888888"">&nbsp;&nbsp;&nbsp;&nbsp;" & oRs.Fields("Description").Value & "</FONT></TD></TR>"
			ors.movenext
		loop
	Else
		%>
		<TR BGCOLOR="#000000"><TD COLSPAN=<%=intCols%>><I>No inactive ballots at this time.</i></TD></TR>
		<%
	End If
	%>
	</TABLE>
	</TD></TR>
	</TABLE>
	<%
	oRS.NextRecordSet
	Call ContentEnd()
End If

%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>