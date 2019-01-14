<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Activate Ballot"

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

if not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Activate a Ballot")

strSQL = "SELECT ballotid, BName FROM tbl_ballot WHERE Avail=1 ORDER BY bName"
oRS.Open strSQL, oConn
If Not(oRs.EOF AND oRS.BOF) Then
	%>
	<TABLE BORDER=0 CElLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
	<TR><TD>
	<TABLE BORDER=0 CElLSPACING=1 CELLPADDING=2>
	<TR BGCOLOR="#000000">
		<TH>Choose a ballot</TH>
	</TR>
	<%
	do while not (ors.EOF)
		If bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		Response.write "<TR BGCOLOR=" & bgc & "><TD><A HREF=""activate.asp?bid=" & oRS("BallotID") & """>" & oRS("bName") & "</A></TD></TR>"
		ors.movenext
	loop
	%>
	</TABLE>
	</TD></TR>
	</TABLE>
	<%
End If
ors.Close 
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>