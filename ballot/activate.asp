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
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<%
If Not(bSysAdmin) Then 
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If
Dim intBallotID
intBallotID = Request.QueryString ("bid")
If Len(intBallotID) = 0 Then
	intBallotID = Request.QueryString ("ballotid")
End If
strSQL = "UPDATE tbl_ballot set isactive=ABS (isActive - 1) WHERE BallotID = " & intBallotID
oconn.Execute strsql

oConn.Close
Set oConn = Nothing
Set oRS = Nothing

Response.Clear 
Response.Redirect "/ballot"

%>
