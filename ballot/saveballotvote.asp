<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Save Ballot"

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

Dim q, playerid, i

%>
<!-- #include virtual="/include/i_funclib.asp" -->
<%
If Not(Session("LoggedIn")) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=2"
End If
PlayerID = Session("PlayerID")

q=Request("qcount")
for i= 1 to q
	if len(request("q_" & i & "_vote")) > 0 then
		strsql = "select VoterID from tbl_votes where QID='" & request("QID_" & i) & "' AND VoterID ='" & playerid & "'"
		ors.Open strSQL, oConn
		If ors.EOF and Ors.BOF then
			strsql="insert into tbl_votes(QID, Choice, VoterID) values (" & request("QID_" & i) & "," & request("q_" & i & "_vote") & "," & playerid & ")"
			oconn.Execute strsql
		End If
		ors.NextRecordset 
	end if
next
oConn.Close 
Set oConn = Nothing
Set oRS = Nothing

Response.Clear 
Response.Redirect "results.asp?ballotid=" & Request("BallotID")
%>
