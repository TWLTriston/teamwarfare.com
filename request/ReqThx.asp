<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("player"), """", "&quot;") 

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLadderAdmin, bLoggedIn
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
bLoggedIn = Session("LoggedIn")

Dim strPlayerName, intPlayerID
strPlayerName = Request("Player")

Dim strTeamName
Dim TMLinkID, DivID, TournamentName
Dim bBarDone
Dim bCanAdmin
Dim strLadderName, intRank, intLosses, intPlayerLadderID
Dim intForfeits, intWins, strStatus, strEnemyName, strResult
Dim linkID, map, opponent, mDate, statusVerbage
bBarDone = False
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Player Name Change Request") 
	%>
		<CENTER>
		Your new user name request has been submitted.<br />
		<br />
		If your name change request is approved, your name will simply change the next time you log in.
		<br />
		<br />
		If your name change request is denied, you will be notified by email.
		</center>
	<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>