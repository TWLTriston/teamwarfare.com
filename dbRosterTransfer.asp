<%
 	response.buffer=true
%>
<!-- #include file="include/i_funclib.asp" -->
<% 	
	Set oConn = Server.CreateObject("ADODB.Connection")
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	Set oRS2 = Server.CreateObject("ADODB.RecordSet")
	
	oConn.ConnectionString = Application("ConnectStr")
	oConn.Open 	
 	
If Not(IsSysAdmin()) Then
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

If request.form("saveType")="LeagueToLeague" then
	strSQL = "EXECUTE RosterTransferLeagueToLeague @FromLeague = '" & Request.Form("selFromLeagueID") & "', @ToLeague = '" & Request.Form("selToLeagueID") & "'"
	oConn.Execute(strSQL)
	oConn.Close
	Set oConn = nothing
	Response.Clear
	Response.Redirect "RosterTransfer.asp?s=1"
End If

If request.form("saveType")="LeagueToTournament" then
	strSQL = "EXECUTE RosterTransferLeagueToTournament @FromLeagueID = '" & Request.Form("selFromLeagueID") & "', @ToTournamentID = '" & Request.Form("selToTournamentID") & "'"
	oConn.Execute(strSQL)
	oConn.Close
	Set oConn = nothing
	Response.Clear
	Response.Redirect "RosterTransfer.asp?s=1"
End If

%>
 		
 		