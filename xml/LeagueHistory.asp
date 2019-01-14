<% Option Explicit %>
<!-- #include virtual="/include/xml.asp" -->
<%
Server.ScriptTimeout = 45

Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Dim strLeagueName
strLeagueName = Request.QueryString("League")

Dim intLeagueID
strSQL = "SELECT LeagueID FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
Else
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=7"
End If
oRs.NextRecordset

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<%
Response.Write "<leaguehistory>" & vbCrLf
Response.Write vbTab & "<leagueinformation>" & vbCrLf
Response.Write vbTab & vbTab & "<leaguename>" & XMLEncode(Server.HTMLEncode(strLeagueName)) & "</leaguename>" & vbCrLf
Response.Write vbTab & vbTab & "<leaguelink>" & vbCrLf
Response.Write vbTab & vbTab & vbTab & "<httplink>" & Server.HTMLEncode("http://www.teamwarfare.com/xml/viewleague.asp?league=" & XMLEncode(Server.URLEncode(strLeagueName))) & "</httplink>" & vbCrLf
Response.Write vbTab & vbTab & vbTab & "<xmllink>" & Server.HTMLEncode("http://www.teamwarfare.com/xml/viewleague.asp?league=" & XMLEncode(Server.URLEncode(strLeagueName))) & "</xmllink>" & vbCrLf
Response.Write vbTab & vbTab & "</leaguelink>" & vbCrLf
Response.Write vbTab & vbTab & "<servertime>" & Now() & "</servertime>" & vbCrLf
Response.Write vbTab & "</leagueinformation>" & vbCrLf

Dim strWeek
strWeek = Request.Querystring("weekago")
if len(strWeek) = 0 Then
	strWeek = 0
End If
strSQL = "EXECUTE LeagueGetHistory @LeagueID = '" & intLeagueID & "', @XFactor='"  & strWeek  & "'"
'Response.Write strSQL
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	Do While Not(oRs.EOF) 
		Response.Write vbTab & "<match matchid=""" & oRs.Fields("LeagueMatchID").Value & """ conference=""" & XMLEncode(oRs.Fields("ConferenceName").Value)  & """ division=""" & XMLEncode(oRs.Fields("DivisionName").Value) & """>" & vbCrLf
		Response.Write vbTab & vbTab & "<team position=""home"">" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<name>" & XMLEncode(Server.HTMLEncode(oRs.Fields("HomeTeamName").Value)) & "</name>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<tag>" & XMLEncode(Server.HTMLEncode(oRs.Fields("HomeTeamTag").Value)) & "</tag>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<teamlink>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>" & Server.HTMLEncode("http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLEncode(oRs.Fields("HomeTeamName").Value))) & "</httplink>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>" & Server.HTMLEncode("http://www.teamwarfare.com/xmp/viewteam_v2.asp?team=" & XMLEncode(Server.URLEncode(oRs.Fields("HomeTeamName").Value))) & "</xmllink>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "</teamlink>" & vbCrLf
		Response.Write vbTab & vbTab & "</team>" & vbCrLf
		Response.Write vbTab & vbTab & "<team position=""visitor"">" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<name>" & XMLEncode(Server.HTMLEncode(oRs.Fields("VisitorTeamName").Value)) & "</name>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<tag>" & XMLEncode(Server.HTMLEncode(oRs.Fields("VisitorTeamTag").Value)) & "</tag>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "<teamlink>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & vbTab & "<httplink>" & Server.HTMLEncode("http://www.teamwarfare.com/viewteam.asp?team=" & XMLEncode(Server.URLEncode(oRs.Fields("VisitorTeamName").Value))) & "</httplink>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & vbTab & "<xmllink>" & Server.HTMLEncode("http://www.teamwarfare.com/xmp/viewteam_v2.asp?team=" & XMLEncode(Server.URLEncode(oRs.Fields("VisitorTeamName").Value))) & "</xmllink>" & vbCrLf
		Response.Write vbTab & vbTab & vbTab & "</teamlink>" & vbCrLf
		Response.Write vbTab & vbTab & "</team>" & vbCrLf
		Response.Write vbTab & vbTab & "<matchdate>" & FormatDateTime(oRs.Fields("MatchDate").Value, 2) & "</matchdate>" & vbCrLf

		Dim i
		For i = 1 to 5
			If Len(oRs.Fields("Map" & i).Value) > 0 Then
				Response.Write vbTab & vbTab & "<map order=""" & i & """"
				Response.Write " name=""" & XMLEncode(oRs.Fields("Map" & i).Value) & """"
				Response.Write " homescore=""" & oRs.Fields("Map" & i & "VisitorScore").Value & """"
				Response.Write " visitorscore=""" & oRs.Fields("Map" & i & "HomeScore").Value & """ />"
			End If
		Next
		Response.Write vbTab & "</match>" & vbCrLf

		oRs.MoveNext
	Loop
End If
oRs.NextRecordSet

Response.Write "</leaguehistory>" & vbCrLf

oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>