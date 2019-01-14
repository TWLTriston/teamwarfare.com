<% Option Explicit %>
<!-- #include virtual="/include/xml.asp" -->
<%
Server.ScriptTimeout = 45

Response.ContentType = "text/xml"
Response.Write "<?xml version=""1.0"" encoding=""ISO-8859-1""?>" & vbCrLf

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Dim strLeagueName, intLeagueID
strLeagueName = Request.QueryString("League")
If Len(Trim(strLeagueName)) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If

strSQL = "SELECT LeagueID, LeagueName FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "' AND LeagueActive = 1"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
	strLeagueName = oRs.Fields("LeagueName").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
oRs.NextRecordSet

Dim dtmDate, intTimeZoneDifference, strDate, strTime
intTimeZoneDifference = 0

Dim strDateMask, bln24HourTime, blnVerticalBars, strColumnColor1, strColumnColor2
strDateMask = "MM-DD-YYYY"
bln24HourTime = False

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<%
Response.Write "<leaguematches>" & vbCrLf
Response.Write vbTab & "<leagueinformation>" & vbCrLf
Response.Write vbTab & vbTab & "<leaguename>" & XMLEncode(Server.HTMLEncode(strLeagueName)) & "</leaguename>" & vbCrLf
Response.Write vbTab & vbTab & "<leaguelink>" & vbCrLf
Response.Write vbTab & vbTab & vbTab & "<httplink>" & Server.HTMLEncode("http://www.teamwarfare.com/xml/viewleague.asp?league=" & XMLEncode(Server.URLEncode(strLeagueName))) & "</httplink>" & vbCrLf
Response.Write vbTab & vbTab & vbTab & "<xmllink>" & Server.HTMLEncode("http://www.teamwarfare.com/xml/viewleague.asp?league=" & XMLEncode(Server.URLEncode(strLeagueName))) & "</xmllink>" & vbCrLf
Response.Write vbTab & vbTab & "</leaguelink>" & vbCrLf
Response.Write vbTab & vbTab & "<servertime>" & Now() & "</servertime>" & vbCrLf
Response.Write vbTab & "</leagueinformation>" & vbCrLf


strSQL = "EXECUTE LeagueGetMatches @LeagueID = '" & intLeagueID & "'"
If Len(Request.QueryString("X")) > 0 AND IsNumeric(Request.QueryString("X")) Then
	strSQL = strSQL & ", @XFactor = '-" & Request.QueryString("X") & "'"
End If
'Response.Write strSQL
oRs.Open strSQL, oConn
If (oRs.State = 1) Then 
	If Not(oRs.EOF AND oRs.BOF) Then
		Do While Not oRs.EOF
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
					Response.Write " name=""" & XMLEncode(oRs.Fields("Map" & i).Value) & """ />"
				End If
			Next
			Response.Write vbTab & "</match>" & vbCrLf
			oRs.MoveNext
		Loop
	End if
End if
Response.Write "</leaguematches>" & vbCrLf

oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>