<%
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' BracketView.asp
' Myles Angell 
' triston@teamwarfare.com
' 5/1/2002
' Version: 1.0
'
' Purpose: Produce a single 
'	elimination simple bracket display 
'
' Future Enhancements:
'	Have the brackets come in at each other, 
'	instead of one long list. One potential problem
'	Exists with the final match up.. how to display it :)
'	Most tournament brackets use the "slots" to display
'	who is playing who, not who is at that slot (as we do)
'	Other than that, it's just some more math, 
'	and doing some absolute value (you know, neg numbers) checking
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CONST ROWS_BETWEEN_TEAMS = 3
Dim iCurrentRound, iTeamPosition
Dim oConn, oRS, strSQL, intTournamentID
Dim strTournamentName, intTeamsPerDivision, intDivisions
Dim intRows, intRounds, iRound, iPowerOf2, iNextPowerOf2
Dim intRoundForLabel, iDivision, blnOnProduction 

blnOnProduction = True
intTournamentID = 16
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open Application("ConnectStr")

Set oRS = Server.CreateObject("ADODB.RecordSet")

strSQL = "SELECT TournamentName, TeamsPerDiv, Divisions FROM tbl_tournaments WHERE TournamentID = " & intTournamentID
oRS.Open strSQL, oConn, 1, 1
If Not(oRS.EOF AND oRS.BOF) Then
	strTournamentName = oRS.Fields("TournamentName").Value
	intTeamsPerDivision = oRS.Fields("TeamsPerDiv").Value
	intDivisions = oRS.Fields("Divisions").Value
End If
oRS.NextRecordset
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
	<title>TWL Bracket</title>
	<style>
	<!--
	.t { BACKGROUND-COLOR: <%=Application("bgcone")%>; BORDER: 1px SOLID #444444; PADDING: 2px; FONT-SIZE: 9px;}
	.d { PADDING: 2px; FONT-SIZE: 9px; FONT-WEIGHT: 600; }
	.r { FONT-SIZE: 4px; BORDER-RIGHT: 1px SOLID #444444; }
	.l { FONT-SIZE: 4px; BORDER-LEFT: 1px SOLID #444444; }
	.e { FONT-SIZE: 4px; }
	.c { COLOR: #FF0000; }
	A, A:hover, A:link, A:active, A:visited { COLOR: #FFD142; FONT-FAMILY: Verdana; FONT-SIZE: 9px; }
	//-->
	</style>
	<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
</head>
<body bgcolor="#000000" text="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<%
Dim arrRounds(8, 256)
Dim arrTeamNames(8, 256)
Dim arrTeamTags(8, 256)
Dim i, j, iSeed
For i = 0 to 8
	For j = 0 to 256
		arrRounds(i,j) = i & ", " & j
		arrTeamNames(i,j) = Null
		arrTeamTags(i,j) = Null
	Next
Next
intRoundForLabel = Log(intTeamsPerDivision) / Log(2) + 1
'' Figure out how many rows and rounds we are expecting to display
intRows = ((intDivisions / 2) * intTeamsPerDivision) + ROWS_BETWEEN_TEAMS * ((intDivisions / 2) * intTeamsPerDivision - 1)
intRounds = Log(intTeamsPerDivision) / Log(2) + Log(intDivisions / 2) / Log(2) + 1

If blnOnProduction Then
	For iRound = 1 to intRounds
		strSQL = "EXECUTE GetTournamentArray " & intTournamentID & ", " & iRound
		'Response.Write strSQL
		oRS.Open strSQl, oConn
		If oRS.State = 1 Then
			If Not(oRS.EOF AND oRS.BOF) Then
				Do While Not(oRS.EOF)
					arrTeamNames(iRound, oRS.Fields("ArrayNumber").Value) = oRS.Fields("TeamName").Value
					arrTeamTags(iRound, oRS.Fields("ArrayNumber").Value) = oRS.Fields("TeamTag").Value
					oRS.MoveNext
				Loop
			End If
			oRS.NextRecordset
		End If
	Next
End If
'' Ahh, finally the good stuff
''' This is where all the power is (and unfortunately all the processor time
''' This thing causes major processor spikes to do this looping math
''' Remember, you are dealing with a matrix of rounds * teams * brackets (3 + (teams * brackets))
''' In short, a 4 team 1 division tournament consists of an 18X3 matrix (it's big)

'' Start the table
Response.Write "<table cellspacing=""0"" align=""center"" cellpadding=""0"" border=""0"" width=""100%"">" & vbCrLf 
Response.Write "<tr>" & vbCrLf
Response.Write vbTab & "<th nowrap=""nowrap"" colspan=""" & (intRounds * 2 - 1) & """>" & strTournamentName & "</th>" & vbCrLf
Response.Write "</tr>" & vbCrLf

Response.Write "<tr>" & vbCrLf 

'' Give the table a header.. in this case, rounds (may change the placement of the round header later
'' but for now it works well
For iRound = 1 to (intRounds * 2 - 1)
	iCurrentRound = IntRounds - Abs(intRounds - iRound)
	Response.Write vbTab & "<th nowrap=""nowrap"">Round " & iCurrentRound & "</th>" & vbCrLf 
Next

Response.Write "</tr>" & vbCrLf 
'' Start going row by row 
For iTeamPosition = 2 to intRows + 1
	Response.Write "<tr>" & vbCrLf
	'' Then go column by column
	For iRound = 1 to (intRounds * 2 - 1)
		iCurrentRound = IntRounds - Abs(intRounds - iRound)
		' This system uses power's of 2 to figure out how to display a bracket
		' There may be another faster way, but this is the pattern i discovered in my testing
		'' do tell if there is another way
		iPowerOf2 = 2 ^ iCurrentRound
		iNextPowerOf2 = 2 ^ (iCurrentRound + 1)
		'' First check to see if this is a "seeded" table cell. If so, give it some color, 
		'' at a later date we can reference an array in this slot to plop names into these colored boxes
		If intRoundForLabel - 1 = (iCurrentRound) AND ((iTeamPosition) MOD (iNextPowerOf2)) = 0 AND ((iTeamPosition) / iNextPowerOf2) MOD 2 = 1 Then
			If iRound <= intRounds Then
				iDivision = ((iTeamPosition / iNextPowerOf2) + 1) / 2
				Response.Write vbTab & "<td align=""left"" nowrap=""nowrap"" class=""d"">Division " & iDivision & "</td>" & vbCrLf 
			Else
				iDivision = (((iTeamPosition / iNextPowerOf2) + 1) / 2) + 2 ^ (iRound - intRounds) - intDivisions / 2
				Response.Write vbTab & "<td align=""right"" nowrap=""nowrap"" class=""d"">Division " & iDivision & "</td>" & vbCrLf 
			End If
		ElseIf (iTeamPosition MOD (iPowerOf2)) = 0 AND (iTeamPosition / iPowerOf2) MOD 2 = 1 Then
			If iRound <= intRounds Then
				iSeed = ((iTeamPosition / iPowerOf2) + 1) / 2
			Else
				iSeed = (((iTeamPosition / iPowerOf2) + 1) / 2) + 2 ^ (iRound - intRounds)
			End If
			If IsNull(arrTeamNames(iCurrentRound, iSeed)) Then
				If iCurrentRound = 1 Then
					Response.Write vbTab & "<td nowrap=""nowrap"" class=""t""><b>Open</b></td>" & vbCrLf 
				Else
					Response.Write vbTab & "<td nowrap=""nowrap"" class=""t"">TBD</td>" & vbCrLf 
				End If
			Else
				If Len(arrTeamTags(iCurrentRound, iSeed)) > 0 Then
					Response.Write vbTab & "<td nowrap=""nowrap"" class=""t""><a href=""http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode(arrTeamNames(iCurrentRound, iSeed)) & """>" & Server.HTMLEncode(arrTeamTags(iCurrentRound, iSeed)) & "</a></td>" & vbCrLf 
				Else
					Response.Write vbTab & "<td nowrap=""nowrap"" class=""t""><a href=""http://www.teamwarfare.com/viewteam.asp?team=" & Server.URLEncode(arrTeamNames(iCurrentRound, iSeed)) & """>" & Server.HTMLEncode(Left(arrTeamNames(iCurrentRound, iSeed), 10)) & "</a></td>" & vbCrLf 
				End If
			End If
		Else
			'' If this isnt a seeded table cell, see if the cell isnear the next cell that will be seeded, 
			'' and give it a border so the lines can be tracked
			If (iTeamPosition / iNextPowerOf2 ) MOD 2 = 1 And iCurrentRound <> intRounds Then
				If iRound < intRounds Then
					Response.Write vbTab & "<td nowrap=""nowrap"" class=""r"">&nbsp;</td>" & vbCrLf 
				Else
					Response.Write vbTab & "<td nowrap=""nowrap"" class=""l"">&nbsp;</td>" & vbCrLf 
				End If
			Else
				'' Other wise, this is an empty cell, just put the equivilance to nothing in here.
				Response.Write vbTab & "<td nowrap=""nowrap"" class=""e"">&nbsp;</td>"& vbCrLf 
			End If
		End If
	'' Go for a round trip
	Next
	Response.Write "</tr>"& vbCrLf 
'' Next row...
Next

'' All done, wasn't that fun, I thought so !
Response.Write "</table>" & vbCrLf 
%>
</body>
</html>
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>