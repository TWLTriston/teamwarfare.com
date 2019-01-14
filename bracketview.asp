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

'' Define the maximums for the display. More for speed than computational limits. 
'' Technically the computer will crunch until it get's the bracket, but even IE has a 
'' render slow down with over 256 bracket slots
CONST MAX_DIVISIONS = 16
CONST MAX_TEAMS_PER_DIVISION = 256
CONST MAX_TOTAL_TEAMS = 1024

'' Const because 3 just looks wierd in the algorythms. 
'' Changings this totally messes up the display. So dont touch it :P
CONST ROWS_BETWEEN_TEAMS = 3

'' Variables, bleh figure it out
Dim intDivisions, intTeamsPerDivision
Dim blnChangedTeams, blnCHangedDivisions, blnChangedAll
Dim iPowerOf2, iNextPowerOf2, iRound, iTeamPosition
Dim intRows, intRounds
Dim blnChangedDivisionMod2

'' Allow for dynamic changes to the display
'' Get querystring values for # of teams and # of divisions
'' Then check to make sure we arent calculating too many numbers
intTeamsPerDivision  = Request.QueryString("TeamsPerDivision")
If Len(intTeamsPerDivision) = 0 Or Not(IsNumeric(intTeamsPerDivision)) Then
	intTeamsPerDivision = MAX_TEAMS_PER_DIVISION
	blnChangedTeams = True
Else
	intTeamsPerDivision = cInt(intTeamsPerDivision)
End If

If intTeamsPerDivision > MAX_TEAMS_PER_DIVISION Then
	intTeamsPerDivision = MAX_TEAMS_PER_DIVISION
	blnChangedTeams = True
End If

intDivisions = Request.QueryString("Divisions")
If Len(intDivisions) = 0 Or Not(IsNumeric(intDivisions)) Then
	intDivisions = MAX_DIVISIONS
	blnChangedDivisions = True
Else
	intDivisions = cInt(intDivisions)
End if

If intDivisions > MAX_DIVISIONS Then
	intDivisions = MAX_DIVISIONS
	blnChangedDivisions = 1
End If

If intDivisions MOD 2 <> 0 Then
	intDivisions = 2
	blnChangedDivisionMod2 = 1
End If

If intDivisions * intTeamsPerDivision > MAX_TOTAL_TEAMS Then
	intTeamsPerDivision = MAX_TOTAL_TEAMS 
	intDivisions = 2
	blnChangedAll = True
End If
%>
<html>
<head>
	<title>TWL Brackets</title>
	<style>
	<!--
	body, th, td { BACKGROUND-COLOR: #000000; COLOR: #FFFFFF; FONT-FAMILY: Verdana; FONT-SIZE: 10px;}
	.cssTeamSlot { BACKGROUND-COLOR: #005555; BORDER: 1px SOLID #FFD142;}
	.cssFillerRight { FONT-SIZE: 4px; BORDER-RIGHT: 1px SOLID #FFD142;}
	.cssFillerLeft { FONT-SIZE: 4px; BORDER-LEFT: 1px SOLID #FFD142;}
	.cssEmpty { FONT-SIZE: 4px; }
	.cssChanged { COLOR: #FF0000; }
	A, A:link, A:hover, A:active, A:visited { TEXT-DECORATION: Underline; COLOR: #FFFF00; }
	//-->
	</style>
</head>
<body>
<%
'' Figure out how many rows and rounds we are expecting to display
intRows = ((intDivisions / 2) * intTeamsPerDivision) + ROWS_BETWEEN_TEAMS * ((intDivisions / 2) * intTeamsPerDivision - 1)
intRounds = Log(intTeamsPerDivision) / Log(2) + Log(intDivisions / 2) / Log(2) + 1

'' Tell the user what they are seeing
Response.Write "Showing " & intTeamsPerDivision
If intTeamsPerDivision = 1 Then
	Response.Write " team "
Else
	Response.Write " teams "
End If
Response.Write " across " & intDivisions 

If intDivisions = 1 Then
	Response.Write " division."
Else
	Response.Write " divisions."
End If

Response.Write " ( " & intTeamsPerDivision * intDivisions

If intTeamsPerDivision * intDivisions = 1 Then
	Response.Write " team."
Else
	Response.Write " total teams."
End If

Response.Write " )<br>"

'' If we changed the values of the querystring, tell the user so we dont 
'' confuse them, and have them telling us our stuff doesnt work
If blnChangedTeams Then
	Response.Write "<span class=""cssChanged""># of teams changed due to no 'TeamsPerDivision' on the querystring, or a value of greater than " & MAX_TEAMS_PER_DIVISION & ".</span><br>"
End If
If blnChangedDivisions Then
	Response.Write "<span class=""cssChanged""># of divisions changed due to no 'Divisions' on the querystring, or a value of greater than " & MAX_DIVISIONS & ".</span><br>"
End If
If blnChangedDivisionMod2 Then
	Response.Write "<span class=""cssChanged""># of divisions changed due to divisions not being a multiple of 2 (shame shame).</span><br>"
End If
If blnChangedAll Then
	Response.Write "<span class=""cssChanged""># of divisions and # of teams changed due to too many total teams, greater than " & MAX_TOTAL_TEAMS & ". Preventing dangerous loop size.</span><br>"
End If

'' Give a link that will format the querysting
Response.Write "<a href=""bracketview.asp?TeamsPerDivision=8&Divisions=1"" alt=""Properly formatted querysting."">Click here for a properly formatted querystring to allow for changes to the display.</a><br>"	
Response.Write "<br>"'' Ahh, finally the good stuff
''' This is where all the power is (and unfortunately all the processor time
''' This thing causes major processor spikes to do this looping math
''' Remember, you are dealing with a matrix of rounds * teams * brackets (3 + (teams * brackets))
''' In short, a 4 team 1 division tournament consists of an 18X3 matrix (it's big)

'' Start the table
Response.Write vbCrLf & vbCrLf 
Response.Write "<table cellspacing=""0"">" & vbCrLf 
Response.Write "<tr>" & vbCrLf 

Dim iCurrentRound
'' Give the table a header.. in this case, rounds (may change the placement of the round header later
'' but for now it works well
For iRound = 1 to (intRounds * 2 - 1)
	iCurrentRound = IntRounds - Abs(intRounds - iRound)
	Response.Write vbTab & "<th width=""60"">Round " & iCurrentRound & "</th>" & vbCrLf 
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
		If (iTeamPosition MOD (iPowerOf2)) = 0 AND (iTeamPosition / iPowerOf2) MOD 2 = 1 Then
			Response.Write vbTab & "<td width=""60"" class=""cssTeamSlot"">&nbsp;</td>" & vbCrLf 
		Else
			'' If this isnt a seeded table cell, see if the cell isnear the next cell that will be seeded, 
			'' and give it a border so the lines can be tracked
			If (iTeamPosition / iNextPowerOf2 ) MOD 2 = 1 And iCurrentRound <> intRounds Then
				If iRound < intRounds Then
					Response.Write vbTab & "<td width=""60"" class=""cssFillerRight"">&nbsp;</td>" & vbCrLf 
				Else
					Response.Write vbTab & "<td width=""60"" class=""cssFillerLeft"">&nbsp;</td>" & vbCrLf 
				End If
			Else
				'' Other wise, this is an empty cell, just put the equivilance to nothing in here.
				Response.Write vbTab & "<td width=""60"" class=""cssEmpty"">&nbsp;</td>"& vbCrLf 
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
	