<% response.buffer = true %>
<% 
dim bigolearray(21)
set oconn = server.CreateObject ("ADODB.Connection")
oconn.Open Application("ConnectSTR")
set ors = server.CreateObject ("ADODB.Recordset")
strSQL = "select top 5 * from vHistory where matchLadderid = 5 and WinnerRank > 0 AND WinnerRank <= 5 AND MatchForfeit <> 1 order by matchdate desc"
ors.Open strsql, oconn
if not(ors.EOF and ors.BOF) then
'	Response.Write "strWinnerName, strLoserName, intWinnerRank, intLoserRank, intWinnerID (1)"
'	Response.write "strMap1WinnerScore, strMap1LoserScore, strMap1Name, "
'	Response.write "strMap2WinnerScore, strMap2LoserScore, strMap2Name, "
'	Response.write "strMap3WinnerScore, strMap3LoserScore, strMap3Name, "
'	Response.write "dtimMatchDate, strLadderURL, intLadderID " & VBCRLF
	do while not(ors.EOF)
		if ors.Fields("WinnerDefending").Value then
			bigolearray(0) = ors.Fields("WinnerName").Value
			bigolearray(1) = ors.Fields("LoserName").Value
		ELSE
			bigolearray(1) = ors.Fields("WinnerName").Value
			bigolearray(0) = ors.Fields("LoserName").Value
		END IF
		bigolearray(2) = ors.Fields("WinnerRank").Value
		bigolearray(3) = ors.Fields("LoserRank").Value
		
		if ors.Fields("WinnerDefending").Value then
			bigolearray(4) = "1"
		else
			bigolearray(4) = "2"
		end if				
		bigolearray(5) = ors.Fields("MatchMap1DefenderScore").Value
		bigolearray(6) = ors.Fields("MatchMap1AttackerScore").Value
		bigolearray(7) = ors.Fields("MatchMap1").Value

		bigolearray(8) = ors.Fields("MatchMap2DefenderScore").Value
		bigolearray(9) = ors.Fields("MatchMap2AttackerScore").Value
		bigolearray(10) = ors.Fields("MatchMap2").Value

		bigolearray(11) = ors.Fields("MatchMap3DefenderScore").Value
		bigolearray(12) = ors.Fields("MatchMap3AttackerScore").Value
		bigolearray(13) = ors.Fields("MatchMap3").Value

		bigolearray(14) = abs(datediff("s", ors.Fields("MatchDate").Value, "1/1/1970"))
		bigolearray(15) = "http://www.teamwarfare.com/viewladder.asp?ladder=Tribes+2+CTF"
		bigolearray(16) = "5"

		for i = 0 to 15
			Response.Write bigolearray(i) & VBTAB
		next
		Response.Write bigolearray(16) & VBCRLF	
		ors.MoveNext
	loop
else
	Response.Write "Problem accessing data, please try again."
end if
ors.Close
oconn.Close
set ors = nothing
set oconn = nothing
%>
<% response.end %>