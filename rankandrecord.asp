<%
team = replace(Request.QueryString ("team"), "'", "''")
ladder = replace(Request.QueryString ("ladder"), "'", "''")
rank = 0
if len(trim(team)) <> 0 and len(trim(ladder)) <> 0 then
	set oconn = server.CreateObject ("ADODB.Connection")
	set ors = server.CreateObject ("ADODB.Recordset")
	oconn.Open Application("ConnectStr")
	strsql = "select rank, wins, losses, forfeits from lnk_t_l lnk, tbl_teams t, tbl_ladders l "
	strsql = strsql & " where l.ladderid = lnk.ladderid "
	strsql = strsql & " AND t.teamid = lnk.teamid "
	strsql = strsql & " AND t.teamname = '" & team & "'"	
	strsql = strsql & " AND l.laddername = '" & ladder & "'"
	ors.Open strsql, oconn
	if not(ors.EOF and ors.BOF) then
		rank = ors.Fields ("rank").Value
		wins = ors.Fields ("wins").Value
		losses = ors.Fields ("losses").Value
		forfeits = ors.Fields ("forfeits").Value
	end if
	ors.Close
	set ors = nothing
	oconn.Close
	set oconn = nothing
end if
Response.Write "Rank-" & Rank & vbCrLf
Response.Write "Wins-" & wins & vbCrLf
Response.Write "Losses-" & losses & vbCrLf
Response.Write "Forfeits-" & forfeits & vbCrLf

%>