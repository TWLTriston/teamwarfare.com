 <link rel="stylesheet" href="http://www.wsbn.cc/tribes_dev/tribes.css" type="text/css">
<style>
p.small
{
    COLOR: #66ccff;
    FONT-FAMILY: Arial, Helvetica, sans-serif;
    FONT-SIZE: 11px;
    FONT-WEIGHT: bold;
    TEXT-DECORATION: none
}
</style>
<%
Function JavaEncode(TheString)
   if TheString <> "" then 
   	JavaEncode = Replace(TheString, "'", "\'")
   	JavaEncode = Replace(JavaEncode, """", "&quot;")
   else
 	JavaEncode=""
   end if
End Function

	bgcone = "#212152"
	bgctwo = "#314D6B"
LadderName = Request.QueryString("ladder")

perpage = Request.QueryString("perpage")

if perpage= "" then
	perpage = 25
end if
	set oConn = Server.CreateObject("ADODB.Connection")
	oConn.connectionstring="file name=c:\twl.udl"
	oConn.Open 
	set oRs = Server.CreateObject("ADODB.Recordset")
	set oRs2 = Server.CreateObject("ADODB.Recordset")
	set oRs1 = Server.CreateObject("ADODB.Recordset")
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#102C4A">
	<tr><td align=center>
       <%
	strSQL = "select LadderID from tbl_Ladders where LadderName='" & replace(LadderName, "'", "''") & "'"
	ors.Open strSQL, oConn
	if not (ors.EOF and ors.BOF) then
		ladderid = ors.Fields(0).Value 
	end if
	ors.Close 
	strSQL="SELECT top " & perpage & " tbl_Teams.TeamName, lnk_T_L.Rank, tbl_Ladders.LadderName, lnk_T_L.Status, lnk_T_L.TLLinkID, lnk_T_L.wins, lnk_T_L.losses, lnk_T_L.forfeits, tbl_teams.TeamTag "
	strSQL= strSQL & "FROM tbl_Ladders INNER JOIN (tbl_Teams INNER JOIN "
	strSQL= strSQL & "lnk_T_L ON tbl_Teams.TeamID = lnk_T_L.TeamID) ON tbl_Ladders.LadderID = lnk_T_L.LadderID "
	strSQL= strSQL & "WHERE tbl_Ladders.LadderName='" & replace(LadderName, "'", "''") & "' and lnk_T_L.isactive=1 ORDER BY lnk_T_L.Rank"
	oRs.Open strSQL, oConn

if not (ors.EOF and ors.BOF) then
	%>
	<table align=center border=0 cellspacing=1 width=97%>
	<tr><td width=40 align=center><p class=whiteLargeText><b>Rank</b></p></td>
	<td width=160 align=center><p class=whiteLargeText><b>Team</b></p></td>
	<td width=120 align=right><p class=whiteLargeText><b>Record (Llamas)</b></p></td>
	<td width=300 align=center><p class=whiteLargeText><b>Status</b></p></td></tr>
	<%
	bgc = bgctwo
	do while not ors.EOF
		%>
		<tr bgcolor=<%=bgc%> height=40 valign=center ><td width=40 align=center><font size=1><%= ors.Fields(1).Value%></b></font></td>
		<td width=160 align=left><p class="small">&nbsp;<a href=viewTeam.asp?team=<% = server.urlencode(ors.Fields(0).Value) %> onMouseOver="(window.status='View profile for <%=javaencode(ors.fields(0).value)%>'); return true" onMouseOut="(window.status='<%=javatitle%>'); return true"><% = server.htmlencode(ors.Fields(0).Value) & " " & server.htmlencode(ors.fields(8).value)%></a></p></td>
		<td width=80 align=right><font size=1><%=ors.Fields(5).Value%> / <%=ors.Fields(6).Value%> - (<%=ors.Fields(7).Value%>)</b></font></td>
		<%
		if (ors.Fields(3).Value <> "Available" and left(ors.Fields(3).Value,8) <> "Defeated" and left(ors.Fields(3).Value,6) <> "Immune" and ors.Fields(3).Value <> "Resting") then
			if ors.Fields(3).Value = "Defending" then
				strSQL = "select MatchAttackerID, MatchDate, MatchMap1ID, Matchmap2ID, MatchMap3ID from tbl_Matches where MatchDefenderID = " & ors.Fields(4).Value & " and MatchLadderID=" & ladderid
				set ors2 = Server.CreateObject("ADODB.Recordset")
				ors2.Open strSQL, oconn
				enemyID=-1
				if not (ors2.EOF and ors2.BOF) then
					enemyId = ors2.Fields(0).Value
					mDate = ors2.Fields(1).Value
					map1 = ors2.fields(2).value
					map2 = ors2.fields(3).value
					map3 = ors2.fields(4).value
				end if
				ors2.Close 
				strSQL = "SELECT  tbl_Teams.TeamName, tbl_Matches.MatchDate, tbl_Teams.TeamTag FROM tbl_Matches, tbl_Teams INNER JOIN "
				strSQL = strSQL & "lnk_T_L ON tbl_Teams.TeamID = lnk_T_L.TeamID "
				strSQL = strSQL & "WHERE (((lnk_T_L.TLLinkID)=" & enemyID & "));"
				ors2.Open strSQL, oconn
				if not (ors2.EOF and ors2.BOF) then 
					enemy = ors2.Fields(0).value
				end if
				ors2.Close 
			end if
			if ors.Fields(3).Value = "Attacking" then
				strSQL = "select MatchDefenderID, MatchDate, MatchMap1ID, MatchMap2ID, MatchMap3ID from tbl_Matches where MatchAttackerID = " & ors.Fields(4).Value & " and MatchLadderID=" & ladderid
				set ors2 = Server.CreateObject("ADODB.Recordset")
				ors2.Open strSQL, oconn
				enemyID=-1
				if not (ors2.EOF and ors2.BOF) then
					enemyId = ors2.Fields(0).Value
					mDate = ors2.Fields(1).Value
					map1 = ors2.fields(2).value
					map2 = ors2.fields(3).value
					map3 = ors2.fields(4).value
				end if
				ors2.Close 
				strSQL = "SELECT  tbl_Teams.TeamName, tbl_Matches.MatchDate, tbl_teams.teamtag FROM tbl_Matches, tbl_Teams INNER JOIN "
				strSQL = strSQL & "lnk_T_L ON tbl_Teams.TeamID = lnk_T_L.TeamID "
				strSQL = strSQL & "WHERE (((lnk_T_L.TLLinkID)=" & enemyID & "));"
				ors2.Open strSQL, oconn
				if not (ors2.EOF and ors2.BOF) then 
					enemy = ors2.Fields(0).value
				end if
				ors2.Close 
			end if
			if mDate <> "TBD" then
				newMDate = right(mDate, len(mDate)-instr(1, mDate, ","))
				newMDate = left(newmdate, len(newmdate)-4)
				newmdate=formatdatetime(newmdate, 2)
				mm=month(newmdate)
				dd=day(newmdate)
				pDate=mm & "/" & dd
			else
				pdate="TBD"
			end if
			%>
			<td width=300 align=center><font size=1><% = left(ors.Fields(3).Value,3) %> v. <a href=viewTeam.asp?team=<%=server.urlencode(enemy)%> onMouseOver="(window.status='View profile for <%=javaencode(enemy)%>'); return true" onMouseOut="(window.status='<%=javatitle%>'); return true"><%=server.htmlencode(enemy)%></a> (<%=pDate%>)
			<% if pdate <> "TBD" then
				response.write "<br>(" & server.htmlencode(map1) & ", " & server.htmlencode(map2) & ", " & server.htmlencode(map3) & ")"
			   end if
			  %>
			  </font></td></tr>
			<%
		else
			if (left(ors.Fields(3).Value,6) = "Defeat") or (left(ors.Fields(3).Value,6) = "Immune") or (left(ors.Fields(3).Value,6) = "Restin") then
				%>
				<td width=300 align=center><font size=1><%=ors.Fields(3).Value %></font></td></tr>
				<%
			else
			%>
			<td width=300 align=center><font size=1>Open</font></td></tr>
			<%
			end if
		end if
		ors.MoveNext
		if bgc=bgcone then
			bgc=bgctwo
		else
			bgc=bgcone
		end if
	loop
	%>
	</table>
	<%
end if

%>

        </td>
      </tr>
</table>
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
Set oRs1 = Nothing
%>