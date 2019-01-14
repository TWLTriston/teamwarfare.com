<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Match Extras"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<%
Dim ReturnURL, m3Info, m1Side, m2Side, m3side, matchID, cnt
returnURL=request("returnURL")

if request("laddertype") = "MW4" then
	m3info="Visibility=" & request("mvis") & "<br>Drop Weight Adjustment=0<br>Time of Day=" & request("mToD") & "<br>Drop Zone Selection=Defender/Challenger"
	strSQL="update tbl_map_extras set Map3Extra='" & m3info & "' where matchid=" & Request("matchid")
else
	m1side=Request("m1side")
	m2side=Request("m2side")
	m3side=request("m3side")
	matchid=Request("matchid")
	'Response.write "The defender chose the " & m1side & " side for the first map.<br>The defender chose the " & m2side & " side for the Second map.<br>Return to " & returnURL & "<br>Match ID: " & matchid
	strSQL="select count(*) from tbl_map_extras where matchid=" & matchid
	ors.Open strSQL, oconn
	cnt=ors.Fields(0).Value 
	ors.Close 
	if cnt=0 then
		if m3side="" then
			strSQL="insert into tbl_map_extras (MatchID, Map1Extra, Map2Extra) values ('" & matchid & "','" & m1side & "','" & m2side & "')"
		else
			strSQL="insert into tbl_map_extras (MatchID, Map3Extra) values ('" & matchid & "','" & m3side & "')"
		end if
	else
		if m3side="" then
			strSQL="update tbl_map_extras set Map1Extra='" & m1side & "', Map2Extra='" & m2side & "' where matchid=" & matchid
		else
			strSQL="update tbl_map_extras set Map3Extra='" & m3side & "' where matchid=" & matchid
		end if
	end if
end if
'Response.Write server.htmlencode(strSQL)
oconn.Execute strsql

oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Response.clear 
Response.Redirect returnurl
Response.End 
%>	
