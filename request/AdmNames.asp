<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("player"), """", "&quot;") 

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack
Dim intCounter, intCount

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLadderAdmin, bLoggedIn, PlayerStatus
bSysAdmin = IsSysAdmin()
bLoggedIn = Session("LoggedIn")

%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Name Change Request Listing") 
if bSysAdmin then
	strSQL = "select tbl_PlayerNameChange.RequestID, tbl_PlayerNameChange.PlayerID, tbl_PlayerNameChange.OldName, tbl_PlayerNameChange.NewName, tbl_PlayerNameChange.RequestDate, tbl_Players.Suspension, tbl_Players.PlayerCanActivate from tbl_PlayerNameChange INNER JOIN tbl_Players ON tbl_PlayerNameChange.PlayerID = tbl_Players.PlayerID where Approved='0' order by OldName"
	oRs.Open strSQL, oConn
		If oRS.EOF and oRS.BOF Then
			%>
			<CENTER><b><font color=red>No Current Requests.</font></b></center>
			<%
		Else
			%>
			<table align=center border=0 CELLSPACING=0 cellpadding=0 class="cssBordered">
			<tr>
				<th><b>Player Name (Reqs)</b></th><th><b>Requested Name</b></th><th><b>Status</b></th></b></th><th>&nbsp;</th><th>&nbsp;</th><th><b>Request Date</b></th>
			</tr>
			<%
				bgc=bgctwo
				do while not oRS.eof
					strSQL = "select count(playerhandle) as cnt from tbl_players where playerhandle='" + CheckString(oRS.Fields("NewName")) + "'"
					oRS2.Open strSQL, oConn
					intCount = oRS2.Fields("cnt")
					oRS2.close 
					intCounter=0
					strSQL = "select count(playerID) as cnt from tbl_playernamechange where playerid=" & oRS.Fields("PlayerID")
					oRS2.Open strSQL, oConn
					intCounter = oRS2.Fields("cnt")
					oRS2.close
					Response.Write "<tr bgcolor=" & bgc & "><td><a href=/viewplayer.asp?player=" & Server.URLEncode(oRS.Fields("OldName")) & ">" & oRS.Fields("OldName") & "</a> (" & intCounter & ")</td>"
					Response.Write "<td>" & oRS.Fields("NewName") & "&nbsp;&nbsp;</td>"
					PlayerStatus = "Active"
					if (oRs.Fields("Suspension") = 1) Then
						PlayerStatus = "SUSPENDED"
					End if
					if (oRs.Fields("PlayerCanActivate") = 0)Then
						PlayerStatus = "BANNED"
					End If
					If Not (PlayerStatus = "Active") Then
						Response.Write "<td><font color=#ff0000><b>" & PlayerStatus & "</b></font></td>"
					Else
						Response.Write "<td>" & PlayerStatus & "</td>"
					End If
					if intCount = 0 then
						Response.Write "<td><a href=AdmNameApprove.asp?status=1&req=" & oRS.Fields("RequestID") & ">Approve</a>&nbsp;</td>"
					else 
						Response.Write "<td><font color=#ff0000><b>In Use</b></font></td>"
					end if
					Response.Write "<td><a href=AdmNameApprove.asp?status=9&req=" & oRS.Fields("RequestID") & ">Deny</a></td>"
					Response.Write "<td>" & oRS.Fields("RequestDate") & "</td></tr>"
					if bgc = bgcone then
						bgc=bgctwo
					else
						bgc=bgcone
					end if
					oRS.MoveNext
				loop
			%>		
			</table>
			<%
		end if
	else
			%>
			<CENTER><b><font color=red>You are not permitted to request a name change for users other than yourself.</font></b></center>
			<%
	end if
			
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Set oRS2 = Nothing
%>