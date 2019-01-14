<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: " & Replace(Request.Querystring("player"), """", "&quot;") 

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bLadderAdmin, bLoggedIn
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
bLoggedIn = Session("LoggedIn")

Dim strPlayerName, intPlayerID, Date90Day, NumberofRequests
strPlayerName = Request("Player")
if Len(strPlayerName) = 0 Then
	strPlayerName = Session("uName")
End if

Dim strTeamName, test1, test2, test3
Dim TMLinkID, DivID, TournamentName
Dim bBarDone
Dim bCanAdmin
Dim strLadderName, intRank, intLosses, intPlayerLadderID
Dim intForfeits, intWins, strStatus, strEnemyName, strResult
Dim linkID, map, opponent, mDate, statusVerbage
bBarDone = False
%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<script>
function fQuickSearch() {
	if (document.frmRequest.txtNewName.value.length != 0) {
		window.location.href = "/searchPlayerByName.asp?item=" + document.frmRequest.txtNewName.value;
	}
}
</script>
<%
Call ContentStart("Player Name Change Request") 
	If session("uName") = strPlayerName  then
		strSQL = "select PlayerHandle, PlayerID from tbl_Players where PlayerHandle='" & CheckString(strPlayerName) & "'"
		oRs.Open strSQL, oConn
		If oRS.EOF and oRS.BOF Then
			%>
			<CENTER><b><font color=red>Member not found... please check your URL.</font></b></center>
			<%
		Else
		  intPlayerID = ors.Fields("PlayerID").Value
		  Date90Day = DateAdd("d", -91, Now)
		  ' Look for duplicate requests
		  strSQL = "SELECT PlayerID, RequestDate FROM tbl_PlayerNameChange WHERE PlayerID = '" & intPlayerID & "' AND Approved = '0'"
		  oRs2.Open strSQL, oConn,3
		  If Not(oRs2.EOF AND oRs2.BOF) Then
		    %>
			<CENTER><b><font color=red>You already have a pending name change request. Please wait 24-48 hours for processing before attempting again.</font></b></center>
			<%
		  Else
		  oRs2.Close
		  ' end dupe search
		  strSQL = "SELECT PlayerID, RequestDate FROM tbl_PlayerNameChange WHERE PlayerID = '" & intPlayerID & "' AND RequestDate > '" & Date90Day & "'"
		  oRs2.Open strSQL, oConn,3
		  NumberofRequests = oRs2.RecordCount
		  If oRs2.RecordCount > 5 Then
		    %>
			<CENTER><b><font color=red>Too many name change requests. You are permitted 5 in a 90 day period. You currently have <%=Server.HTMLEncode(NumberofRequests)%>.</font></b></center>
			<%
		  Else
			%>
			<table align=center border=0 CELLSPACING=1 cellpadding=2 width="500">
			<tr><td>
			Not all name requests will be granted.  <br />
			<br />
			Requested names must be unique.  The system will verify that your requested name is available before submitting the request.
			This validation will happen again when the administrators approve the name change request.  It is possible that your requested 
			name will be taken between the time of request and the approval. You are permitted a maximum of 5 name changes in 90 days. Any more
			will be automatically rejected. You currently have <%=Server.HTMLEncode(NumberofRequests)%> requests in the last 90 days.
			<br />
			<br />
			It is also possible that an administrator will deny a name change request for reasons such as
			<ul>
			<li>Vulgar or inappropriate names</li>
			<li>Abuse of the change request system (excessive requests will be denied)</li>
			</ul>
			If you would like to verify the availability of your requested name click the "quick search" link.
			<br />
			<br />
			Name changes are processed manually within 24-48 hours.
			</td><tr></table>
			<br />
			<br />
			
			
			<form name="frmRequest" action="ReqCheckName.asp" method="get">
			<% if Request("e") = 1 then %><center><font color=#ff0000>Requested name is in use</font></center><% end if %>
				<table align="center" border="0" cellpadding="0" cellspacing="0" width="50%" class="cssBordered">
				
				<TR BGCOLOR="#000000">
					<TH COLSPAN=2>Current User: <%=Server.HTMLEncode(strPlayerName)%></TH>
				</TR>
				<%
				intPlayerID = ors.Fields("PlayerID").Value 
				strPlayerName = oRs.Fields("PlayerHandle").Value 
				%>
				<tr bgcolor=<%=bgctwo%>>
					<td>Requested Name</td>
					<td>
						<input type="text" name="txtNewName" id="txtNewName" />
						<input type="hidden" name="hdnPlayerID" value="<%=intPlayerID%>" />	
						&nbsp;&nbsp;<input type="submit" value="Submit Request" />
						&nbsp;&nbsp;<a href="javascript:fQuickSearch();">Quick Search</a>
						<br /></td>
				</tr>
				</table>
			<%
			end if
			end if
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