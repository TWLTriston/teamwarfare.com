<% Option Explicit %>
<%
Response.Buffer = True

Const adOpenForwardOnly = 0
Const adLockReadOnly = 1
Const adCmdTableDirect = &H0200
Const adUseClient = 3

Dim strPageTitle, intDisputeTeamID
intDisputeTeamID = Request.QueryString("DisputeTeamID")

strPageTitle = "TWL: Dispute Match" 

Dim strSQL, oConn, oRs, oRs2
Dim bgcone, bgctwo, strHeaderColor
strHeaderColor	= Application("HeaderColor")

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRs = Server.CreateObject("ADODB.RecordSet")
Set oRs2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLeagueAdmin, bLoggedIn, bAnyLadderAdmin
bLoggedIn = Session("LoggedIn")
bSysAdmin = IsSysAdmin()
bAnyLeagueAdmin = IsAnyLeagueAdmin()
bAnyLadderAdmin =  IsAnyLadderAdmin()

Dim strLeagueName, intLeagueID, intGameID, intDisputeForumID

strLeagueName = Request.QueryString("League")
If Len(Trim(strLeagueName)) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If

Dim intMatchID
intMatchID = Trim(Request.QueryString("MatchID"))
If Not(IsNumeric(intMatchID)) OR Len(intMatchID) = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
	
strSQL = "SELECT LeagueID, LeagueGameID, LeagueName FROM tbl_leagues WHERE LeagueName = '" & CheckString(strLeagueName) & "' AND LeagueActive = 1"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intLeagueID = oRs.Fields("LeagueID").Value
	strLeagueName = oRs.Fields("LeagueName").Value
	intGameID = oRs.FieldS("LeagueGameID").Value
Else
	oRs.Close
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=7"
End If
oRs.NextRecordSet

strSQL = "SELECT DisputeForumID FROM tbl_games WHERE GameID = '" & CheckString(intGameID) & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) Then
	intDisputeForumID = oRs.FieldS("DisputeForumID").Value
End If
oRs.NextRecordSet

'' Get match details
Dim strHomeTeam, strVisitorTeam, strHomeTeamTags, strVisitorTeamTags, intHomeTeamID, intVisitorTeamID
DIm intHomeLinkID, intVisitorLinkID, strMaps(6), i

strSQL = "EXECUTE LeagueMatchDetails @LeagueMatchID = '" & CheckString(intMatchID) & "'"
oRs.Open strSQL, oConn
If oRs.State = 1 Then
	If Not(oRs.EOF AND oRs.BOF) Then
		strHomeTeam = oRs.Fields("HomeTeamName").Value
		strVisitorTeam = oRs.Fields("VisitorTeamName").Value
		strHomeTeamTags = oRs.Fields("HomeTeamTag").Value
		strVisitorTeamTags = oRs.Fields("VisitorTeamTag").Value
		intHomeTeamID = oRs.Fields("HomeTeamID").Value
		intVisitorTeamID = oRs.Fields("VisitorTeamID").Value
		intHomeLinkID = oRs.Fields("HomeTeamLinkID").Value
		intVisitorLinkID = oRs.Fields("VisitorTeamLinkID").Value
	End If
	oRs.nextRecordSet
End if

If intDisputeTeamID <> CStr(intHomeLinkID  & "") AND intDisputeTeamID <> cStr(intVisitorLinkID & "") Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=32"
End If

'' Verify access to dispute
If Not(bSysAdmin or IsLeagueAdmin(strLeagueName) or _
	IsTeamFounder(strHomeTeam) Or IsTeamFounder(strVisitorTeam) Or _
	IsLeagueTeamCaptain(strHomeTeam, strLeagueName) Or IsLeagueTeamCaptain(strHomeTeam, strLeagueName)) Then 
	
	Response.Clear
	Response.Redirect "errorpage.asp?error=3"
End If

Dim bgc
Dim intTimeZoneDifference, strDate, strTime
Dim strCurrentTime, strCurrentDate
Dim strDateMask, bln24HourTime

intTimeZoneDifference = Session("intTimeZoneDifference")
strDateMask = "MM-DD-YYYY"
bln24HourTime = False

Dim strYourTeam, strTheirTeam
'' Identify disputing team
If cStr(intHomeLinkID) = intDisputeTeamID Then 
	strTheirTeam = strVisitorTeam
	strYourTeam = strHomeTeam
Else
	strTheirTeam = strHomeTeam
	strYourTeam = strVisitorTeam
End if

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart(strLeagueName & " Ladder Match Dispute Form")
%>	
<script language="javascript" type="text/javascript">
<!--
function fSubmit(objForm) {
	var errFlag = 0;
	var errMsg = "Error:\n";
	if (objForm.DisputeReason.value.length == 0) {
		errFlag = 1;
		errMsg = errMsg + "You must choose a reason for the dispute.\n";
	} 
	if (objForm.Details.value.length < 100) {
		errFlag = 1;
		errMsg = errMsg + "Details must be provided. Please elaborate on your dispute, use at minimum of 100 characters.\n";
	}
	if (errFlag == 0) {
		if (confirm("Are you certain you want to submit this dispute? \nDid you put all the information possible in the details?\nIf yes, click ok, otherwise, click cancel.")) {
			objForm.submit();
		} else {
			// nothing
		}
	} else {
		alert(errMsg);
	}
}
//-->
</script>

	Use the form below to dispute a match. <br />
	<br />
	<table border="0" align="center" width="52%"><tr><td>
	Fill in with as much detail
	as possible regarding your complaint against the other team
	so that we may faciliate your claim as fast as possible.
	Please allow 24 hours for a response from your ladder admin. If no
	response is received, you can follow up via <a href="staff.asp">email</a>, or irc.<br /><br />
	<center><b>Do not submit this form more than once. It will only delay the processing of the dispute.</b></center>
	</td></tr>
	</table>
	<form name="frmDispute" id="frmDispute" action="saveitem.asp" method="post">
	<input type="hidden" id="hdnMatchID" name="hdnMatchID" value="<%=intMatchID%>" />
	<input type="hidden" id="hdnGameID" name="hdnGameID" value="<%=intGameID%>" />
	<input type="hidden" id="hdnDisputeForumID" name="hdnDisputeForumID" value="<%=intDisputeForumID%>" />
	<input type="hidden" id="SaveType" name="SaveType" value="MatchDispute" />
	<input type="hidden" id="hdnCompetitionType" name="hdnCompetitionType" value="League" />
	<input type="hidden" id="hdnSubmittor" name="hdnSubmittor" value="<%=Server.HTMLEncode("" & Session("uname"))%>" />
	<input type="hidden" id="hdnSubmittorID" name="hdnSubmittorID" value="<%=Session("UserID")%>" />
	<input type="hidden" id="hdnLeagueName" name="hdnLeagueName" value="<%=Server.HTMLEncode("" & strLeagueName)%>" />
	<input type="hidden" id="hdnLadderID" name="hdnLadderID" value="<%=Server.HTMLEncode("" & intLeagueID)%>" />
	<input type="hidden" id="hdnDisputingTeam" name="hdnDisputingTeam" value="<%=Server.HTMLEncode("" & strYourTeam)%>" />
	<input type="hidden" id="hdnDisputedTeam" name="hdnDisputedTeam" value="<%=Server.HTMLEncode("" & strTheirTeam)%>" />
	<table width="60%" border="0" cellspacing="0" cellpadding="0" BACKGROUND="" bgcolor="#444444">
	<tr><td>
	<table width="100%" border="0" cellspacing="1" cellpadding="4" BACKGROUND="">
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right">Your name</td>
		<td bgcolor="<%=bgctwo%>"><b><%=Server.HTMLEncode("" & Session("uName"))%></b></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right">Ladder</td>
		<td bgcolor="<%=bgcone%>"><b><%=Server.HTMLEncode("" & strLeagueName)%></b></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right">Disputing team</td>
		<td bgcolor="<%=bgctwo%>"><b><%=Server.HTMLEncode("" & strYourTeam)%></b></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right">Disputing against</td>
		<td bgcolor="<%=bgcone%>"><b><%=Server.HTMLEncode("" & strTheirTeam)%></b></td>
	</tr>
	<tr>
		<td bgcolor="<%=bgctwo%>" align="right">Dispute reason</td>
		<td bgcolor="<%=bgctwo%>"><select name="DisputeReason" id="DisputeReason">
			<option value="">&lt;--   Choose One   --&gt;</option>
			<option value="Bug Exploit">Bug Exploit</option>
			<option value="Cheating">Cheating</option>
			<option value="Invalid Server">Invalid Server</option>
			<option value="No Show">No Show</option>
			<option value="Rules Violation">Rules Violation</option>
			<option value="Smurfing">Smurfing</option>
			<option value="Unsportsmanlike Conduct">Unsportsmanlike Conduct</option>
			</select>
		</td>
	</tr>
	<tr>
		<td align="right" bgcolor="<%=bgcone%>" valign="top">Details<br />
		<span style="font-size: 9px;">(be as specific as possible, must be at least 100 characters)</span>
		</td>
		<td bgcolor="<%=bgcone%>"><textarea id="Details" name="Details" cols="40" rows="8"></textarea></td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#000000" align="center"><input onclick="fSubmit(this.form);" type="button" value="Submit Dispute" /></td>
	</tr>
	</table>
	</td></tr>
	</table>
	</form>

<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing

Function ListPages(byVal iPageNum, byVal iTotalPages)
	Dim i
	If iTotalPages > 1 Then
		Response.Write "<TR><TD><IMG SRC=""/images/spacer.gif"" HEIGHT=5></TD></TR>"
		Response.Write "<TR>"
		Response.Write "<TD CLASS=""pagelist"">"
		Response.Write "Pages (" & iTotalPages & "): <B>"
		If iPageNum > 5 Then
			Response.Write " <a alt=""First Page"" href=""viewmatch.asp?ladder=" & Server.URLENcode(strLeagueName & "") & "&matchid=" & intMatchID & "&page=1"">&laquo; First</A> ... "
		End If
		If iPageNum > 1 Then
			Response.Write " <a alt=""Previous Page"" href=""viewmatch.asp?ladder=" & Server.URLENcode(strLeagueName & "") & "&matchid=" & intMatchID &"&page=" & iPageNum - 1 & """>&laquo;</A> "
		End If
		For i = iPageNum - 5 To iPageNum + 5 
			If i > 0 Then
				If i = iPageNum Then
					Response.Write " <span class=""currentpage"">[" & i & "]</span>"
				ElseIf i <= iTotalPages Then
					Response.Write " <a href=""viewmatch.asp?ladder=" & Server.URLENcode(strLeagueName & "") & "&matchid=" & intMatchID & "&page=" & i & """>" & i & "</a>"
				End If				
			End If
		Next
		If iPageNum < iTotalPages Then
			Response.Write " <a alt=""Next Page"" href=""viewmatch.asp?ladder=" & Server.URLENcode(strLeagueName & "") & "&matchid=" & intMatchID & "&page=" & iPageNum + 1 & """>&raquo;</A> "
		End If
		If iPageNum + 5 < iTotalPages Then
			Response.Write " ... <a alt=""Last Page"" href=""viewmatch.asp?ladder=" & Server.URLENcode(strLeagueName & "") & "&matchid=" & intMatchID & "&page=" & iTotalpages & """>Last &raquo;</A>"
		End If
		Response.Write "</B>"
		Response.Write "</TD></TR>"
		Response.Write "<TR><TD><IMG SRC=""/images/spacer.gif"" HEIGHT=5></TD></TR>"
	End If
End Function
%>