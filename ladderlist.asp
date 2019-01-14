<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Competition list"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim intGameID
intGameID = -1

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call Content33BoxStart("Competition List") 
If Request.QueryString("error") = "1" Then
	Response.Write "<font color=""#ff0000"">Invalid ladder name. Please choose from a ladder below.</font><br />"
End If
Dim strType
Dim intCompetitions, intActiveTeams
intActiveTeams = 0
intCompetitions = 0
strType = ""
strSQL = "EXECUTE CompetitionList"
oRs.Open strSQL, oConn
'' build the game ids
%>

<table border="0" cellspacing="0" cellpadding="0" class="cssBordered" align="center" width="97%">
<tr>
	<th colspan="2">Choose a Game</th>
</tr>
<%
Dim intThisGame, arrGameIDs(50), iGameCounter
intThisGame = -1
iGameCounter = 0
Do While Not(oRs.EOF)
	If intThisGame <> oRs.Fields("GameID").Value Then
		intThisGame = oRs.Fields("GameID").Value
		arrGameIDs(iGameCounter) = intThisGame
		iGameCounter = iGameCounter + 1
		%>
		<tr>
			<td style="background:<%=bgcone%>;cursor: pointer;" onMouseOver="this.style.background='<%=bgctwo%>';" onMouseOut="this.style.background='<%=bgcone%>';" onClick="fShowGame(<%=oRs.Fields("GameID").Value%>)">
				<a href="javascript:fShowGame(<%=oRs.Fields("GameID").Value%>)"><%=Server.HTMLEncode(oRs.Fields("GameName").Value & "")%></a>
			</td>
		</tr>
		<%
	End If
	oRs.MoveNext
Loop
%>
</table>
<script language="javascript" type="text/javascript">
var arrGameIDs = new Array();
<% 
Dim I
For i = 0 to uBound(arrGameIDs)
	If Len(arrGameIDs(i)) > 0 Then
	%>
	arrGameIDs.push(<%=arrGameIDs(i)%>);<%
	End If
Next
%>
function fShowGame(intGameID) {
	for (i=0;i<arrGameIDs.length;i++) {
		strDivName = "divGame" + arrGameIDs[i]; //intGameID;
		if (arrGameIDs[i] == intGameID) {
			document.getElementById(strDivName).style.visibility = "visible";
			document.getElementById(strDivName).style.display = "inline";
		} else {
			document.getElementById(strDivName).style.visibility = "hidden";
			document.getElementById(strDivName).style.display = "none";
		}
	}
}
</script>

<% Call Content33BoxMiddle() %>

<%
Set oRs = oRs.NextRecordSet
bgc=bgctwo
if not (ors.EOF and ors.BOF) then
		do while not ors.EOF
			intCompetitions = intCompetitions + 1
			intActiveTeams = intActiveTeams + oRs.Fields("ActiveTeams").Value
			if intGameID <> oRS.Fields("GameID") Then
				strType = oRs.fields("LadderType").Value
				If intGameID <> -1 Then
					Response.Write "</TABLE><BR><BR></div>"
				End If
				intGameID = oRS.Fields("GameID").Value
				%>
				<div id="divGame<%=intGameID%>" style="visibility: hidden; display: none;">
				<a name="Game<%=intGameID%>"></a>
				<table border="0" cellspacing="0" cellpadding="0" WIDTH=98% class="cssBordered" align="center">
				<TR BGCOLOR="#000000">
					<TH COLSPAN=<% If bSysAdmin Then Response.write "4" Else Response.Write "3" End If %>><%=oRS.Fields("GameName").Value%> ( <%=oRs.Fields("GameAbbreviation").Value%> )</TH>
				</TR>
				<TR BGCOLOR="#000000">
					<TH>Name</TH>
					<TH WIDTH=100>Active Teams</TH>
					<% If bSysAdmin Then %>
					<TH WIDTH=50>Edit</TH>
					<% End If %>
					<TH WIDTH=50>Info</TH>
				</TR>
				<%
			End If
			If strType <> oRs.Fields("ladderType").Value THen
				%>
				<tr><td colspan="<% If bSysAdmin Then Response.write "4" Else Response.Write "3" End If %>" bgcolor="#000000"><img src="/images/spacer.gif" height="3" width="1" alt="" border="0" /></td></tr>
				<%
				strType = oRs.Fields("ladderType").Value
			End If
			%>
			<tr bgcolor=<%=bgc%> ><td>
			<% 
			If oRS.Fields ("LadderType").Value = "T" Then
				%>
				<a href=viewladder.asp?ladder=<% Response.Write server.urlencode(ors.Fields("LadderName").Value) %> ><% =Server.HTMLEncode(ors.Fields("LadderName").Value) %> Ladder</a> 
				<%
			ElseIf oRS.Fields ("LadderType").Value = "U" Then
				%>
				<a href=viewscrimladder.asp?ladder=<% Response.Write server.urlencode(ors.Fields("LadderName").Value) %> ><% =Server.HTMLEncode(ors.Fields("LadderName").Value) %> Ladder</a> 
				<%
			ElseIf oRS.Fields ("LadderType").Value = "P" Then
				%>
				<a href=viewPlayerladder.asp?ladder=<% Response.Write server.urlencode(ors.Fields("LadderName").Value) %> ><% =Server.HTMLEncode(ors.Fields("LadderName").Value) %> Ladder</a>
				<%
			ElseIf oRS.Fields ("LadderType").Value = "L" Then
				%>
				<a href=viewleague.asp?league=<% Response.Write server.urlencode(ors.Fields("LadderName").Value) %> ><% =Server.HTMLEncode(ors.Fields("LadderName").Value) %> League</a>
				<%
			ElseIf oRS.Fields ("LadderType").Value = "A" Then
				%>
				<a href=tournament/default.asp?tournament=<% Response.Write server.urlencode(ors.Fields("LadderName").Value) %> ><% =Server.HTMLEncode(ors.Fields("LadderName").Value) %> Tournament</a>
				<%
			End If
			%>
			</TD>
			<td align=center><%=oRs.Fields("ActiveTeams").Value %></td>
			<% If bSysAdmin Then
				If oRS.Fields ("LadderType").Value = "T" Then
					%>
					<TD ALIGN=CENTER><A href="/addladder.asp?IsEdit=true&ladder=<%=server.URLEncode(oRs.Fields("LadderName").Value)%>">Edit</A></TD>
					<%
				ElseIf oRS.Fields ("LadderType").Value = "U" Then
					%>
					<TD ALIGN=CENTER><A href="/scrim/LadderAdd.asp?IsEdit=true&ladder=<%=server.URLEncode(oRs.Fields("LadderName").Value)%>">Edit</A></TD>
					<%
				ElseIf oRS.Fields ("LadderType").Value = "P" Then
					%>
					<TD ALIGN=CENTER><A href="/addplayerladder.asp?IsEdit=true&name=<%=server.URLEncode(oRs.Fields("LadderName").Value)%>">Edit</A></TD>
					<%
				ElseIf oRS.Fields ("LadderType").Value = "L" Then
					%>
					<TD ALIGN=CENTER>&nbsp;</TD>
					<%
				ElseIf oRS.Fields ("LadderType").Value = "A" Then
					%>
					<TD ALIGN=CENTER>&nbsp;</TD>
					<%
				End If
			End if %>
			<% If Ors.fields("LadderType").Value = "T" Then %>
			<td align="center"><a href="viewladderdetails.asp?ladder=<% Response.Write server.urlencode(ors.Fields("LadderName").Value) %>">info</a></td>
			<% Else %>
			<td>&nbsp;</td>
			<% ENd If %>
			
			</tr>
			<%
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.MoveNext
		loop
	end if
	
%>
</table><br /><br />
</div>
<%
	If Session("uName") = "Triston" Then
		%>
		&nbsp;&nbsp;Competitions: <%=intCompetitions%> <br />
		&nbsp;&nbsp;Active Teams: <%=intActiveTeams%> <br />
		
		<%
	End If
%>
<% Call Content33BoxEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

