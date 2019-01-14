<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Search"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call Content2BoxStart("Find Player By In Game Identifier") %>
<%
Dim strSearch
strSearch = trim(Request("item"))
Dim FoundOne
%>

		<table border="0" cellspacing="0" cellpadding="0" class="cssBordered" align="center">
		<form method=post action="" name=searchscript id=1>
				<tr>
					<th bgcolor="#000000" colspan="2">New Search</th>
				</tr>
				<tr>
					<td align="right" bgcolor="<%=bgcone%>"><b>Identifier Type</b></td>
					<td bgcolor="<%=bgcone%>">
						<select name="selIdentifierID" id="selIdentifierID">
						<%
						strSQL = "SELECT IdentifierName, IdentifierID FROM tbl_identifiers ORDER BY IdentifierName ASC "
						oRs.Open strSQL, oConn
						If Not(oRs.EOF AND oRs.BOF) Then
							Do While Not(oRs.EOF)
								Response.Write "<option value=""" & oRs.Fields("IdentifierID").Value & """"
								If CStr(oRs.Fields("IdentifierID").Value & "") = CStr(Request.Form("selIdentifierID")) Then
									Response.Write " selected=""selected"""
								End If
								Response.Write ">" & Server.HTMLEncode(oRs.Fields("IdentifierName").Value & "") & "</option>" & VbCrLf
								oRs.MoveNext
							Loop
						End if
						oRs.NextRecordSet
						%>
						</select>
					</td>
				</tr>
				<tr>
					<td bgcolor="<%=bgcone%>" align="right"><b>Relevant Value:</b></td>
					<td bgcolor="<%=bgctwo%>"><input type="text" name="txtIdentifierValue" id="txtIdentifierValue" size="20" value="<%=Request.Form("txtIdentifierValue")%>" /></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#000000" align="center"><input type="submit" value="Perform Search" /></td>
			</tr>
		</form>
		</table>

<% Call Content2BoxMiddle() %>
<div align="center">
	<% if strSearch="" then %>
	Please use the form on the left to perform a search
	<% else %>
	Previous search:<br><b><%=strSearch%></b>
	<% end if %>
</div>
<% Call Content2BoxEnd() %>
<%
if Request.Form("txtIdentifierValue") <> "" then
	Call ContentStart("Search Results")
	%>
        <table border=0 width=97% cellspacing=0 cellpadding=0 align=center>

<%
	foundone=false
	bgc=bgcone
	if bSysAdmin Then
	strsql="SELECT PlayerHandle, Identifiervalue FROM tbl_players p INNER JOIN lnk_player_identifier i ON p.PlayerID = i.PlayerID WHERE IdentifierID = '" & CheckString(Request.Form("selIdentifierID")) & "' AND IdentifierValue LIKE '%" & CheckString(SearchString(Request.Form("txtIdentifierValue"))) & "%' order by PlayerHandle"
	else
	strsql="SELECT PlayerHandle, Identifiervalue FROM tbl_players p INNER JOIN lnk_player_identifier i ON p.PlayerID = i.PlayerID WHERE IdentifierID = '" & CheckString(Request.Form("selIdentifierID")) & "' AND IdentifierValue LIKE '%" & CheckString(SearchString(Request.Form("txtIdentifierValue"))) & "%' AND IdentifierActive = 1 order by PlayerHandle"
	end if
	ors.Open strsql, oconn
	if not(ors.EOF and ors.BOF) then
		foundone=true
		Response.Write "<tr><td>&nbsp;</td></tr>"
		Response.Write "<tr><td><b>The following player(s) were found: </b></td></tr>"
		do while not (ors.EOF)
			Response.Write "<tr bgcolor=" & bgc & " height=20><td>&nbsp;<a href=viewplayer.asp?player=" & server.URLEncode(ors.Fields("PlayerHandle").Value) & ">" & Server.HTMLEncode(ors.Fields("PlayerHandle").Value) & "</a> (" & oRs.Fields("Identifiervalue").Value & ")</td></tr>"
			if bgc=bgcone then
				bgc=bgctwo
			else
				bgc=bgcone
			end if
			ors.MoveNext
		loop
	end if
	ors.Close

	if not(foundone) then
		Response.Write "<tr align=center><td><b>No data was found matching the search criteria you specified.</td></tr>"
	end if
%>
	</table>
	<%
	Call ContentEnd()
end if
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>