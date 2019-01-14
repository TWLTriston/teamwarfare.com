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

<% Call Content2BoxStart("Find Player By Name") %>
<%
Dim strSearch
strSearch = trim(Request("item"))
Dim FoundOne
%>
	
<form method=post action="" name=searchscript id=1>
<table border="0" cellspacing="0" cellpadding="0" align="center" class="cssBordered">
<tr>
	<th>New Search</th>
</tr>
<tr>
	<td bgcolor="<%=bgcone%>"><input type=text name=item class=bright id=searchitem size=15 style="width:200px; height:18px;"></td>
</tr>
<tr>
	<td bgcolor="#000000" align="center"><input type=submit value="New Search" name=submitsearch id=submitsearch class=bright></td>
</tr>
</table>
</form>

<% Call Content2BoxMiddle() %>
	<% if strSearch="" then %>
	Please use the form on the left to perform a search
	<% else %>
	Previous search:<br /><b><%=strSearch%></b>
	<% end if %>
<% Call Content2BoxEnd() %>
<%
if strSearch <> "" then
	Call ContentStart("Search Results")
	%>
 <table width="100%">

<%
	foundone=false

	bgc=bgcone
	strsql="Select TOP 50 * from tbl_players where playerhandle like '%" & CheckString(SearchString(strSearch)) & "%' order by playerhandle"
	ors.Open strsql, oconn
	if not(ors.EOF and ors.BOF) then
		foundone=true
		Response.Write "<tr><td>&nbsp;</td></tr>"
		Response.Write "<tr><td><b>The following player(s) were found: (by name)</b></td></tr>"
		do while not (ors.EOF)
			Response.Write "<tr bgcolor=" & bgc & " height=20><td>&nbsp;<a href=viewplayer.asp?player=" & server.URLEncode(ors.Fields(1).Value) & ">" & ors.Fields(1).Value & "</a></td></tr>"
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