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

If Not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.clear
	Response.Redirect "errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call Content2BoxStart("Email Search") %>
<%
Dim searchStr
searchStr = Request("email")
Dim FoundOne
%>
	<table width=780 border="0" cellspacing="0" cellpadding="0" BACKGROUND="">
	<tr>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	<td width=380 ALIGN=CENTER>
	
<form method=post action=emailsearch.asp name=searchscript id=searchscript>
<font face=Arial size=2 color=white><b>Enter the exact email address below:</b></font><br>
&nbsp;&nbsp;	<input type=text name=email class=bright id=email size=15 style="width:200px; height:18px;"><br><br>&nbsp;&nbsp;<input type=submit value="New Search" name=submitsearch id=submitsearch class=bright><br>
</form>

	</td>
	<td><img src="/images/spacer.gif" width="10" height="1"></td>
	<td width=379 ALIGN=CENTER>
	<% if searchStr="" then %>
	Please use the form on the left to perform a search
	<% else %>
	Previous search:<br><b><%=searchStr%></b>
	<% end if %>
	</td>
	<td><img src="/images/spacer.gif" width="5" height="1"></td>
	</tr>
	</table>

<% Call Content2BoxEnd() %>
<%
if searchStr <> "" then
	Call ContentStart("Search Results")
	%>
        <table border=0 width=97% cellspacing=0 cellpadding=0 align=center>

<%
	foundone=false	
	bgc=bgcone
	strsql="Select TOP 50 * from tbl_players where playeremail = '" & CheckString(searchStr) & "' order by playerhandle"
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