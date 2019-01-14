<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add a League"

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
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("Add a League") %>
<FORM action="saveleague.asp" METHOD="post">
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER>
<TR><TD>
<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2>
<tr bgcolor=<%=bgcone%>><td align=right>League Name:</td><td width=300>&nbsp;<INPUT id=text1 name=LeagueName style=" WIDTH: 250px" class=text value=""></td></tr>
<tr bgcolor=<%=bgctwo%>><td align=right>Admin:</td><td>&nbsp;<INPUT id=text3 name=LeagueAdmin style=" WIDTH: 100px" class=text value=""></td></tr>
<tr bgcolor=<%=bgcone%>><td align=right>Game:</td><td>&nbsp;<INPUT id=text3 name=LeagueGame style=" WIDTH: 100px" class=text value=""></td></tr>
<tr bgcolor=<%=bgctwo%>><td align=right>Start Date:</td><td>&nbsp;<INPUT id=text3 name=LeagueStart style=" WIDTH: 100px" class=text value=""></td></tr>
<tr bgcolor=<%=bgcone%>><td align=right>End Date:</td><td>&nbsp;<INPUT id=text3 name=LeagueEnd style=" WIDTH: 100px" class=text value=""></td></tr>
<tr bgcolor=<%=bgctwo%>><td align=right>Division 1:</td><td>&nbsp;<INPUT id=text3 name=LeagueDiv_1 style=" WIDTH: 200px" class=text value=""></td></tr>
<tr bgcolor=<%=bgcone%>><td align=right>Division 2:</td><td>&nbsp;<INPUT id=text3 name=LeagueDiv_2 style=" WIDTH: 200px" class=text value=""></td></tr>
<tr bgcolor=<%=bgctwo%>><td align=right>Division 3:</td><td>&nbsp;<INPUT id=text3 name=LeagueDiv_3 style=" WIDTH: 200px" class=text value=""></td></tr>
<tr bgcolor=<%=bgcone%>><td align=right>Division 4:</td><td>&nbsp;<INPUT id=text3 name=LeagueDiv_4 style=" WIDTH: 200px" class=text value=""></td></tr>
<tr bgcolor=<%=bgctwo%>><td align=right>Division 5:</td><td>&nbsp;<INPUT id=text3 name=LeagueDiv_5 style=" WIDTH: 200px" class=text value=""></td></tr>
<tr bgcolor=<%=bgcone%>><td align=right>Division 6:</td><td>&nbsp;<INPUT id=text3 name=LeagueDiv_6 style=" WIDTH: 200px" class=text value=""></td></tr>
<tr bgcolor=<%=bgctwo%>><td align=right>Division 7:</td><td>&nbsp;<INPUT id=text3 name=LeagueDiv_7 style=" WIDTH: 200px" class=text value=""></td></tr>
<tr bgcolor=<%=bgcone%>><td align=right>Division 8:</td><td>&nbsp;<INPUT id=text3 name=LeagueDiv_8 style=" WIDTH: 200px" class=text value=""></td></tr>
<tr bgcolor=<%=bgctwo%>><td align=right>Teams per Division:</td><td>&nbsp;<INPUT id=text3 name=LeagueDivSize style=" WIDTH: 20px" class=text value=""></td></tr>
<tr bgcolor=<%=bgcone%>><td colspan=2 align=middle><INPUT id=submit1 name=submit1 type=submit value=Save class=text></td></tr>
<input type=hidden name=SaveMethod value="addleague">
</TABLE>
</TD></TR>
</tABLE>
<input type=hidden name=SaveType value=addleague>
</form>

<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>