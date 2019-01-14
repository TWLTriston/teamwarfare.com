<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Record Change"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
Dim strTeamName , strLadderName
strTeamName = Request.QueryString("team")
strLadderName = Request.QueryString("Ladder")

Dim intWins, intLosses, intForfeits, intTLLinkID
strSQL = "SELECT TLLinkID, Wins, Losses, Forfeits FROM lnk_t_l lnk INNER JOIN tbl_teams t ON t.TeamID = lnk.TeamID INNER JOIN tbl_ladders l ON l.LadderID = lnk.LadderID WHERE TeamName = '" & CheckString(strTeamName) & "' AND LadderName='" & CheckString(strLadderName) & "'"
'Response.Write strSQL
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRs.BOF) THen
	intTLLinkID = oRs.FieldS("TLLinkID").Value
	intWins = oRs.FieldS("Wins").Value
	intLosses = oRs.FieldS("Losses").Value
	intForfeits = oRs.FieldS("Forfeits").Value
	
Else 
	Response.Clear
	Response.Write "Bad Data"
	Response.end
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<html>
<head>
	<title>TWL: Chage team record</title>
	<link REL=STYLESHEET HREF="/core/style.css" TYPE="text/css">
</head>

<body bgcolor="#000000" leftmargin="0" topmargin="00" marginwidth="000" marginheight="0000">
<TABLE height=100% width=100% border=0 cellspacing=0 cellpadding=0 valign=center align=center>
<tr valign=center>
	<td align="center">
	<form name="frmChangeRecord" id="frmChangeRecord" method="post" action="saveitem.asp">
	<input type="hidden" name="TLLinkID" id="TLLinkID" value="<%=intTLLinkID%>" />
	<input type="hidden" name="SaveType" id="SaveType" value="ChangeRecord" />
	<input type="hidden" name="Team" id="Team" value="<%=Server.HTMLEncode(strTeamName)%>" />
	<input type="hidden" name="Ladder" id="Ladder" value="<%=Server.HTMLEncode(strLadderName)%>" />
	
	<table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444">
	<tr>
		<td>
			<table border="0" cellspacing="1" cellpadding="4">
			<tr>
				<th bgcolor="#000000" colspan="2">Change Team Record</th>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right">Team:</td>
				<td bgcolor="<%=bgcone%>"><%=Server.HTMLEncode(strTeamName)%></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgctwo%>" align="right">Ladder:</td>
				<td bgcolor="<%=bgctwo%>"><%=Server.HTMLEncode(strLadderName)%></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right">Wins:</td>
				<td bgcolor="<%=bgcone%>"><input type="text" name="wins" id="wins" value="<%=intWins%>" size="5" maxlength="3" /></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgctwo%>" align="right">Losses:</td>
				<td bgcolor="<%=bgctwo%>"><input type="text" name="Losses" id="Losses" value="<%=intLosses%>" size="5" maxlength="3" /></td>
			</tr>
			<tr>
				<td bgcolor="<%=bgcone%>" align="right">Forfeits:</td>
				<td bgcolor="<%=bgcone%>"><input type="text" name="Forfeits" id="Forfeits" value="<%=intForfeits%>" size="5" maxlength="3" /></td>
			</tr>
			<tr>
				<td colspan="2" align="center" bgcolor="#000000"><input type="submit" value="Change Record" /></td>
			</tr>
			</table>
		</td>
	</tr>
	</table>
	</form>
	</td>
</tr>
</table>
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>