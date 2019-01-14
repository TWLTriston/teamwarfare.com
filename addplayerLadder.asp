<%' Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add a Player Ladder"

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
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim strVerbage, bIsEdit, strLadderName, strMethod, intGameID
Dim lAdmin, lGame, lActive, lid, labbr, llocked, lchallenge
bIsEdit = cBool(Request.QueryString("IsEdit"))
If bIsEdit Then
	strVerbage = "Edit a player ladder"
	strLadderName = Request.QueryString("name")
	strMethod = "Edit"
	strSQL = "select * from tbl_playerladders where Playerladdername='" & CheckString(strLaddername) & "'"
	oRs.Open strSQL, oConn
	if not (ors.eof and ors.BOF) then
		lactive = ors.Fields("Active").Value 
		lid = ors.Fields("PlayerLadderID").Value 
		labbr = ors.Fields("Abbreviation").Value 
		llocked = ors.Fields("Locked").Value 
		lchallenge = ors.Fields("Challenge").Value 
		intGameID = oRS.Fields("GameID").Value
		sMethod="Edit"
	end if
	ors.Close
Else
	strVerbage = "Add a player ladder"
	strMethod = "New"
End If
strPageTitle = strVerbage
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart(strVerbage)
%>
	<form name=frmAddPlayerLadder action=saveItem.asp method=post>
	<table align=center BACKGROUND="" BGCOLOR="#444444" CELLSPACING=0 CELLPADDING=0>
	<TR><TD>
	<table align=center CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
		<tr bgcolor=<%=bgcone%>><td align=right>Name:</td><td width=300>&nbsp;<INPUT id=LadderName name=PlayerLadderName style=" WIDTH: 250px" class=text value="<%=Server.HTMLEncode(strLadderName)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Abbreviation:</td><td>&nbsp;<INPUT id=LadderAbbreviation name=Abbreviation style=" WIDTH: 100px" class=text value="<%=Server.HTMLEncode(labbr)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Challenge Rungs:</td><td>&nbsp;<INPUT id=LadderChallenge name=Challenge style=" WIDTH: 75px" class=text value="<%=Server.HTMLEncode(lchallenge)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Active:</td><td>&nbsp;<SELECT NAME=Active Class=text>
					<OPTION VALUE="0" <%If lactive = "0" Then Response.Write " SELECTED " %>>No</OPTION>
					<OPTION VALUE="1" <%If lactive = "1" Then Response.Write " SELECTED " %>>Yes</OPTION>
					</SELECT></td></tr>
<tr bgcolor=<%=bgcone%>><td align=right>Locked:</td><td>&nbsp;<SELECT NAME=Locked Class=text>
					<OPTION VALUE="0" <%If llocked = "0" Then Response.Write " SELECTED " %>>No</OPTION>
					<OPTION VALUE="1" <%If llocked = "1" Then Response.Write " SELECTED " %>>Yes</OPTION>
					</SELECT></td></tr>
	<tr bgcolor=<%=bgcone%>><td align=right>Game:</td><td>&nbsp;<SELECT NAME=GameID Class=text>
				<%
					strSQL = "SELECT GameID, GameName FROM tbl_Games WHERE GameID > 0 ORDER BY GameName ASC "
					oRS.Open strSQL, oConn
					If Not(oRS.EOF AND oRS.BOF) Then
						Do While Not(oRS.EOF)
							Response.Write "<OPTION VALUE=""" & oRS.Fields("GameID").Value & """ "
							If cStr(oRS.Fields("GameID").Value  & "") = cStr(intGameID & "") Then
								Response.Write " SELECTED "
							End If
							Response.Write ">" & Server.HTMLEncode(oRS.Fields("GameName").Value & "") & "</OPTION>" & vbCrLf
							oRs.MoveNext
						Loop					
					End If
					oRs.NextRecordset
					%>
					</SELECT></td></tr>
	<tr bgcolor=<%=bgctwo%>><td colspan=2 align=middle><INPUT id=submit1 name=submit1 type=submit value="Save Ladder Information"></td></tr>
<input type=hidden name=SaveMethod value="<%=sMethod%>">
<input type=hidden value="<%=Server.HTMLEncode(strLadderName)%>" name=OldName>
</TABLE>
</TD></TR>
</TABLE>
<input type=hidden name=SaveType value=playerladder>
</form>
<%
Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>