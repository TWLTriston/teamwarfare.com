<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add a Scrim Ladder"

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

Dim strVerbage, bIsEdit, strLadderName, strMethod
Dim strRules, intActive, intLocked, strAbbr, intMinPlayer, intForumID
Dim intMaxRatingDiff, intGameID, intLadderID, intIdentifierID

bIsEdit = cBool(Request.QueryString("IsEdit"))
If bIsEdit Then
	strVerbage = "Edit an Elo ladder"
	strLadderName = Request.QueryString("ladder")
	strMethod = "Edit"
	strSQL = "select * from tbl_elo_ladders where EloLadderName='" & CheckString(strLadderName) & "'"
	oRs.Open strSQL, oConn
	if not (ors.eof and ors.BOF) then
		intLadderID = oRs.Fields("EloLadderID").Value
		intGameID = oRs.Fields("EloGameID").Value
		strRules = oRs.Fields("EloRulesName").Value
		intActive = oRs.Fields("EloActive").Value
		intLocked = oRs.Fields("EloLocked").Value
		strAbbr = oRs.Fields("EloAbbreviation").Value
		intMinPlayer = oRs.Fields("EloMinPlayer").Value
		intForumID = oRs.Fields("EloForumID").Value
		intMaxRatingDiff = oRs.Fields("EloMaxRatingDiff").Value
		intIdentifierID =  oRs.Fields("EloIdentifierID").Value
		strMethod="Edit"
	end if
	ors.Close
Else
	strVerbage = "Add an Scrim ladder"
	strMethod = "New"
End If
strPageTitle = "TWL: " &  strVerbage

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart(strVerbage)
%>
<% If Request.QueryString("e") = "1" Then %>
<b>Duplicate ladder name found. Choose another</b>
<% End If %>
	<form name="frmAddLadder" id="frmAddLadder" action=savescrim.asp method=post>
	<table align=center BACKGROUND="" BGCOLOR="#444444" CELLSPACING=0 CELLPADDING=0>
	<TR><TD>
	<table align=center CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
		<TR BGCOLOR="#000000"><TH COLSPAN=2><%=strVerbage%></TH></TR>
		<tr bgcolor=<%=bgcone%>><td align=right>Name:</td><td width=300>&nbsp;<INPUT id=LadderName name=LadderName style=" WIDTH: 250px" maxlength="100" class=text value="<%=Server.HTMLEncode(strLadderName)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Abbreviation:</td><td>&nbsp;<INPUT id=Abbreviation name=Abbreviation maxlength="100" style=" WIDTH: 100px" class=text value="<%=Server.HTMLEncode(strAbbr)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Rule Set Name:</td><td>&nbsp;<INPUT id=Rules name=Rules maxlength="100" style=" WIDTH: 250px" class=text value="<%=Server.HTMLEncode(strRules & "")%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Minimum Roster to Challenge:</td><td>&nbsp;<INPUT id=MinPlayer name=MinPlayer style=" WIDTH: 50px" class=text value="<%=Server.HTMLEncode(intMinPlayer & "")%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Max Rating Diff to Challenge:</td><td>&nbsp;<INPUT id=MaxRatingDiff name=MaxRatingDiff style=" WIDTH: 50px" class=text value="<%=Server.HTMLEncode(intMaxRatingDiff & "")%>"></td></tr>
		 
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
		<tr bgcolor=<%=bgcone%>><td align=right>Anti-Smurf Criteria:</td><td>&nbsp;<select name="selIdentifierID" id="selIdentifierID">
				<option value="">No identifier tracked</option>
				<%
				strSQL = "SELECT IdentifierID, IdentifierName FROM tbl_identifiers ORDER BY IdentifierName ASC "
				oRs.Open strSQL, oConn
				If Not(oRS.EOF AND oRs.BOF) THen
					Do While Not(oRs.EOF)
						Response.Write "<option value=""" & oRs.Fields("IdentifierID") & """"
						If CStr(oRs.FieldS("IdentifierID").Value & "") = CStr(intIdentifierID & "") Then
							Response.Write "Selected=""selected"""
						End if
						Response.Write ">" & Server.HTMLEncode(oRs.Fields("IdentifierName").Value) & "</option>" & vBCrLf
						oRs.MoveNext
					Loop
				End If
				oRs.NextRecordSet
				%>
				</select>			
		</td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Active:</td><td>&nbsp;<SELECT NAME=Active Class=text>
					<OPTION VALUE="0" <%If intActive = "0" Then Response.Write " SELECTED " %>>No</OPTION>
					<OPTION VALUE="1" <%If intActive= "1" Then Response.Write " SELECTED " %>>Yes</OPTION>
					</SELECT></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Locked:</td><td>&nbsp;<SELECT NAME=Locked Class=text>
					<OPTION VALUE="0" <%If intLocked = "0" Then Response.Write " SELECTED " %>>No</OPTION>
					<OPTION VALUE="1" <%If intLocked = "1" Then Response.Write " SELECTED " %>>Yes</OPTION>
					</SELECT></td></tr>
<tr bgcolor=<%=bgctwo%>><td colspan=2 align=middle><INPUT id=submit1 name=submit1 type=submit value="Save Ladder Information"></td></tr>
<input type=hidden name=SaveMethod value="<%=strMethod%>">
<input type=hidden value="<%=Server.HTMLEncode(strLadderName)%>" name="OldName">
<input type=hidden value="<%=Server.HTMLEncode(intLadderID)%>" name="EloLadderID">
</TABLE>
</TD></TR>
</TABLE>
<input type=hidden name=SaveType value="EloLadder" />
</form>
<%
Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>