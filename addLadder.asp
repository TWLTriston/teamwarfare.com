<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add a Ladder"

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

Dim strVerbage, bIsEdit, strLadderName, strMethod, lrules
Dim lAdmin, lGame, lActive, lid, labbr, llocked, lchallenge
Dim maps, mapconfiguration, timezone, timeoptions, intGameID
Dim minPlayer, Scoring, chrRequireDistinctMaps
Dim intChallengeDays, intMatchDays, intMaxRoster, restrank, intIdentifierID
timezone = "EST"
timeoptions = "8:30|9:00|9:30|10:00|10:30"
maps = "3"
mapConfiguration = "RRD"
chrRequireDistinctMaps = "Y"
intMaxRoster = 0
bIsEdit = cBool(Request.QueryString("IsEdit"))
If bIsEdit Then
	strVerbage = "Edit a ladder"
	strLadderName = Request.QueryString("ladder")
	strMethod = "Edit"
	strSQL = "select * from tbl_ladders where laddername='" & CheckString(strLaddername) & "'"
	oRs.Open strSQL, oConn
	if not (ors.eof and ors.BOF) then
		restrank = ors.Fields("RestRank").Value
		lactive = ors.Fields("LadderActive").Value 
		lid = ors.Fields("LadderID").Value 
		labbr = ors.Fields("LadderAbbreviation").Value 
		llocked = ors.Fields("LadderLocked").Value 
		lchallenge = ors.Fields("LadderChallenge").Value 
		lrules = ors.Fields("LadderRules").Value
		minPlayer = oRS.Fields("MinPlayer").Value 
		maps = oRS.Fields("maps").Value 
		mapconfiguration = oRS.Fields("mapconfiguration").Value 
		timezone = oRS.Fields("timezone").Value 
		timeoptions = oRS.Fields("timeoptions").Value 
		Scoring = ors.Fields("Scoring").Value
		intGameID = oRS.FIelds("GameID").Value 
		chrRequireDistinctMaps = oRs.Fields("RequireDistinctMaps").Value 
		intChallengeDays = oRS.Fields("ChallengeDays").Value
		intMatchDays = oRS.Fields("MatchDays").Value
		intMaxRoster = oRs.Fields("RosterLimit").Value
		intIdentifierID = oRs.Fields("IdentifierID").Value
		strMethod="Edit"
	end if
	ors.Close
Else
	strVerbage = "Add a ladder"
	strMethod = "New"
	intChallengeDays = 254
	intMatchDays = 254
	restrank = 20
End If
strPageTitle = strVerbage

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart(strVerbage)
%>
	<form name="frmAddLadder" id="frmAddLadder" action=saveItem.asp method=post>
	<table align=center BACKGROUND="" BGCOLOR="#444444" CELLSPACING=0 CELLPADDING=0>
	<TR><TD>
	<table align=center CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
		<TR BGCOLOR="#000000"><TH COLSPAN=2><%=strVerbage%></TH></TR>
		<tr bgcolor=<%=bgcone%>><td align=right>Name:</td><td width=300>&nbsp;<INPUT id=LadderName name=LadderName style=" WIDTH: 250px" maxlength="100" class=text value="<%=Server.HTMLEncode(strLadderName)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Abbreviation:</td><td>&nbsp;<INPUT id=LadderAbbreviation name=LadderAbbreviation maxlength="100" style=" WIDTH: 100px" class=text value="<%=Server.HTMLEncode(labbr)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Rule Set Name:</td><td>&nbsp;<INPUT id=LadderRules name=LadderRules maxlength="100" style=" WIDTH: 250px" class=text value="<%=Server.HTMLEncode(lrules & "")%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Min Rung to Rest: (default 20)</td><td>&nbsp;<INPUT id=restrank name=restrank style=" WIDTH: 50px" class=text value="<%=Server.HTMLEncode(restrank)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Challenge Rungs:</td><td>&nbsp;<INPUT id=LadderChallenge name=LadderChallenge style=" WIDTH: 50px" class=text value="<%=Server.HTMLEncode(lchallenge)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Minimum Roster to Challenge:</td><td>&nbsp;<INPUT id=MinPlayer name=MinPlayer style=" WIDTH: 50px" class=text value="<%=Server.HTMLEncode(minPlayer & "")%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Maximum Roster:</td><td>&nbsp;<INPUT id=MaxRoster name=MaxRoster style=" WIDTH: 50px" class=text value="<%=Server.HTMLEncode(intMaxRoster & "")%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Time Zone:</td><td>&nbsp;<INPUT id=TimeZone name=TimeZone style=" WIDTH: 50px" class=text value="<%=Server.HTMLEncode(TimeZone & "")%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Time Options (Pipe Delimit "|"):</td><td>&nbsp;<INPUT id=timeoptions name=timeoptions style=" WIDTH: 200px" class=text value="<%=Server.HTMLEncode(timeoptions & "")%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Maps (Number):</td><td>&nbsp;<INPUT id=Maps name=Maps style=" WIDTH: 25px" class=text value="<%=Server.HTMLEncode(Maps & "")%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Map Configuration:</td><td>&nbsp;<INPUT id=mapconfiguration name=mapconfiguration style=" WIDTH: 50px" class=text value="<%=Server.HTMLEncode(MapConfiguration & "")%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Anti-Smurf Criteria:</td><td>&nbsp;
			<select name="selIdentifierID" id="selIdentifierID">
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
		
		<tr bgcolor=<%=bgcone%>><Td colspan=2>This defines how maps are selected. Possible value: RRD meaning Map 1 is R, Map 2 is R, and Map 3 is D. Options are: <UL>
		<LI>C : At time of challenge, attacker picks map
		<LI>D : At time of acceptance, defender picks map
		<LI>A : After acceptance, attacker picks map
		<LI>R : After acceptace, system randomly chooses map</UL></TD></TR>
		<tr bgcolor=<%=bgcone%>><td align=right>Scoring:</td><td>&nbsp;<SELECT id=Scoring name=Scoring>
					<OPTION>NOTHING SELECTED</OPTION>
					<OPTION VALUE="B" <% If Scoring = "B" THen Response.Write " SELECTED " End If %>>Each map must have a winner.</OPTION>
					<OPTION VALUE="M" <% If Scoring = "M" THen Response.Write " SELECTED " End If %>>Count points for each map, winner defined by total scores.</OPTION>
					</SELECT></TD></TR>					

		<tr bgcolor=<%=bgcone%>><td align=right>Active:</td><td>&nbsp;<SELECT NAME=LadderActive Class=text>
					<OPTION VALUE="0" <%If lactive = "0" Then Response.Write " SELECTED " %>>No</OPTION>
					<OPTION VALUE="1" <%If lactive = "1" Then Response.Write " SELECTED " %>>Yes</OPTION>
					</SELECT></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Locked:</td><td>&nbsp;<SELECT NAME=LadderLocked Class=text>
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
		<tr bgcolor=<%=bgcone%>><td align=right>Require Distinct Maps:</td><td>&nbsp;<SELECT NAME=RequireDistinctMaps Class=text>
					<OPTION VALUE="N" <%If chrRequireDistinctMaps = "N" Then Response.Write " SELECTED " %>>No</OPTION>
					<OPTION VALUE="Y" <%If chrRequireDistinctMaps = "Y" Then Response.Write " SELECTED " %>>Yes</OPTION>
					</SELECT></td></tr>
		<tr bgcolor=<%=bgcone%>><Td colspan=2>This defines if all maps must be distinct, for most cases, this should be Yes</TD></TR>
		<tr bgcolor=<%=bgctwo%>><td align=right valign="top">Can challenge days:</td><td>
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbSunday%>" name="chkChallengeSunday" id="chkChallengeSunday" <% if (intChallengeDays and 2 ^ vbSunday) Then %> checked="checked" <% end if %>> Sunday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbMonday%>" name="chkChallengeMonday" id="chkChallengeMonday" <% if (intChallengeDays and 2 ^ vbMonday) Then %> checked="checked" <% end if %>> Monday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbTuesday%>" name="chkChallengeTuesday" id="chkChallengeTuesday" <% if (intChallengeDays and 2 ^ vbTuesday) Then %> checked="checked" <% end if %>> Tuesday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbWednesday%>" name="chkChallengeWednesday" id="chkChallengeWednesday" <% if (intChallengeDays and 2 ^ vbWednesday) Then %> checked="checked" <% end if %>> Wednesday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbThursday%>" name="chkChallengeThursday" id="chkChallengeThursday" <% if (intChallengeDays and 2 ^ vbThursday) Then %> checked="checked" <% end if %>> Thursday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbFriday%>" name="chkChallengeFriday" id="chkChallengeFriday" <% if (intChallengeDays and 2 ^ vbFriday) Then %> checked="checked" <% end if %>> Friday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbSaturday%>" name="chkChallengeSaturday" id="chkChallengeSaturday" <% if (intChallengeDays and 2 ^ vbSaturday) Then %> checked="checked" <% end if %>> Saturday<br />
		&nbsp;&nbsp;<input type="checkbox" class="borderless" onclick="if(this.checked) {CheckAllChallenge(this.form); }"> All
		</td></tr>
		
		<tr bgcolor=<%=bgcone%>><td align=right valign="top">Can play match on days:</td><td>
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbSunday%>" name="chkMatchSunday" id="chkMatchSunday" <% if (intMatchDays and 2 ^ vbSunday) Then %> checked="checked" <% end if %>> Sunday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbMonday%>" name="chkMatchMonday" id="chkMatchMonday" <% if (intMatchDays and 2 ^ vbMonday) Then %> checked="checked" <% end if %>> Monday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbTuesday%>" name="chkMatchTuesday" id="chkMatchTuesday" <% if (intMatchDays and 2 ^ vbTuesday) Then %> checked="checked" <% end if %>> Tuesday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbWednesday%>" name="chkMatchWednesday" id="chkMatchWednesday" <% if (intMatchDays and 2 ^ vbWednesday) Then %> checked="checked" <% end if %>> Wednesday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbThursday%>" name="chkMatchThursday" id="chkMatchThursday" <% if (intMatchDays and 2 ^ vbThursday) Then %> checked="checked" <% end if %>> Thursday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbFriday%>" name="chkMatchFriday" id="chkMatchFriday" <% if (intMatchDays and 2 ^ vbFriday) Then %> checked="checked" <% end if %>> Friday<br />
		&nbsp;<input type="checkbox" class="borderless" value="<%=2 ^ vbSaturday%>" name="chkMatchSaturday" id="chkMatchSaturday" <% if (intMatchDays and 2 ^ vbSaturday) Then %> checked="checked" <% end if %>> Saturday<br />
		&nbsp;&nbsp;<input type="checkbox" class="borderless" onclick="if(this.checked) {CheckAllMatch(this.form); }"> All
		</td></tr>
					
<tr bgcolor=<%=bgctwo%>><td colspan=2 align=middle><INPUT id=submit1 name=submit1 type=submit value="Save Ladder Information"></td></tr>
<input type=hidden name=SaveMethod value="<%=strMethod%>">
<input type=hidden value="<%=Server.HTMLEncode(strLadderName)%>" name=OldName>
</TABLE>
</TD></TR>
</TABLE>
<input type=hidden name=SaveType value=ladder>
</form>
<script language="javascript" type="text/javascript">
<!--
function CheckAllChallenge(objForm) {
	objForm.chkChallengeSunday.checked = true;
	objForm.chkChallengeMonday.checked = true;
	objForm.chkChallengeTuesday.checked = true;
	objForm.chkChallengeWednesday.checked = true;
	objForm.chkChallengeThursday.checked = true;
	objForm.chkChallengeFriday.checked = true;
	objForm.chkChallengeSaturday.checked = true;
}
function CheckAllMatch(objForm) {
	objForm.chkMatchSunday.checked = true;
	objForm.chkMatchMonday.checked = true;
	objForm.chkMatchTuesday.checked = true;
	objForm.chkMatchWednesday.checked = true;
	objForm.chkMatchThursday.checked = true;
	objForm.chkMatchFriday.checked = true;
	objForm.chkMatchSaturday.checked = true;
}
//-->
</script>

<%
Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>