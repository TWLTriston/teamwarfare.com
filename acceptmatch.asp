<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Accept Match"

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

Dim strEnemyName, strLadderName, strTeamName
strTeamName = Request.QueryString("team")
strEnemyName = Request.QueryString("enemy")
strLadderName = Request.QueryString("ladder")

if not(bSysAdmin OR IsTeamFounder(strTeamName) OR IsTeamCaptain(strTeamName, strLadderName) OR IsLadderAdmin(strLadderName)) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If

Dim strTimeZone

Dim theDay, dotw, mName, dy, i, j
Dim dayStr, dayArray, strMapConfiguration, intMaps, intMatchDays

strSQL = "SELECT Maps, MapConfiguration, TimeZone, TimeOptions, MatchDays FROM tbl_ladders WHERE LadderName = '" & CheckString(strLadderName) & "'"
oRs.Open strSQL, oConn
If Not(oRS.EOF and oRS.BOF) Then
	dayStr = oRS.Fields("TimeOptions").Value 
	strTimeZone = oRS.Fields("TimeZone").Value 
	intMaps = oRS.Fields("Maps").Value 
	strMapConfiguration = oRS.Fields("MapConfiguration").Value 
	intMatchDays = oRS.Fields("MatchDays").Value
End If
oRS.Close

dayArray = Split(dayStr, "|")
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Accept Match with " & strEnemyName & " on the " & strLadderName & " Ladder")
%>
<SCRIPT LANGUAGE="JavaScript">
<!-- 
	function acceptSubmit(objForm) {
		var err = 'n';
		var errStr = 'Error: \n';
		date1 = objForm.Day1.value.split(" ");
		date2 = objForm.Day2.value.split(" ");
		<% if Not((intMatchDays = 2^vbSunday) _
					OR (intMatchDays = 2^vbMonday) _
					OR (intMatchDays = 2^vbTuesday) _
					OR (intMatchDays = 2^vbWednesday) _
					OR (intMatchDays = 2^vbThursday) _
					OR (intMatchDays = 2^vbFriday) _
					OR (intMatchDays = 2^vbSaturday)) Then %>
		if (date1[2] == date2[2]) {
			err = 'y';
			errStr = errStr + "The dates you offered are on the same day. You must choose different days.\n";
		}
		<% End If %>
		if (date1[3] == date2[3]) {
			err = 'y';
			errStr = errStr + "The time of day you offered is the same. You must choose different times of day.\n";
		}
		if (err == 'n') {
			//alert('Submitting...');
			objForm.submit();
		} else {
			errStr = errStr + "Please fix the error, and try again.\n";
			alert(errStr);
		}
		
	}
//-->
</SCRIPT>
<form name=frmAccept action=saveitem.asp method=post>
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER WIDTH="75%">
	<TR><TD>
    <table width="100%" border="0" cellpadding=4 CELLSPACING=1>
      <tr bgcolor=<%=bgctwo%> height=35>
     <td align=center>
     <p class=small>
     From the following drop downs, please propose two dates and times. These dates must be different, as well as the times.
     <BR><BR>
     For example:
     <center>
     <blockquote><FONT COLOR="red">
     Monday, January 1, 8:30 PM EST<BR>
     Tuesday, January 2, 8:30 PM EST<BR>
     Not acceptable, the times must be different.</FONT>
     </blockquote>
     <blockquote><FONT COLOR="red">
     Monday, January 1, 8:30 PM EST<BR>
     Monday, January 1, 9:00 PM EST<BR>
     Not acceptable, the dates must be different.</FONT>
     </blockquote>
     <blockquote>
     <B>Monday, January 1, 8:30 PM EST<BR>
     Tuesday, January 2, 9:00 PM EST<BR>
     Acceptable, times and dates are different.</B>
     </blockquote>
     </center>
     </TD>
     </TR>
      <tr bgcolor=<%=bgcone%> height=35>
     <td align=center>
Match Date One: <select name=Day1 class="bright">
<%
for i = 2 to 9
	theday = now + i
	dotw = weekdayname(weekday(theday))
	mname = monthname(month(theday))
	dy = day(theday)
	IF intMatchDays AND 2^weekday(theday) Then
		For j = lBound(dayArray) to uBound(dayArray)
			If month(theday) <> 9 OR Not(day(theday) = 11) Then
				Response.Write "<option VALUE=""" & dotw & ", " & mname & " " & dy & " " & dayArray(j) & " PM " & strTimeZone & """>" & dotw & ", " & mname & " " & dy & " " & dayArray(j) & " PM " & strTimeZone & "</OPTION>" & vbCrLf
			End If
		Next
	End If
next
%>
</select>
</td></tr>
<tr height=35 bgcolor=<%=bgctwo%>><td align=center>
Match Date Two: <select name=Day2 class=bright>
<%
for i = 2 to 9
	theday = now + i
	dotw = weekdayname(weekday(theday))
	mname = monthname(month(theday))
	dy = day(theday)
	IF intMatchDays AND 2^weekday(theday) Then
		For j = lBound(dayArray) to uBound(dayArray)
			If month(theday) <> 9 OR Not(day(theday) = 11) Then
				Response.Write "<option VALUE=""" & dotw & ", " & mname & " " & dy & " " & dayArray(j) & " PM " & strTimeZone & """>" & dotw & ", " & mname & " " & dy & " " & dayArray(j) & " PM " & strTimeZone & "</OPTION>" & vbCrLf
			End If
		Next
	ENd If
next
%></select>
</td></tr>
<%
For i = 1 To Len(strMapConfiguration)
	If Mid(strMapConfiguration, i, 1) = "D" Then
		%>
		<TR>
			<TD ALIGN=CENTER BGCOLOR=<%=bgcone%>>Choose Map <%=i%>: <SELECT Name=Map<%=i%> CLASS=bright>
			<%
			strSQL = "EXEC GetMapList '" & Request("matchid") & "', " & i
'			Response.Write strSQL
			ors.Open strsql, oconn
			if not (ors.EOF and ors.BOF) then
				do while not ors.EOF
					Response.Write "<option VALUE=""" & ors.Fields("MapName").Value & """>" & ors.Fields("MapName").Value & "</OPTION>" & vbCrLf
					ors.MoveNext 
				loop
			End If
			oRS.Close 
			%>
			</TD>
		</TR>
		<%
	End If
Next
%>
<tr height=35 bgcolor=<%=bgctwo%>><td align=center>Approved for Shoutcasting: <input type=checkbox name=scApproved value=true checked></td></tr>
<tr height=30 bgcolor=<%=bgcone%>><td align=center>
<input type=hidden name=matchid value=<%=Request.QueryString("matchid")%>>
<input type=hidden name=SaveType value=AcceptMatch>
<input type=hidden name=team value="<%=Server.HTMLEncode(strTeamName)%>">
<input type=hidden name=ladder value="<%=Server.HTMLEncode(strLadderName)%>">
<input type=hidden name="MC" value="<%=Server.HTMLEncode(strMapConfiguration)%>">
<input type=hidden name="MD" value="<%=Server.HTMLEncode(intMatchDays)%>">
<%
if i > 0 then
	response.write "<input type=BUTTON name=Button value=""Accept Match"" class=bright ONCLICK=""javaScript:acceptSubmit(this.form);"">"
Else
	Response.Clear
	Response.Redirect "/errorpage.asp?error=20"
End if
%>
</tr></td>
</form>
</td>
</tr>
</table>
</td>
</tr>
</table>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>