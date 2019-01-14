<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Accdept Match"

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

Dim strOpponent, strLadder, strPlayerName
strOpponent = Request.QueryString("enemy")
strLadder = Request.QueryString ("ladder")
strPlayerName = Request.QueryString("player")

If Not(bSysAdmin Or IsPlayerLadderAdmin(strLadder) Or Session("uName") = strPlayerName) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "/errorpage.asp?error=3"
end if

Dim intMaps
strSQL = "SELECT COUNT(*) FROM lnk_pl_m lnk, tbl_playerLadders l "
		strSQL = strSQL & " WHERE l.Playerladderid = lnk.Playerladderid "
		strSQL = strSQL & " AND l.PlayerLadderName='" & CheckString(strLadder) & "'"
oRs.Open strSQL, oConn
If Not(oRS.EOF AND oRS.BOF) THen
	intMaps = oRs.Fields(0).Value
End If
oRs.NextRecordSet

If intMaps = 0 Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "errorpage.asp?error=31"
End If

Dim strTimeZone, i, theday, dotw, mname, dy, j
Dim dayStr, dayArray

strTimeZone = "EST"
dayStr = "8:30|9:00|9:30|10:00|10:30|11:00|11:30"
dayArray = Split(dayStr, "|")
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("Accept Match with " & strOpponent & " on the " & strLadder & " Ladder") %>
<SCRIPT LANGUAGE="JavaScript">
<!-- 
	function acceptSubmit(objForm) {
		var err = 'n';
		var errStr = 'Error: \n';
		
		date1 = objForm.Day1.value.split(" ");
		date2 = objForm.Day2.value.split(" ");
		if (date1[2] == date2[2]) {
			err = 'y';
			errStr = errStr + "The dates you offered are on the same day. You must choose different days.\n";
		}
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
<form name=frmAccept action=../saveitem.asp method=post>
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
	For j = lBound(dayArray) to uBound(dayArray)
		Response.Write "<option VALUE=""" & dotw & ", " & mname & " " & dy & " " & dayArray(j) & " PM " & strTimeZone & """>" & dotw & ", " & mname & " " & dy & " " & dayArray(j) & " PM " & strTimeZone & "</OPTION>" & vbCrLf
	Next
next
%>
</SELECT>
</td></tr>
<tr height=35 bgcolor=<%=bgctwo%>><td align=center>
Match Date Two: <SELECT name=Day2 class=bright>
<%
for i = 2 to 9
	theday = now + i
	dotw = weekdayname(weekday(theday))
	mname = monthname(month(theday))
	dy = day(theday)
	For j = lBound(dayArray) to uBound(dayArray)
		Response.Write "<option VALUE=""" & dotw & ", " & mname & " " & dy & " " & dayArray(j) & " PM " & strTimeZone & """>" & dotw & ", " & mname & " " & dy & " " & dayArray(j) & " PM " & strTimeZone & "</OPTION>" & vbCrLf
	Next
next
%></SELECT>
</td></tr><tr  height=30 bgcolor=<%=bgcone%>><td align=center>
<input type=hidden name=matchid value=<%=Request.QueryString("matchid")%>>
<input type=hidden name=SaveType value=PlayerAcceptMatch>
<input type=hidden name=PlayerName value="<%=Server.HTMLEncode(strPlayerName)%>">
<input type=hidden name=ladder value="<%=Server.HTMLEncode(strLadder)%>">
<INPUT TYPE=BUTTON ONCLICK="javascript:acceptSubmit(this.form)" VALUE="Accept Match">
</tr></td>
</form>
</td>
</tr>
</table></td>
</tr>
</table>
<%
Call ContentEnd()
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
%>

