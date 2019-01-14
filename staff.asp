<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Staff"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin, intColSpan
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
intColSpan = 3
If bSysAdmin Then
	intColSpan = 5
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<%
Dim blnFirstTime, strThisGroup
Call ContentStart("Teamwarfare Staff")
%>
<div align="center">
For any competition specific questions, or suggestions, please choose the game from the list below to contact the correct person.
<form name="frmStaff" id="frmStaff" method="get" action="">
<select name="StaffGroupID" id="StaffGroupID">
<option value="">&lt;-- choose group --&gt;</option>
<option value="0">Senior Staff &amp; Operations</option>
<%
strSQL = "SELECT Description, StaffGroupID FROM tbl_staff_group WHERE SeqNum = 100 ORDER BY sg.Description "
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRS.BOF) Then
	Do While Not(oRs.EOF)
		Response.Write "<option value=""" & oRs.Fields("StaffGroupID").Value & """"
		If Request.QueryString("StaffGroupID") = cStr(oRs.Fields("StaffGroupID").Value) Then
			Response.Write " selected=""selected"""
		End if
		Response.Write ">" & oRs.Fields("Description").Value & "</option>"
		oRs.MoveNext
	Loop
End If
oRs.NextRecordSet
%>
</select>
<input type="submit" value="Display Staff" />
</form></div>
<%
If Len(Request.QueryString("StaffGroupID")) > 0 AND IsNumeric(Request.QueryString("StaffGroupID")) Then
	strSQL = "SELECT s.*, 'Group' = sg.Description FROM tbl_staff s, tbl_staff_group sg WHERE sg.staffgroupid = s.staffgroupid AND sg.StaffGroupID = '" & Request.QueryString("StaffGroupID") & "' ORDER BY sg.SeqNum, sg.Description, s.SeqNum, s.displayname "
	oRS.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		%>
		<BR>
		<table class="cssbordered" width="100%">
		<TR>
			<TH COLSPAN=<%=intColSpan%> BGCOLOR="#000000"><%=oRs.Fields("Group").Value%></TH>
		</TR>
		
		<%
		Do While Not(oRS.EOF)
			call DisplayStaff(ors("staffid").value, ors("playerID").value, oRS("displayname").value, ors("email").value, ors("title").value, ors("description").value, oRS.Fields("seqnum").Value, bgc)
			oRs.MoveNext
		Loop
		%>
		</table>
		<br /><br />
		<%
	End If
	oRs.NextRecordSet

	bgc = bgcone
	blnFirstTime  = True
	strThisGroup = ""
	strSQL = "SELECT s.*, 'Group' = sg.Description FROM tbl_staff s, tbl_staff_group sg WHERE sg.staffgroupid = s.staffgroupid AND sg.SeqNum < 100 AND sg.SeqNum != 0 ORDER BY sg.SeqNum, sg.Description, s.SeqNum, s.displayname "
	oRS.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		Do While Not(oRS.EOF)
			If strThisGroup <> ors("group").value Then
				If Not(blnFirstTime) Then
					%>
					</TABLE>
					<%
				End If
				blnFirstTime = False
				strThisGroup = oRS("group").value
				%>
				<BR>
				<table border=0 cellspacing=0 cellpadding=0 class="cssbordered" width="97%" align=center>
				<TR>
					<TH COLSPAN=<%=intColSpan%> BGCOLOR="#000000"><%=strThisGroup%></TH>
				</TR>
				<%
			End If
			call DisplayStaff(ors("staffid").value, ors("playerID").value, oRS("displayname").value, ors("email").value, ors("title").value, ors("description").value, oRS.Fields("seqnum").Value, bgc)
			oRS.MoveNext
		Loop
	End If
	oRs.NextRecordSet
	%>
</TABLE>
<br /><br />
	
	<%
End If
%>

<% If bSysAdmin Then %>
<A href="editstaff.asp">New Staffer</A>
<% End If %>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn= Nothing
Set oRS = Nothing

Sub DisplayStaff(byVal staffID, byVal playerID, byVal Name, byVal email, byVal title, byVal description, byVal staffSeqNum, byRef bgc)
		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		%>
		<tr bgcolor=<%=bgc%>>
			<td width=200 align=left><a href="mailto:<%=email%>"><%=Name%></a></td>
			<td width=150 align=left><%=title%></td>
			<td><%=description%></td>
			<% If bSysAdmin Then %>
			<Td WIDTH=50><%=staffSeqNum%></TD>
			<td WIDTH=50><a href="editstaff.asp?staffid=<%=staffid%>&isEdit=True">Edit</a></td>
			<% End If %>
		</tr>         
		<%
End Sub
%>
