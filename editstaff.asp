<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Staff Member Maint"

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

Dim intPlayerID, intSTaffID, intStaffGroupID, strDisplayName, strTitle, strDescription
Dim strEmail, intSeqNum, strVerbage, strMethod

Dim bIsEdit
bIsEdit = cBool(Request.QueryString("IsEdit"))
If bIsEdit Then
	strVerbage = "Edit a Staff Member"
	intStaffID = Request.QueryString("StaffID")
	strMethod = "Edit"
	strSQL = "select * from tbl_staff where staffID='" & CheckString(intStaffID) & "'"
	oRs.Open strSQL, oConn
	if not (ors.eof and ors.BOF) then
		intPlayerID = oRS("PlayerID")
		intStaffGroupID = oRs("StaffGroupID")
		strDisplayName =oRS("DisplayName")
		strTitle = oRs("Title")
		strDescription = oRS("Description")
		strEmail = oRS("email")
		intSeqNum = oRS("seqNum")
		strMethod="Edit"
	end if
	ors.Close
Else
	strVerbage = "Add a Staffer"
	strMethod = "New"
End If
strPageTitle = strVerbage
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart(strVerbage)
%>
	<form name=frmAddTeam action=saveItem.asp method=post>
	<table align=center BACKGROUND="" BGCOLOR="#444444" CELLSPACING=0 CELLPADDING=0>
	<TR><TD>
	<table align=center CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
		<tr bgcolor=<%=bgcone%>><td align=right>Display Name:</td>
			<td width=300>&nbsp;<INPUT id=DisplayName name=DisplayName style=" WIDTH: 250px" class=text value="<%=Server.HTMLEncode(strDisplayName)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Title:</td>
			<td width=300>&nbsp;<INPUT id=Title name=Title style=" WIDTH: 250px" class=text value="<%=Server.HTMLEncode(strTitle)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Description:</td>
			<td width=300>&nbsp;<INPUT id=Description name=Description style=" WIDTH: 250px" class=text value="<%=Server.HTMLEncode(strDescription)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Email:</td>
			<td width=300>&nbsp;<INPUT id=Email name=Email style=" WIDTH: 250px" class=text value="<%=Server.HTMLEncode(strEmail)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>SeqNum:</td>
			<td width=300>&nbsp;<INPUT id=SeqNum name=SeqNum style=" WIDTH: 250px" class=text value="<%=Server.HTMLEncode(intSeqNum)%>"></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Group:</td>
			<td width=300>&nbsp;
				<SELECT NAME=StaffGroup>
					<% 
					strSQL = "Select StaffGroupID, Description From tbl_staff_group ORDER by Description"
					oRS.Open strSQL, oConn
					If Not(ors.eof and ors.bof) Then
						do while not(ors.eof)
							Response.write "<OPTION VALUE=""" & ors("staffgroupid") & """ "
							If ors("staffGroupID") = intStaffGroupID Then
								Response.write " SELECTED "
							End If
							Response.write ">" & oRS("Description") & "</OPTION>"
							
							ors.movenext
						loop
					End If
					oRS.NextRecordSet
					%>
					</SELECT>
					
					</td></tr>
<tr bgcolor=<%=bgctwo%>><td colspan=2 align=middle><INPUT type=submit value="Save Ladder Information"></td></tr>
<input type=hidden name=SaveMethod value="<%=strMethod%>">
<input type=hidden value="<%=intStaffID%>" name=StaffID>
</TABLE>
</TD></TR>
</TABLE>
<input type=hidden name=SaveType value=StaffMember>
</form>
<%
Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>