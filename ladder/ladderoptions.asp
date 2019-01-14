<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Ladder Options"

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

Dim intLadderID, strLadderName
intLadderID = Request.QueryString("LadderID")

if not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If
Dim intCols, intCurrent
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<%
Call ContentStart("Jump to Ladder")
%>
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444" ALIGN=CENTER WIDTH="97%">
<TR>
	<TD>
	<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH="100%">
	<TR BGCOLOR="#000000">
		<TH COLSPAN=3>Choose a ladder:</TH>
	</TR>
	<%
	intCols = 3
	intCurrent = 0
	strSQL = "SELECT LadderName, LadderID FROM tbl_ladders WHERE LadderShown = 1 ORDER BY LadderName ASC"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		Do While Not(oRS.EOF)
			If intCurrent MOD intCols = 0 Then
				If intCurrent > 0 Then
					Response.Write "</TR>"
				End If					
				If bgc = bgcone Then
					bgc = bgctwo 
				Else
					bgc = bgcone
				End If
				Response.Write "<TR BGCOLOR=""" & bgc & """>"
			End If
			%>
				<TD ALIGN=CENTER><B><A href="ladderoptions.asp?ladderid=<%=oRS.Fields("LadderID").Value%>"><%=Server.HTMLEncode(oRS.Fields("LadderName").Value)%></A></B></TD>
			<%
			oRS.MoveNext
			intCurrent = intCurrent + 1
		Loop
		While intCurrent Mod intCols <> 0
			Response.Write "<TD></TD>"
			intCurrent = intCurrent + 1
		Wend
		Response.Write "</TR>"
	End If
	oRS.NextRecordset 
	%>
	</TABLE>
	</TD>
</TR>
</TABLE>
<%
Call ContentEnd()
If Len(intLadderID) > 0 Then
	strSQL = "SELECT LadderName FROM tbl_ladders WHERE ladderID = '" & intLadderID & "'"
	oRS.Open strSQL, oConn
	If Not(oRS.BOF AND oRS.EOF) Then
		strLadderName = oRS.Fields("LadderName").Value 
	End If
	oRS.NextRecordset 
	Call ContentStart("Configure " & strLadderName & " Ladder Options")
	%>
	<A HREF="addoption.asp?ladderid=<%=intLadderID%>">Add a match option</A><BR>
	<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 BGCOLOR="#444444">
	<TR><TD>
		<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=2 WIDTH=100%>
		<TR BGCOLOR="#000000">
			<TH>Option Name</TH>
			<TH>Map Number</TH>
			<TH>Selected By</TH>
			<TH>Side Choice</TH>
			<TH>Edit</TH>
			<TH>Delete</TH>
		</TR>
		<%
		strSQL = "SELECT lo.* "
		strSQL = strSQL & " FROM tbl_ladder_options lo "
		strSQL = strSQL & " WHERE lo.LadderID = '" & intLadderID & "' ORDER BY OptionName, MapNumber ASC"
		oRS.Open strSQL, oConn
		If Not(oRS.EOF AND oRS.BOF) Then
			Do While Not(oRS.EOF)
				If bgc = bgcone Then
					bgc = bgctwo
				Else
					bgc = bgcone
				End If
				Response.Write "<TR BGCOLOR=""" & bgc & """>"
				Response.Write "<TD>" & Server.HTMLEncode("" & oRS.Fields("OptionName").Value) & "</TD>"
				Response.Write "<TD>" & Server.HTMLEncode("" & oRS.Fields("MapNumber").Value) & "</TD>"
				Select Case uCase(oRS.Fields("SelectedBy").Value)
					Case "A"
						Response.Write "<TD>Attacker</TD>"
					Case "D"
						Response.Write "<TD>Defender</TD>"
					Case "R"
						Response.Write "<TD>Random</TD>"
					Case Else
						Response.Write "<TD>Unknown</TD>"
				End Select
				
				Select Case uCase(oRS.Fields("SideChoice").Value)
					Case "Y"
						Response.Write "<TD>Yes</TD>"
					Case "N"
						Response.Write "<TD>No</TD>"
					Case Else
						Response.Write "<TD>Unknown</TD>"
				End Select
				Response.Write "<TD><A HREF=""addoption.asp?ladderid=" & intLadderID & "&optionid=" & oRS.Fields("OptionID").Value & "&isedit=true"">Edit</A></TD>"
				If uCase(oRS.Fields("Active").Value) = "Y" Then
					Response.Write "<TD><A HREF=""option_saveitem.asp?optionid=" & oRS.Fields("OptionID").Value & "&ladderid=" & intLadderID & "&savetype=deleteoption"">Delete</A></TD>"
				Else
					Response.Write "<TD><A HREF=""option_saveitem.asp?optionid=" & oRS.Fields("OptionID").Value & "&ladderid=" & intLadderID & "&savetype=restoreoption"">Restore</A></TD>"
				End If
				Response.Write "</TR>"
				oRS.MoveNext
			Loop
		Else
			%>
			<TR BGCOLOR="#000000">
				<TD COLSPAN=6><I>No options configured</I></TD>
			</TR>
			<%
		End If
		%>
		</TABLE>
	</TD></TR>
	</TABLE>
	<%
	oRS.NextRecordset 
	Call ContentEnd()
End If
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>