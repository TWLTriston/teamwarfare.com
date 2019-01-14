<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Files"

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
Dim strCategory
If Not(bSysAdmin or bAnyLadderAdmin) Then
	oConn.Close
	set oConn = Nothing
	Set oRS = Nothing
	Response.Clear  
	Response.Redirect "/errorpage.asp?error=3"
End If

Dim intFileID, intCategoryID, strDesc
intFileID = Request.QueryString("FileID")
If Not(isNumeric(intFileID)) Then
	oConn.CLose
	Set oCOnn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "default.asp?e=1"
End If
strSQL = "SELECT * FROM tbl_files WHERE FileID = '" & intFileID & "'"
oRs.Open strSQL, oConn
If Not(oRs.EOF AND oRS.BOF) Then
	intCategoryID = oRS.Fields("FileCategoryID").Value
	strDesc = oRs.Fields("Description").Value
EnD If
oRs.nextRecordset
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("TWL Essential Files") %>
	<table align=center border=0 cellpadding=0 cellspacing=0 BGCOLOR="#444444">
	<FORM METHOD="POST" ACTION="SaveItem.asp" id=form1 name=form1>
	<input type=hidden name=PlayerID value="<%=Session("PlayerID")%>">
	<input type=hidden name=SaveType value="EditFile">
	<input type=hidden name=FileID value="<%=intFileID%>">
	<TR>
		<TD>
			<table align=center border=0 cellpadding=4 cellspacing=1>
				<TR BGCOLOR="#000000">
					<TH COLSPAN=2>Edit a File</TH>
				</TR>
				<TR BGCOLOR="<%=bgctwo%>">
					<TD ALIGN=RIGHT><B>Category:</B></TD>
					<TD><SELECT NAME="CategoryID">
						<%
						strSQL = "SELECT CategoryName, FileCategoryID, Directory FROM tbl_file_category ORDER BY CategoryName"
						oRs.Open strSQL, oCOnn
						If Not(ors.EOF and ors.bof) Then
							Do While Not(ors.EOF)
								Response.Write "<OPTION VALUE=""" &  oRS.Fields("FileCategoryID").Value & """"
								If intCategoryID = oRs.Fields("FileCategoryID").Value Then
									Response.Write " selected "
								End If
								Response.Write ">" & oRS.Fields("CategoryName").Value & "</OPTION>"
								oRS.MoveNext
							Loop
						End If
						ors.Close
						%></SELECT>
						
				<TR BGCOLOR="<%=bgcone%>">
					<TD ALIGN=RIGHT VALIGN=TOP><B>Description:</B></TD>
					<TD><TEXTAREA COLS=20 ROWS=5 NAME="Description"><%=strDesc%></TEXTAREA>
				</TR>
				<tr BGCOLOR="<%=bgctwo%>">
					<td colspan=2 align=center><input type=submit name="Submit1" value=" Upload File "></td>
				</tr>		
			</table>
		</TD>
	</TR>
	</form>
	</TABLE>

<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn= Nothing
Set oRS = Nothing

Function CalculateSizeDisplay(byVal intSize)
	If Not(isNumeric(intSize)) Then
		CalculateSizeDisplay = "#ERROR#"
		Exit Function
	End If
	' Force conversion into a numeric format
	' using double just in case file is really big
	intSize = cDbl(intSize)
	If intSize < 1000 Then
		CalculateSizeDisplay = cStr(intSize) & " B"
	Elseif intSize < 1000000 Then
		CalculateSizeDisplay = FormatNumber(intSize/1000, 0,0,0,-1) & " KB"
	Else
		CalculateSizeDisplay = FormatNumber(intSize/1000000, 0, 0, 0, -1) & " MB"
	End If
End Function
%>

