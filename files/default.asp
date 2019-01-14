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
Dim strCategory, oFile

If Not(Session("LoggedIn")) Then
	oConn.Close
	Set oConn = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=2"
End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("TWL Essential Files") %>
<div align="center">
	 Welcome to the TWL essential file downloads. Listed below are the various files that 
 we feel are neccessary to enjoy your gaming experience. Most likely, we missed a file, and would like you to tell us about it.
 Visit our forums, and let us know of any files we should add.
</div>
<% If bSysAdmin AND Session("uName") = "Triston" Then %>
<FORM NAME=NewCategory ACTION="saveitem.asp" METHOD="POST">
<INPUT TYPE="HIDDEN" NAME="SaveType" VALUE="AddCategory">
 <table width="60%" class="cssBordered" align="center">
	<TR BGCOLOR="#000000">
		<TH COLSPAN=2>New Category</TH>
	</TR>
	<TR BGCOLOR="<%=bgcone%>">
		<TD ALIGN=RIGHT>Category Name:</TD>
		<TD><INPUT TYPE=TEXT NAME="CategoryName"></TD>
	</TR>
	<TR BGCOLOR="<%=bgctwo%>">
		<TD ALIGN=RIGHT>Directory:</TD>
		<TD><INPUT TYPE=TEXT NAME="Directory"></TD>
	</TR>
	<TR BGCOLOR="<%=bgcone%>">
		<TD COLSPAN=2 ALIGN=CENTER><INPUT TYPE=SUBMIT VALUE="Add Category">
	</TR>
</TABLE>
</FORM>

<% End If %>
<% Call ContentEnd() %>
<% Call ContentStart("") %>
<%
Dim intGameID
intGameID = Request.QuerySTring("Game")
%>
<div align="center">
	<% If Request.QueryString("e") = "1" Then %>
	<b><font color="#ff0000">Invalid file id passed.</font></b>
	<% End If %>
	<form name="frmChooseGame" action="" method="get">
	Select a game: <select name="game" id="game">
		<%
		strSQL = "SELECT FileCategoryID, CategoryName FROM tbl_file_category WHERE IsActive = 1 ORDER BY CategoryName ASC" 
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND ORs.BOF) Then
			Do While Not(oRS.EOF)
				%>
				<option value="<%=oRs.Fields("FileCategoryID").Value%>" <% if CStr(intGameID & "") = CStr(oRs.Fields("FileCategoryID").Value & "") Then Response.Write " selected " End If %>><%=oRs.Fields("CategoryName").Value%></option>
				<%
				oRs.MoveNext
			Loop
		End If
		oRs.NextRecordSet
		%>
		</select>
		<br /><br />
		<input type="submit" value="View Downloads" />
	</form>
<%
If Len(intGameID) > 0 AND IsNumeric(intGameID) Then 
	%>
<table width="100%" class="cssBordered">
 <% If bSysAdmin Or bAnyLadderAdmin Then %>
 <TR BGCOLOR="#000000">
  <TH COLSPAN=7><A HREF="FileAdd.asp">Add New File</A></TH>
 </TR>
 <% End If %>
	
	<%
	Dim oFS
	Set oFS = Server.CreateObject ("Scripting.FileSystemObject")
	If bSysAdmin Then
		strSQL = "SELECT F.FileID, F.FileName, p.PlayerHandle, F.Description, FC.CategoryName, F.upload_time, FC.Directory "
		strSQL = strSQL & " FROM tbl_files f, tbl_file_category fc, tbl_players p"
		strSQL = strSQL & " WHERE p.PlayerID = f.PlayerID AND f.FileCategoryID = fc.FileCategoryID "
		strSQL = strSQL & " AND fc.FileCategoryID='" & intGameID & "'"
		strSQL = strSQL & " ORDER BY FC.CategoryName, F.FileName "
	Else
		strSQL = "SELECT F.FileID, F.FileName, F.Description, FC.CategoryName, FC.Directory "
		strSQL = strSQL & " FROM tbl_files f, tbl_file_category fc "
		strSQL = strSQL & " WHERE f.FileCategoryID = fc.FileCategoryID "
		strSQL = strSQL & " AND fc.FileCategoryID='" & intGameID & "'"
		strSQL = strSQL & " ORDER BY FC.CategoryName, F.FileName "
	End If
	oRS.Open strSQL, oConn
	If Not(oRS.EOF and oRS.BOF) Then
		Do While Not(oRS.EOF)
			If oFS.FileExists ("D:\TWLFILES\UPLOAD\" & oRS.Fields("Directory").Value & "\" & oRS.Fields("FileName").Value) Then
				Set oFile = oFS.GetFile("D:\TWLFILES\UPLOAD\" & oRS.Fields("Directory").Value & "\" & oRS.Fields("FileName").Value)
				If bgc = bgcone Then
					bgc = bgctwo
				Else
					bgc = bgcone
				End If
				If strCategory <> oRs.Fields("CategorYName").Value Then
					If Len(strCategory ) > 0 Then
						Response.write "<TR BGCOLOR=""#000000"">"
						If bSysAdmin Then
							Response.Write "<TD COLSPAN=6>"
						Else
							Response.Write "<TD COLSPAN=3>"
						End If
						Response.Write "&nbsp;</TD></TR>"
					End If
					strCategory = oRs.Fields("CategoryName").Value
					Response.Write "<TR BGCOLOR=""#000000"">"
					If bSysAdmin Then
						Response.Write "<TD COLSPAN=7>"
					Else
						Response.Write "<TD COLSPAN=3>"
					End If
					Response.Write "<B>" & oRS.Fields("CategoryName").Value & "</B></TD></TR>"
					Response.Write "<TR BGCOLOR=""#000000"">"
					Response.Write "<TH WIDTH=200>File Name</TH>"
					Response.Write "<TH WIDTH=100>File Size</TH>"
					Response.Write "<TH>Description</TH>"
					If bAnyLadderAdmin or bSysAdmin Then
						Response.Write "<TH>Edit</TH>"
					End If
					If bSysAdmin Then
						Response.Write "<TH>Up By</TH>"
						Response.Write "<TH>Up Time</TH>"
						Response.Write "<TH>Delete</TH>"
					End If				
					Response.Write "</TR>"
					
				End If
				%>
				<TR BGCOLOR="<%=bgc%>">
					<TD>&nbsp;<A href="dl/<%=oRs.Fields("Directory").Value%>/<%=oRS.Fields("FileName").Value%>"><%=Server.HTMLEncode(oRS.Fields("FileName").Value )%></A></TD>
					<TD ALIGN=RIGHT>&nbsp;<%=CalculateSizeDisplay(oFile.Size)%> </TD>
					<TD>&nbsp;<%=oRS.Fields("Description").Value %> </TD>
					<% If bAnyLadderAdmin or bSysAdmin Then %>
						<TD>&nbsp;<a href="FileEdit.asp?fileid=<%=oRS.Fields("FileID").Value %>">edit</A></TD>
					<% End If %>
					<% If bSysAdmin Then %>
						<TD><%=oRS.Fields("PlayerHandle").Value%></TD>
						<TD><%=FormatDateTime(oRS.Fields("Upload_time").Value, 2)%></TD>
						<TD>&nbsp;<a href="saveitem.asp?savetype=DeleteFile&fileid=<%=oRS.Fields("FileID").Value %>&category=<%=oRS.Fields("Directory").Value%>&filename=<%=Server.URLEncode(oRS.Fields("FileName").Value)%>">delete</A></TD>
					<% End If %>
				</TR>
				<%
			
			End If
			oRs.MoveNext
		Loop
	Else
		%>
		<tr bgcolor="#000000"><td colspan="6">No files are available for download at this time.</td></tr>
		<%
	End If
	oRS.Close 
	%>
	</TABLE>
	<%
End If
%> 
</div>
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
	Elseif intSize < 10000000 Then
		CalculateSizeDisplay = FormatNumber(intSize/1000, 0,0,0,-1) & " KB"
	Else
		CalculateSizeDisplay = FormatNumber(intSize/1000000, 0, 0, 0, -1) & " MB"
	End If
End Function
%>

