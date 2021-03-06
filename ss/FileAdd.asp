<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: ScreenShots"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc

bgcone = Application("bgcone")
bgctwo = Application("bgctwo")
	
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
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("TWL Essential Files") %>
	<table align=center border=0 cellpadding=0 cellspacing=0 BGCOLOR="#444444">
	<FORM ENCTYPE="MULTIPART/FORM-DATA" METHOD="POST" ACTION="FileSave.asp" id=form1 name=form1>
	<input type=hidden name=PlayerID value="<%=Session("PlayerID")%>">
	<TR>
		<TD>
			<table align=center border=0 cellpadding=4 cellspacing=1>
				<TR BGCOLOR="#000000">
					<TH COLSPAN=2>Upload to ScreenShot library</TH>
				</TR>
				<tr BGCOLOR="<%=bgcone%>">
					<td valign=MIDDLE ALIGN=RIGHT width=50%><b>File:</b></td>
					<td width=50%><INPUT TYPE="FILE" NAME="FILE1" style="width: 300px;"></td>
				</tr>		
				<TR BGCOLOR="<%=bgcone%>">
					<TD>&nbsp;</TD>
					<TD>All file types accepted</TD>
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

