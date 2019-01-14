<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Files"

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

'If Not(bAnyLadderAdmin or bSysAdmin) Then
'	oConn.Close
'	Set oConn = Nothing
'	Response.Clear
'	Response.Redirect "/errorpage.asp?error=3"
'End If
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->

<% Call ContentStart("TWL ScreenShot Library") %>


<table width=500 cellspacing=0 cellpadding=0 border=0 BGCOLOR="#444444">
 <TR><TD>
 <table width=100% Cellspacing=1 cellpadding=2 border=0>
 <TR BGCOLOR="#000000">
  <TH COLSPAN=2><A HREF="FileAdd.asp">Add New ScreenShot</A></TH>
 </TR>
 <tr bgcolor="#000000">
 	<th>File</th>
 	<th>Size</th>
 </tr>
	
	<%
	Dim oFS, oFolder, oFile
	Set oFS = Server.CreateObject ("Scripting.FileSystemObject")
	Set oFolder = oFS.GetFolder("d:\TWLFiles\ScreenShots")
	bgc = bgcone
	For Each oFile in oFolder.Files
		if bgc = bgcone then
			bgc = bgctwo
		else
			bgc = bgcone
		end if
		%>
		<tr>
			<td bgcolor="<%=bgc%>"><a href="img/<%=oFile.Name%>"><%=oFile.Name%></a></td>
			<td bgcolor="<%=bgc%>" align="right"><%=CalculateSizeDisplay(oFile.Size)%></td>
		</tr>
		<%
	Next	
	%>
	 </TABLE>
	</TD></TR>
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
	Elseif intSize < 10000000 Then
		CalculateSizeDisplay = FormatNumber(intSize/1000, 0,0,0,-1) & " KB"
	Else
		CalculateSizeDisplay = FormatNumber(intSize/1000000, 0, 0, 0, -1) & " MB"
	End If
End Function
%>

