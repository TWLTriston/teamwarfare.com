<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add an Ladder Option"

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

Dim bIsEdit, strOptionName, chrActive
Dim chrSide, chrSelectedBy, intLadderID
Dim intOptionID, strMethod, strVerbage
Dim intMapNumber

bIsEdit = cBool(Request.QueryString("IsEdit"))
intLadderID = Request("LadderID")

If bIsEdit Then
	intOptionID = Request("OptionID")
	
	strSQL = "SELECT * FROM tbl_ladder_options WHERE OptionID='" & intOptioNID & "'"
	oRS.Open strSQL, oConn
	If Not(oRs.EOF AND oRS.BOF) Then
		strOptionName	= oRS.Fields("OptionName").Value 
		chrSide			= oRS.Fields("SideChoice").Value
		chrSelectedBy	= oRS.Fields("SelectedBy").Value
		intMapNumber	= oRS.Fields("MapNumber").Value
	End If
	oRS.Close
	strMethod = "Edit"
	strVerbage = "Edit Ladder Option"	
Else 
	strMethod = "New"
	strVerbage = "Create Ladder Option"
End If

strPageTitle = "TWL: " & strVerbage
Dim intCounter

if not(bSysAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	response.clear
	response.redirect "errorpage.asp?error=3"
End If

%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart(strVerbage) %>
    <form name=frmAddOption action=option_saveItem.asp method=post>
	<INPUT TYPE=HIDDEN NAME=LadderID VALUE="<%=intLadderID%>">
	<INPUT TYPE=HIDDEN NAME=OptionID VALUE="<%=intOptionID%>">
	<INPUT TYPE=HIDDEN NAME=Method VALUE="<%=strMethod%>">
	<INPUT TYPE=HIDDEN NAME=SaveType VALUE="AddOption">
		<table border=0 align=center cellspacing=0 CELLPADDING=0 BGCOLOR="#444444">
		<TR><TD>
		<table border=0 align=center width="100%" cellspacing=1 CELLPADDING=2>
		<TR BGCOLOR="#000000">
			<TH COLSPAN=2>Configure Option</TH>
		</TR>
		<TR BGCOLOR=<%=bgcone%>>
			<TD ALIGN=RIGHT>Option Name:</TD>
			<TD><INPUT TYPE=TEXT NAME=OptionName VALUE="<%=Server.HTMLEncode("" & strOptionName)%>" SIZE=20></TD>
		</TR>   
		<TR BGCOLOR=<%=bgctwo%>>
			<TD ALIGN=RIGHT>Side Choice?</TD>
			<TD><SELECT NAME=SideChoice>
				<OPTION VALUE="N" <% If chrSide = "N" Then Response.Write " SELECTED " %>>No</OPTION>
				<OPTION VALUE="Y" <% If chrSide = "Y" Then Response.Write " SELECTED " %>>Yes</OPTION>
			</SELECT></TD>
		</TR>
		<TR BGCOLOR=<%=bgcone%>>
			<TD ALIGN=RIGHT>Selected By:</TD>
			<TD><SELECT NAME=SelectedBy>
				<OPTION VALUE="A" <% If chrSelectedBy = "A" Then Response.Write " SELECTED " %>>Attacker</OPTION>
				<OPTION VALUE="D" <% If chrSelectedBy = "D" Then Response.Write " SELECTED " %>>Defender</OPTION>
				<OPTION VALUE="R" <% If chrSelectedBy = "R" Then Response.Write " SELECTED " %>>Random</OPTION>
			</SELECT></TD>
		</TR>
		<TR BGCOLOR=<%=bgctwo%>>
			<TD ALIGN=RIGHT>Map Number:</TD>
			<TD><SELECT NAME=MapNumber>
				<OPTION VALUE="0" <% If intMapNumber = "0" Then Response.Write " SELECTED " %>>Not map specific</OPTION>
				<OPTION VALUE="1" <% If intMapNumber = "1" Then Response.Write " SELECTED " %>>Map 1</OPTION>
				<OPTION VALUE="2" <% If intMapNumber = "2" Then Response.Write " SELECTED " %>>Map 2</OPTION>
				<OPTION VALUE="3" <% If intMapNumber = "3" Then Response.Write " SELECTED " %>>Map 3</OPTION>
				<OPTION VALUE="4" <% If intMapNumber = "4" Then Response.Write " SELECTED " %>>Map 4</OPTION>
				<OPTION VALUE="5" <% If intMapNumber = "5" Then Response.Write " SELECTED " %>>Map 5</OPTION>
			</SELECT></TD>
		</TR>
		<TR BGCOLOR=<%=bgcone%>>
			<TD ALIGN=CENTER COLSPAN=2><INPUT TYPE=SUBMIT VALUE="Save Information"></TD>
		</TR>
		</table>
    </TD></TR>
    </TABLE>
    </form>
    <BR><BR>
    <B>Side Choice:</B> Means that the seleciton will reverse depending on which team you are <BR>looking at, otherwise it will not, and the option
			will be considered a "map" option, like "time of day", or "fog".
    
<% Call ContentEnd() 

If bIsEdit Then
	Call ContentStart("Configure Drop Down Values")
	%>
	<FORM ACTION="option_saveitem.asp" METHOD="post">
	<INPUT TYPE="HIDDEN" VALUE="<%=intOptionID%>" NAME=OptionID>
	<INPUT TYPE="HIDDEN" VALUE="<%=intLadderID%>" NAME=LadderID>
	<INPUT TYPE="HIDDEN" VALUE="OptionValue" NAME=SaveType>
	
	<table border=0 align=center cellspacing=0 CELLPADDING=0 BGCOLOR="#444444" WIDTH="50%">
	<TR><TD>
	<table border=0 align=center width="100%" cellspacing=1 CELLPADDING=2>
	<TR BGCOLOR="#000000">
		<TH COLSPAN=2>Drop Down Values</TH>
	</TR>
	<TR BGCOLOR="#000000">
		<TH>Name</TH>
		<TH>Active</TH>
	</TR>
	<%
	strSQL = "SELECT * FROM tbl_ladder_option_value WHERE OptionID='" & intOptionID & "' ORDER BY Active DESC, ValueName ASC"
	oRS.Open strSQL, oConn
	If Not(oRS.EOF AND oRS.BOF) Then
		Do While Not(oRS.EOF)
			If bgc = bgcone Then
				bgc = bgctwo
			Else
				bgc = bgcone
			End If
			%>
			<TR BGCOLOR="<%=bgc%>">
				<INPUT TYPE="HIDDEN" NAME=OptionValueID VALUE="<%=oRS.Fields("OptionValueID").Value%>">
				<TD><INPUT TYPE=TEXT NAME=ValueName VALUE="<%=Server.HTMLEncode("" & oRS.Fields("ValueName").Value)%>" SIZE=40></TD>
				<TD><SELECT NAME=Active>
					<OPTION VALUE="Y" <% If oRS.Fields("Active").Value = "Y" Then Response.Write " SELECTED "%>>Yes</OPTION>
					<OPTION VALUE="N" <% If oRS.Fields("Active").Value = "N" Then Response.Write " SELECTED "%>>No</OPTION>
				</SELECT>
				</TD>
			</TR>
			<%
			oRS.MoveNext
		Loop
	End If
	oRS.Close
	For intCounter = 1 to 20
		If bgc = bgcone Then
			bgc = bgctwo
		Else
			bgc = bgcone
		End If
		%>
		<TR BGCOLOR="<%=bgc%>">
			<INPUT TYPE="HIDDEN" NAME=OptionValueID VALUE="NEW">
			<TD><INPUT TYPE=TEXT NAME=ValueName VALUE="" SIZE=40></TD>
			<TD><SELECT NAME=Active>
				<OPTION VALUE="Y">Yes</OPTION>
				<OPTION VALUE="N">No</OPTION>
			</SELECT>
			</TD>
		</TR>
		<%
	Next
	%>	
	<TR BGCOLOR="#000000">
		<TD ALIGN=CENTER COLSPAN=2><INPUT TYPE=SUBMIT VALUE="Save Values"></TD>
	</TR>
	</TABLE>
	</TD></TR>
	</TABLE>
	</FORM>
	<%
	Call ContentEnd() 
End If
%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>