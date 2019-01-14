<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Add In Game ID"

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

If Not(Session("LoggedIn")) Then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=2"
End If

%>

<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% Call ContentStart("") %>
<script language="javascript" type="text/javascript">
function fValidateIdentifier(oForm) {
	var arrValidation = new Array()
	<%
	strSQL = "SELECT IdentifierName, IdentifierID, IdentifierFormat FROM tbl_identifiers WHERE IdentifierActive = 1 ORDER BY IdentifierName ASC "
	oRs.Open strSQL, oConn
	If Not(oRs.EOF AND oRs.BOF) Then
		Do While Not(oRs.EOF)
			%>
			arrValidation.push(new Array("<%=oRs.FieldS("IdentifierID").Value%>", <%=oRs.FieldS("IdentifierFormat").Value%>));
			<%
			oRs.MoveNext
		Loop
	End if
	oRs.NextRecordSet
	%>
	var intIdentifierID = oForm.selIdentifierID.options[oForm.selIdentifierID.selectedIndex].value;
	for (i=0;i<arrValidation.length;i++) {
		if (Number(arrValidation[i][0]) == intIdentifierID) {
			regStr = arrValidation[i][1];
		}
	}
	var oRegEx = new RegExp(regStr);
	var strIdentifier = oForm.txtIdentifierValue.value;
	if (oRegEx.test(strIdentifier)) {
		oForm.submit();
	} else {
		alert("Error:\n'" + strIdentifier + "' is not a valid " + oForm.selIdentifierID[oForm.selIdentifierID.selectedIndex].text + ".");
	}
}
</script>
	<%
	IF Len(Request.QuerySTring("InUse")) <> 0 Then
		Dim sPlayerHandle
		strSQL = "SELECT PlayerHandle FROM tbl_players WHERE PlayerID = '" & CheckString(Request.QuerySTring("InUse")) & "'"
		oRs.Open strSQL, oConn
		If Not(oRs.EOF AND oRs.BOF) Then
			sPlayerHandle = oRs.Fields("PlayerHandle").Value
		End If
		oRs.NextrecordSet
		%>
		That identifier is already assigned to another account: <a href="viewplayer.asp?player=<%=Server.URLEncode(sPlayerHandle)%>"><%=Server.HTMLEncode(sPlayerHandle)%></a>.<br />
		<br />
		<%
	End If
	%>
		
		
	<table align=center border=0 CELLSPACING=0 cellpadding=0  BGCOLOR="#444444">
	<form name="frmAntiSmurf" id="frmAntiSmurf" action="saveitem.asp" method="post">
	<input type="hidden" name="hdnPlayer" id="hdnPlayer" value="<%=Server.HTMLEncode(Request.QueryString("Player"))%>" />
	<input type=hidden name=SaveType value="AntiSmurfAdd">
	<tr><td>
	<table align=center border=0 cellspacing=1 cellpadding=4 width="100%">
	<tr>
		<th colspan="2" bgcolor="#000000">Add In Game Identifier</th>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right"><b>ID Type:</b></td>
		<td bgcolor="<%=bgctwo%>">
			<select name="selIdentifierID" id="selIdentifierID">
			<%
			strSQL = "SELECT IdentifierName, IdentifierID FROM tbl_identifiers WHERE IdentifierActive = 1 ORDER BY IdentifierName ASC "
			oRs.Open strSQL, oConn
			If Not(oRs.EOF AND oRs.BOF) Then
				Do While Not(oRs.EOF)
					Response.Write "<option value=""" & oRs.Fields("IdentifierID").Value & """"
					If CStr(oRs.Fields("IdentifierID").Value & "") = CStr(Request.QueryString("Identifier")) Then
						Response.Write " selected=""selected"""
					End If
					Response.Write ">" & Server.HTMLEncode(oRs.Fields("IdentifierName").Value & "") & "</option>" & VbCrLf
					oRs.MoveNext
				Loop
			End if
			oRs.NextRecordSet
			%>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="<%=bgcone%>" align="right"><b>Relevant Value:</b></td>
		<td bgcolor="<%=bgctwo%>"><input type="text" name="txtIdentifierValue" id="txtIdentifierValue" size="20" value="<%=Request.QuerySTring("Relevant")%>" /></td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#000000" align="center"><input type="button" onclick="fValidateIdentifier(this.form);" value="Add Identifier" /></td>
	</tr>
	</form>
	</table>
	</td></tr>
	</table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>