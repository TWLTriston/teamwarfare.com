<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: News"

Dim strSQL, oConn, oRS, oRS2
Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oRS2 = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()

Dim bSysAdmin, bAnyLadderAdmin, bTeamFounder, bLeagueAdmin, bTeamCaptain, bLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()

If Not(bSysAdmin or bAnyLadderAdmin) Then
	oConn.Close
	Set oConn = Nothing
	Set oRS = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
End If	

Dim Action, intGameID
action=Request.QueryString("action")
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #Include virtual="/include/i_header.asp" -->
<SCRIPT LANGUAGE="javascript">
<!--
	function writeheadline(strData) {
		divheadline.innerHTML = strData
	}
	function writecontent(strData) {
		divcontent.innerHTML = strData
	}
//-->
</SCRIPT>
<%
Call ContentStart(Action & " News")
%>
<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 ALIGN=CENTER BGCOLOR="#444444">
<TR><TD>
<%
	strSQL = "select * from tbl_News where NewsID='" & Request.QueryString("newsid") & "'"
	oRs.Open strSQL, oConn	
	if not (ors.EOF and ors.BOF) then
		if action = "Delete" then
			Response.Write "<table width=600 border=0 cellpadding=2 cellspacing=1 align=center>"
			Response.Write "<tr bgcolor=#000000><tH>" & Server.HTMLEncode(ors.Fields("NewsHeadline").Value) & " (" & ors.Fields("NewsDate").Value & " by " & Server.HTMLEncode(ors.Fields("NewsAuthor").Value) & ")</tH></tr>"
			Response.Write "<tr bgcolor=" & bgctwo & "><td>" & ors.Fields("NewsContent").Value & " </td></tr>"
			Response.Write "<tr bgcolor=" & bgcone & "><td align=center><a href=saveitem.asp?SaveType=DeleteNews&newsid=" & Request.QueryString("newsid") & ">Delete Entry</a> - <a href=newsdesk.asp>Return to the News Desk</a></td></tr></table>"
		end if
		if action = "Edit" then
			%>
			<table align=center width=400 border=0 cellspacing=1 cellpadding=2>
			<form name=frmNewsEdit action=saveitem.asp method=post>
			<TR BGCOLOR="#000000"><TH COLSPAN=2>Edit News</TH></TR>
			<tr bgcolor=<%=bgcone%> height=30><td ><B>Headline:&nbsp;</B>&nbsp;&nbsp;&nbsp;<input onkeypress="javascript:writeheadline(this.form.headline.value)" onchange="javascript:writeheadline(this.form.headline.value)" maxlength=30 type=text name=headline class=text value="<%=ors.Fields("NewsHeadline").Value%>" style="width:300">

			<tr bgcolor=<%=bgctwo%> height=200><td align=center><textarea name=newscontent onkeypress="javascript:writecontent(this.form.newscontent.value)" onchange="javascript:writecontent(this.form.newscontent.value)" cols=80 rows=15><%=Server.HTMLEncode(ors.Fields("NewsContent").Value)%></textarea></td></tr>
		<tr bgcolor=<%=bgcone%>><td align=center>Relevant Game:&nbsp;<SELECT NAME=NewsType Class=text><option value="0">Announcement</option>
				<%
					strSQL = "SELECT GameID, GameName FROM tbl_Games WHERE GameID > 0 ORDER BY GameName ASC "
					oRS2.Open strSQL, oConn
					If Not(oRS2.EOF AND oRS2.BOF) Then
						Do While Not(oRS2.EOF)
							Response.Write "<OPTION VALUE=""" & oRS2.Fields("GameID").Value & """ "
							If cStr(oRS2.Fields("GameID").Value  & "") = cStr(oRS.Fields("NewsType").Value & "") Then
								Response.Write " SELECTED "
							End If
							Response.Write ">" & Server.HTMLEncode(oRS2.Fields("GameName").Value & "") & "</OPTION>" & vbCrLf
							oRs2.MoveNext
						Loop					
					End If
					oRs2.NextRecordset
					%>
					</SELECT></td></tr>
			<tr bgcolor=<%=bgctwo%>><td align=center><input type=hidden name=SaveType value=EditNews><input type=hidden name=NewsId value=<%=Request.QueryString("newsid")%>>
				<input type=button name=Preview value="Preview" class=bright onclick="javascript:writecontent(this.form.newscontent.value);javascript:writeheadline(this.form.headline.value)">&nbsp;&nbsp;&nbsp;
				<input type=submit name=submit1 value=submit class=bright>
				</td></tr>
			</form>
			</table>

			<%
		end if
	end if
	if action="Add" then
		%>
			<table align=center width=400 border=0 cellspacing=1 cellpadding=2>
			<form name=frmNewsAdd action=saveitem.asp method=post>
			<TR BGCOLOR="#000000"><TH COLSPAN=2>Add News</TH></TR>
			<tr bgcolor=<%=bgcone%> height=30><td align=right><B>Headline:&nbsp;</B></td><td align=left>&nbsp;&nbsp;&nbsp;<input onkeypress="javascript:writeheadline(this.form.headline.value)" onchange="javascript:writeheadline(this.form.headline.value)" maxlength=30 type=text name=headline class=text style="width:300">
			</td></tr>
			<tr bgcolor=<%=bgctwo%> height=200><td colspan=2 align=center><textarea onkeypress="javascript:writecontent(this.form.content.value)" onchange="javascript:writecontent(this.form.content.value)" name=content cols=80 rows=15></textarea>
			</td></tr>
		<tr bgcolor=<%=bgcone%>><td align=right>Relevant Game:</td><td>&nbsp;<SELECT NAME=NewsType Class=text><option value="0">Announcement</option>
				<%
					strSQL = "SELECT GameID, GameName FROM tbl_Games WHERE GameID > 0 ORDER BY GameName ASC "
					oRS2.Open strSQL, oConn
					If Not(oRS2.EOF AND oRS2.BOF) Then
						Do While Not(oRS2.EOF)
							Response.Write "<OPTION VALUE=""" & oRS2.Fields("GameID").Value & """ "
							Response.Write ">" & Server.HTMLEncode(oRS2.Fields("GameName").Value & "") & "</OPTION>" & vbCrLf
							oRs2.MoveNext
						Loop					
					End If
					oRs2.NextRecordset
					%>
					</SELECT></td></tr>
			<tr bgcolor=<%=bgctwo%>><td align=center colspan=2><input type=hidden name=SaveType value=AddNews>
					<input type=hidden name=NewsId value=<%=Request.QueryString("newsid")%>>
					<input type=submit name=submit1 value="Submit News" class=bright>
			</td></tr>
			</form>
			</table>
		<%	 
	end if
%>
</TD></TR>
</TABLE>
<%
Call ContentEnd()

%>
<script type="text/javascript" src="js/tinymce3/tiny_mce.js"></script>
<script type="text/javascript">
tinymce.init({
        theme : "advanced",
        mode : "textareas"
});
</script>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRs = Nothing
Set oRs2 = Nothing
%>

