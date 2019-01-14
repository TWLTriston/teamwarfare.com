<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Help Administration"

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

Dim oRSChapters, sName, rulesName
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

<% 
if not (bAnyLadderAdmin or bSysAdmin) then
	oConn.Close
	Set oConn = Nothing
	Set oRs = Nothing
	Response.Clear
	Response.Redirect "/errorpage.asp?error=3"
end if

	If Request.QueryString("save")="yes" Then
		Set oRS = Server.CreateObject("ADODB.Recordset")
		'Change ordering...
		Dim sID, nFldAuto
		sID = Request.QueryString("faq_fldAuto")
		sID = CInt( sID)
		oRS.Open "select * from tbl_chapter where faq_fldAuto = " & sID, oConn, 1, 3
		While Not oRS.EOF
			nFldAuto = oRS("fldAuto")
			oRS("orderingfield") = Request.Form("order" & nFldAuto)
			oRS.Update
			oRS.MoveNext
		Wend
		ors.NextRecordset 
	End If

Call ContentStart("Topics For: " & Request.QueryString ("Section"))
%>
<script language="javascript">
function fConfirmDelete(fldAuto, faq_fldAuto) {
	if (confirm("Are you sure you want to delete this ruleset??")) {
		window.location="chapter.asp?save=yes&action=del&faq_fldAuto="+faq_fldAuto+"&fldAuto="+fldAuto;
	}
}
</script>
           <form method="POST" action="<%=sURL%>" name="FrontPage_Form1">
        <table border="0" width="90%" bgcolor="#444444" cellspacing="0" cellpadding="0">
          <tr>
            <td>
			<%
			Dim sURL
			sURL = "chapters.asp?faq_fldAuto=" & Request.QueryString("faq_fldAuto") & "&save=yes"
			Set oRSChapters = oConn.Execute("select * from tbl_chapter where faq_fldAuto=" & Request.QueryString("faq_fldAuto") & " and isactive = 1 order by orderingfield" )
			%>                             
            <table border="0" width="100%" cellspacing="1" cellpadding="4">
              <tr> 
                <th bgcolor="#000000">Topic Name</th>
                <th bgcolor="#000000">Order Number</th>
                <th bgcolor="#000000">Actions</th>
              </tr>
			<%
			Dim nVal 
			nVal = 1
			While Not oRSChapters.EOF
				sName = "order" & oRSChapters("fldAuto").Value
				rulesname = oRSChapters("RuleName")
			%>                        
            <tr>
              <td bgcolor="<%=bgcone%>"><%=oRSChapters("name")%> <br /><b><%=orschapters("rulename")%></b></td>
              <td align="center" bgcolor="<%=bgctwo%>">
              <input type="text" name="<%=sName%>" size="3" id="<%=sName%>" value="<%=nVal%>"></td>
              <td align="center" bgcolor="<%=bgcone%>"><a href="chapter.asp?action=edit&amp;faq_fldAuto=<%=Request.QueryString("faq_fldAuto")%>&amp;fldAuto=<%=oRSChapters("fldAuto")%>">Modify</a> 
              <% if bSysAdmin Then %>
              -<a href="javascript:fConfirmDelete('<%=oRSChapters("fldAuto")%>','<%=request.querystring("faq_fldauto")%>');"> Delete</a>
              <% End If %>
                - <a href="questions.asp?chapter_fldAuto=<%=oRSChapters("fldAuto")%>"> Questions</a>
				<% if rulesname <> "" then %>
                - <a href="/rules.asp?set=<%=server.urlencode(rulesname)%>">View as Rules</a></b>
                <% end if %></td>
            </tr> 
			<%
				nVal = nVal + 1
				oRSChapters.MoveNext
			Wend
			oRSChapters.Close
			Set oRSChapters = Nothing
			%>                        
            <tr>
            	<td bgcolor="#000000" colspan="3">
            		<br />
              		&nbsp;&nbsp;&nbsp;<input type="submit" value="Change Topic ordering" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;
              		<a href="chapter.asp?action=new&amp;faq_fldAuto=<%=Request.QueryString("faq_fldAuto")%>">Add new Topic</a>
              		<br /><br />
					<a href="/help/admin/default.asp"><b>Help Admin</b></a>
				</td>
			</tr>
            </table>
			</td>
		</tr>
		</table>
         </form>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS2 = Nothing
Set oRS = Nothing
%>