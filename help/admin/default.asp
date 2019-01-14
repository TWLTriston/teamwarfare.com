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
Call ContentStart("TWL Help Administration")
%>
    <table width="90%" border="0">
	<tr>
		<td>
		<table border="0" width="100%">
			<tr>
				<td width="100%">
					<p class=small>It works like this....<br>
					<b>SECTION(s)</b>: Any Section you want (Rules, FAQ, Help, etc...)<br>
					<b>TOPIC(s)</b>: Subtopics for a section(General Rules, CTF Rules, Arena Rules, etc...)<br>
					<b>QUESTIONS</b>: Headings under TOPICS(CTF Rules::Unsportsman-like Conduct, etc...)<br>
					<b>CONTENT</b>: The RULE, FAQ, HELP, etc...itself(..."To join a team go to...")<br>
					If you have questions or Comments, contact DannyBoy. 
					</p>                        
                </td>
              </tr>
        </table>
        <table border="0" cellspacing="0" cellpadding="0" bgcolor="#444444" width="100%">
        <tr><td>
        <table border="0" width="100%" cellspacing="1" cellpadding="4">
			<tr>
				<th colspan=2 bgcolor="#000000"><b><i>These Sections are currently available:</i></b></Th>
			</TR>
			<%	
				if bsysadmin then
					oRS.open "select * from tbl_faq ORDER BY name ASC ", oconn
				else
					oRS.open "select * from tbl_faq where adminedit = 0 ORDER BY name ASC ", oconn
				end if				
				bgc = bgcone
				dim active
				while not oRS.EOF
					active = "Active"
					if ors("isActive") = 0 then
						active = "Inactive"
					end if
			%>                        
            <tr>
              <td width="50%" bgcolor="<%=bgc%>"><b><%=oRS("name")%> (<%=active%>)</b></td>
              <td width="50%" bgcolor="<%=bgc%>"><b><a href="/help/admin/faq.asp?fldAuto=<%=oRS("fldAuto")%>&action=edit">Modify</a>
                <% If bSysAdmin Then %>
                - <a href="/help/admin/faq.asp?fldAuto=<%=oRS("fldAuto")%>&save=yes&action=del">Delete</a> 
                <% End If %>
                - <a href="/help/admin/chapters.asp?faq_fldAuto=<%=oRS("fldAuto")%>&section=<%=server.urlencode(ors("name"))%>">Topics/Content</a>
            </tr>
			<%	
					If bgc=bgcone Then
						bgc =bgctwo
					Else
						bgc=bgcone
					End if
					oRS.MoveNext
				Wend
			%>                        
		</table>
		</td></tr>
		</table>
		<p class=small><a href="/help/admin/faq.asp?action=new">Add new Section</a></p>
		</td>
	</tr>
	</table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS2 = Nothing
Set oRS = Nothing
%>