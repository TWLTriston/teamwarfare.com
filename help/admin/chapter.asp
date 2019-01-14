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

Dim rulename, active, orsFAQ, sFAQName, nFaqID

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
		
		Dim sID
		sID = Request.QueryString("fldAuto")
		If sID = "" Then
			sID = 0
		Else
			sID = CInt( sID)
		End If
		oRS.Open "select * from tbl_chapter where fldAuto = " & sID, oConn, 1, 3
		Select Case Request.QueryString("action")
			Case "new"
				oRS.AddNew
				oRS("name").Value = Request.Form("name")
				oRS("rulename").Value = Request.Form("rulename")
				oRS("orderingfield").Value = "0"
				oRS("faq_fldAuto").Value = Request.QueryString("faq_fldAuto")
				oRS("GeneralRuleID").value = Request.Form("GeneralRuleID")
			Case "edit"
				oRS("rulename").Value = Request.Form("rulename")
				oRS("name").Value = Request.Form("name")
				oRS("GeneralRuleID").value = Request.Form("GeneralRuleID")
			Case "del"
				oRS("Isactive").Value = 0
			'oConn.Execute "delete from tbl_question where chapter_fldAuto=" & sID
		End Select
		oRS.Update
		oRS.Close
		Response.Redirect "chapters.asp?faq_fldAuto=" & Request.QueryString("faq_fldAuto")
	End If
Dim name, descr, generalruleid

If Request.QueryString("action") = "edit" Then
	Set oRS2 = oConn.Execute( "select * from tbl_chapter where fldAuto = " & Request.QueryString("fldAuto") )
	name = oRS2("name").Value
	rulename = oRS2("rulename").Value
	generalruleid = ors2("generalruleid").value
	oRS2.Close
Else
End If

Set oRSFAQ = oConn.Execute ("select * from tbl_faq where fldAuto = " & Request.QueryString("faq_fldAuto") )
sFAQName = oRSFAQ("name")
nFaqId = Request.QueryString("faq_fldAuto")
oRSFAQ.Close    
Set oRSFAQ = Nothing

Call ContentStart(sFAQName & "Add/Modify Topic")
%>
    <table width="90%" border="0">
	<tr>
		<td>
        <%
        Dim sURL
        sURL = "chapter.asp?save=yes&action=" & Request.QueryString("action")
        sURL = sURL & "&fldAuto=" & Request.QueryString("fldAuto") 
        sURL = sURL & "&faq_fldAuto=" & Request.QueryString("faq_fldAuto") 
        %>
        <form method="POST" action="<%=sURL%>">
        	General Rule Set: <select name="GeneralRuleID" id="GeneralRuleID">
        	<option value="0" <% if generalruleid = "0" Then response.write "selected=""selected"" "end if %>>None</option>
        	<option value="190" <% if generalruleid = "190" Then response.write "selected=""selected"" "end if %>>League General Rules</option>
        	<option value="45" <% if generalruleid = "45" Then response.write "selected=""selected"" "end if %>>Rung Ladder General Rules</option>
        	<option value="380" <% if generalruleid = "380" Then response.write "selected=""selected"" "end if %>>Power Rating Ladder General Rules</option>
        	<option value="483" <% if generalruleid = "483" Then response.write "selected=""selected"" "end if %>>America's Army General Rules</option>
        	</select><br />
        	
          <p align="left">FAQ Name: <input type="text" name="name" size="40" value="<%=name%>"></p>
          <p align="left">Rule Page Name: <input type="text" name="rulename" size="40" value="<%=rulename%>"></p>
          <p align="left"><input type="submit" value="Submit" name="B1"></p>
        </form>
		<a href="chapters.asp?faq_fldAuto=<%=nFaqId%>">Back to current Section ( <%=sFAQName%> )&nbsp;</a><br>
		<a href="default.asp">Back to home&nbsp;</a><br>
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