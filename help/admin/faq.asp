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

Dim Name, active, descr
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
		oRS.Open "select * from tbl_faq where fldAuto = " & sID, oConn, 1, 3
		Select Case Request.QueryString("action")
			Case "new"
				oRS.AddNew
				oRS("name").Value = Request.Form("name")
				oRS("descr").Value = Request.Form("descr")
				oRS("isactive").Value = Request.Form("isactive")
			Case "edit"
				oRS("name").Value = Request.Form("name")
				oRS("descr").Value = Request.Form("descr")
				oRS("isactive").Value = Request.Form("isactive")
			Case "del"
				ors("IsActive").Value = 0
				'oRS.Delete
				'oConn.Execute "delete from tbl_chapter where faq_fldAuto=" & sID
				'oConn.Execute "delete from tbl_question where chapter_fldAuto not in (select fldAuto from tbl_chapter)"
		End Select
		oRS.Update
		oRS.Close
	Set oRS = Nothing
	oConn.Close
	Set oConn = Nothing
	Response.Redirect "default.asp"
	End If
If Request.QueryString("action") = "edit" Then
	Set oRS2 = oConn.Execute( "select * from tbl_faq where fldAuto = " & Request.QueryString("fldAuto") )
	name = oRS2("name").Value
	active = oRS2("isactive").Value
	descr = oRS2("descr").Value
	oRS2.Close
Else
End If

Call ContentStart("Help Section Admin")

%>
    <table width="90%" border="0">
	<tr>
		<td>
                        <p align="left"><br>
                        <%
                        Dim sURL
                        sURL = "faq.asp?save=yes&action=" & Request.QueryString("action")
                        If Request.QueryString("action") <> "new" Then
                        	sURL = sURL & "&fldAuto=" & Request.QueryString("fldAuto") 
                        End If
                        %>
                        <form method="POST" action="<%=sURL%>">
                          <p align="left">Name: <input type="text" name="name" size="40" value="<%=name%>"></p>
                          <p align="left">Description: <input type="text" name="descr" size="40" value="<%=descr%>"><br>
                          <p align="left">Active: <select name=isactive>
                          <option <% if active=0 then Response.write " selected " end if %>value="0">Inactive</option>
                          <option <% if active=1 then Response.write " selected " end if %>value="1">Active</option></select>
							</p>
                          <p align="left"><input type="submit" value="Submit" name="B1"></p>
                        </form>
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