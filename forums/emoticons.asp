<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: Forum Codes"

Dim bgcone, bgctwo, bgc, bgcheader, bgcblack

Dim oConn, oRS, strSQL
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open Application("ConnectStr")

Set oRS = Server.CreateObject ("ADODB.RecordSet")



Call CheckCookie
Dim bSysAdmin, bAnyLadderAdmin
Dim blnLoggedIn, blnSysAdmin
blnLoggedIn = Session("LoggedIn")
blnSysAdmin = IsSysAdmin()
bSysAdmin = blnSysAdmin
%>
<!-- #Include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->
<%
Dim strHeaderColor, strHighlight1, strHighlight2
Dim strBGC
strHeaderColor	= bgcheader
strHighlight1	= bgcone
strHighlight2	= bgctwo

Dim strColumnColor1, strColumnColor2
strColumnColor1 = strHighlight1
strColumnColor2 = strHighlight2
%>
<% Call ContentStart("") %>
<table width="97%" border="0">
  <tr>
    <td align=center>
	Below are the available emoticons codes that will be replaced by the forums automatically.
    <table border="0" cellspacing="0" cellpadding="0" class="cssBordered" WIDTH=100%>
		<TR bgcolor="<%=bgcblack%>">
			<TH colspan=3>Emoticons</TD>
		</TR>
		<TR bgcolor="<%=strHeaderColor%>">
			<TH CLASS="columnheader">Emoticon Name</TH>
			<TH CLASS="columnheader">You type</TH>
			<TH CLASS="columnheader">Image Replacement</TH>
		</TR>
		<%
		strSQL = "SELECT EmoticonName, EmoticonSearch, EmoticonImage FROM tbl_emoticons ORDER BY  EmoticonName ASC"
		oRS.Open strSQL, oConn 
		If Not(oRS.EOF and oRS.BOF) Then
			strBGC = strHighlight1
			Do While Not(oRS.EOF)
				Response.write "<TR BGCOLOR=""" & strBGC & """>"
				Response.Write "<TD BGCOLOR=""" & strColumnColor1 & """ VALIGN=TOP><B>" & oRS.Fields("EmoticonName").Value & "</B></TD>"
				Response.Write "<TD BGCOLOR=""" & strColumnColor2 & """ VALIGN=TOP align=CENTER>" & Server.htmlencode(oRS.Fields("EmoticonSearch").Value & "") & "</TD>"
				Response.Write "<TD BGCOLOR=""" & strColumnColor1 & """ ALIGN=CENTER>" & oRS.Fields("EmoticonImage").Value & "</TD>"
				Response.Write "</TR>"
				If strBGC = strHighlight1 Then
					strBGC = strHighlight2
				Else
					strBGC = strHighlight1
				End If
				oRS.MoveNext 
			Loop
		End If
		oRS.Close
		%>
		</TABLE>
     </td>
</tr>
<%Call DisplayForumFooter()%>
</table>
<%Call ContentEnd()%>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>