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
    Below are the available BB codes that will be replaced by the forums automatically.
    <table border="0" cellspacing="0" cellpadding="0" class="cssBordered" WIDTH=100%>
		<TR bgcolor="<%=bgcblack%>">
			<TH colspan=3>Forum Codes</TD>
		</TR>
		<TR bgcolor="<%=strHeaderColor%>">
			<TH CLASS="columnheader">Description</TH>
			<TH CLASS="columnheader">You type</TH>
			<TH CLASS="columnheader">What it looks like</TH>
		</TR>
		<TR bgcolor="<%=strHighlight1%>">
			<TD BGCOLOR="<%=strColumnColor1%>"><B>Bold text</B></TH>
			<TD BGCOLOR="<%=strColumnColor2%>" ALIGN=CENTER>[b]Your text here.[/b]</TD>
			<TD BGCOLOR="<%=strColumnColor1%>" ALIGN=CENTER><B>Your text here.</B></TD>
		</TR>
		<TR bgcolor="<%=strHighlight2%>">
			<TD BGCOLOR="<%=strColumnColor1%>"><B>Italic Text</B></TH>
			<TD BGCOLOR="<%=strColumnColor2%>" ALIGN=CENTER>[i]Your text here.[/i]</TD>
			<TD BGCOLOR="<%=strColumnColor1%>" ALIGN=CENTER><i>Your text here.</i></TD>
		</TR>
		<TR bgcolor="<%=strHighlight1%>">
			<TD BGCOLOR="<%=strColumnColor1%>"><B>Underline Text</B></TH>
			<TD BGCOLOR="<%=strColumnColor2%>" ALIGN=CENTER>[u]Your text here.[/u]</TD>
			<TD BGCOLOR="<%=strColumnColor1%>" ALIGN=CENTER><u>Your text here.</u></TD>
		</TR>
		<TR bgcolor="<%=strHighlight2%>">
			<TD BGCOLOR="<%=strColumnColor1%>"><B>Quote</B></TH>
			<TD BGCOLOR="<%=strColumnColor2%>" ALIGN=CENTER>[quote]Someone's text.[/quote]</TD>
			<TD BGCOLOR="<%=strColumnColor1%>" ALIGN=CENTER><BLOCKQUOTE CLASS="forumquote">Your text here.</BLOCKQUOTE></TD>
		</TR>
		<TR bgcolor="<%=strHighlight1%>">
			<TD BGCOLOR="<%=strColumnColor1%>"><B>Images</B></TH>
			<TD BGCOLOR="<%=strColumnColor2%>" ALIGN=CENTER>[img]http://www.teamwarfare.com/forums/smilies/party.gif[/img]</TD>
			<TD BGCOLOR="<%=strColumnColor1%>" ALIGN=CENTER><img src="http://www.teamwarfare.com/forums/smilies/party.gif"></TD>
		</TR>
		<TR bgcolor="<%=strHighlight2%>">
			<TD BGCOLOR="<%=strColumnColor1%>"><B>URL Parsing</B></TH>
			<TD BGCOLOR="<%=strColumnColor2%>" ALIGN=CENTER>[url]http://www.teamwarfare.com/[/url]</TD>
			<TD BGCOLOR="<%=strColumnColor1%>" ALIGN=CENTER><a href="http://www.teamwarfare.com/">http://www.teamwarfare.com/</a></TD>
		</TR>
		<TR bgcolor="<%=strHighlight1%>">
			<TD BGCOLOR="<%=strColumnColor1%>"><B>URL Parsing #2</B></TH>
			<TD BGCOLOR="<%=strColumnColor2%>" ALIGN=CENTER>[url="http://www.teamwarfare.com"]TeamWarfare[/url]</TD>
			<TD BGCOLOR="<%=strColumnColor1%>" ALIGN=CENTER><a href="http://www.teamwarfare.com/">TeamWarfare</a></TD>
		</TR>
		</TABLE>
		</td>
    </tr>
<%Call DisplayForumFooter()%>
</table>
<% Call ContentEnd() %>
<!-- #include virtual="/include/i_footer.asp" -->
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
%>