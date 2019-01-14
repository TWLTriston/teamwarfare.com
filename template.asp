<% Option Explicit %>
<%
Response.Buffer = True

Dim strPageTitle

strPageTitle = "TWL: The Template"

Dim strSQL, oConn, oRS
Dim bgcone, bgctwo, bgcblack, bgcheader

Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")

oConn.ConnectionString = Application("ConnectStr")
oConn.Open

Call CheckCookie()
If Len(Request.QueryString("StyleID")) > 0 Then
	Session("StyleID") =  Request.QueryString("StyleID")
End if

Response.Write "Second Mod 2: " & Second(Now()) Mod 2

Dim bSysAdmin, bAnyLadderAdmin
bSysAdmin = IsSysAdmin()
bAnyLadderAdmin = IsAnyLadderAdmin()
''Response.Write "StyleID: " & Session("StyleID")
%>
<!-- #include virtual="/include/i_funclib.asp" -->
<!-- #include virtual="/include/i_header.asp" -->

    <% Call ContentStart("") %>
    <div align="center">
	    <!-- BEGIN: AdSolution-Website-Tag 4.3 : Teamwarfare.com / Teamwarfare. com_Forums_Leaderboard -->
	<script language="javascript" type="text/javascript">
	Ads_kid=0;Ads_bid=0;Ads_xl=0;Ads_yl=0;Ads_xp='';Ads_yp='';Ads_xp1='';Ads_yp1='';Ads_opt=0;Ads_par='';Ads_cnturl='';
	</script>
	<script type="text/javascript" language="javascript" src="http://a.as-us.falkag.net/dat/cjf/00/09/33/54.js"></script>
	<!-- END:AdSolution-Tag 4.3 -->
		</div>
    <% Call ContentEnd() %>
    <% Call ContentStart("") %>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call ContentEnd() %>


    <% Call ContentStart("Content w/ Header") %>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call ContentEnd() %>

    <% Call ContentNewsStart("News") %>
					<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
					<tr>
						<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
					</tr>
					</table>
	<% Call ContentNewsEnd() %>
	
	<% Call Content2BoxStart("")%>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call Content2BoxMiddle() %>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call Content2BoxEnd() %>

	<% Call Content3BoxStart("") %>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call Content3BoxMiddle1() %>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call Content3BoxMiddle2() %>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call Content3BoxEnd() %>
	<% Call Content33BoxStart("") %>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call Content33BoxMiddle() %>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call Content33BoxEnd() %>
	<% Call Content66BoxStart("") %>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call Content66BoxMiddle() %>
		<table cellspacing="0" cellpadding="0" border="0" width="100%" class="cssBordered" align="center">
		<tr>
			<td>Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Curabitur eget mauris. Donec risus odio, lacinia nonummy, posuere non, nonummy vitae, dolor. Mauris dictum, nulla ac ornare commodo, turpis mauris scelerisque orci, in dignissim magna ante vitae mi. Vestibulum lobortis leo vel lacus. In posuere elit ac enim. Mauris imperdiet, felis sit amet fringilla lacinia, massa turpis rutrum sem, in ornare dolor lectus id augue. Donec felis. Curabitur ut neque eget mi blandit venenatis. Morbi fermentum quam nec purus posuere dignissim. Morbi a dui non felis cursus suscipit. Donec bibendum, ipsum a accumsan elementum, erat lectus rhoncus sapien, sed dictum orci augue eget wisi. Phasellus eu odio adipiscing turpis dapibus venenatis. Curabitur aliquet gravida arcu. Integer ac urna accumsan wisi commodo ullamcorper. Donec eros massa, vestibulum ac, semper nec, interdum non, nisl.</td>
		</tr>
		</table>
	<% Call Content66BoxEnd() %>
	<!-- #include virtual="/include/i_footer.asp" //-->
	
<%
oConn.Close
Set oConn = Nothing
Set oRS = Nothing
Response.End
%>