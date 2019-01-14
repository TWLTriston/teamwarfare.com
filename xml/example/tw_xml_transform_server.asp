<%@ LANGUAGE="VBSCRIPT" ENABLESESSIONSTATE = FALSE %>
<%
option explicit
Response.Buffer = True '-- Buffer ASP Page Until Server-Side Processing Complete 
%>

<%
'*-------------------------------------------- 
'*
'* FILENAME: tw_xml_transform.asp
'*
'* DESCRIPTION
'* -----------
'* This page processes xml to html using xsl all on the server.
'*
'* INPUT PARAMETERS:
'* ----------------
'* Description of parameters passed into and out of page
'* I - <Parameter1> - <Source1> - <Description1>
'* O - <Parameter2> - <Source2> - <Description2>
'*--------------------------------------------
%>

<%
'-- Declarations
	dim xml, xsl
	
'-- Load the XML
	set xml = Server.CreateObject("Microsoft.XMLDOM")
	xml.async = false
	'xml.load(Server.MapPath("TFarena_XML.xml"))
	xml.load("http://www.teamwarfare.com/xml/viewteam.asp?team=Team+Fusion+Arena&extension=.xml")

'-- Load the XSL
	set xsl = Server.CreateObject("Microsoft.XMLDOM")
	xsl.async = false
	xsl.load(Server.MapPath("tw_xsl_transform.xsl"))

'-- Transform the file
	Response.Write(xml.transformNode(xsl))
%>
