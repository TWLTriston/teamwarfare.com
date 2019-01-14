<%
Function XMLEncode(xml)

	If xml <> "" Then
		xml = Replace(xml,"&","&amp;")
		xml = Replace(xml,"<","&lt;")
		xml = Replace(xml,">","&gt;")
		xml = Replace(xml,Chr(34),"&quot;")
		xml = Replace(xml,Chr(39),"&apos;")
	End If
	
	XMLEncode = xml
	
End Function
%>
