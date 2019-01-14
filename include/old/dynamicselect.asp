<SCRIPT RUNAT="server" LANGUAGE="VBScript">

SUB SelectBox(rsOptions, strName, strValue, strDisplay, strSecondary, strArray, ladder)
	Dim endline, insString
'Creates a Select box from a recordset
'rsOptions	the recordset which provides the data to display;
'strName	the name we assign to the list within the HTML form;
'strValue	the field for the value to be passed when the form is submitted;
'strDisplay	the field to be displayed for each option in the list;
'strSecondary		the Select list that will change in response to changes in 
'			the list being created here;
'strArray	the array that will populate the list identified by strSecondary.

	if strArray=" " then
		endline=">" & vbCRLF
	else
		endline="onChange=" & chr(34) & "ChangeOptions('" & Replace(strName, "'", "\'") & "', '" & Replace(strSecondary, "'", "\'") & "', '" & Replace(strArray , "'", "\'") & "')" & chr(34) & ">" & vbCRLF
	end if
   
	Response.Write vbCRLF & "<SELECT name=" & chr(34) & Replace(strName, "'", "\'") & chr(34) & " id=" & chr(34) & Replace(strName, "'", "\'") & chr(34) & " size=1 style='width:200' " & endline
   
	Response.Write "<option value=""0"">&nbsp;</option>"
	Do Until rsOptions.EOF
		if rsOptions.Fields(strDisplay)= ladder then
			insstring=" selected "
		else
			insstring= ""
		end if
		Response.Write "<OPTION value=" & chr(34) & Replace(rsOptions.Fields(strValue), "'", "\'") & chr(34) & insstring & ">"
		Response.Write Trim(rsOptions.Fields(strDisplay)) & "</OPTION>" & vbCRLF
		rsOptions.MoveNext
	LOOP

Response.Write "</SELECT>"
END SUB

SUB FillArray(strArrName, rsSource, strKey, strValue, strDisplay)

'Populate arrOptions as a two dimensional array.
'strArrName, 	the name we give the array;
'rsSource, 	the recordset used to populate the array;
'strKey, 	the name of the recordset field that will provide the "key" values
'		in the list’s options;
'strValue, 	the name of the recordset field that will provide the option
'		values submitted as part of the HTML form;
'strDisplay, 	the name of the recordset field that will provide the display 'values in the list’s options.
'
'[n][0] =matches the selected key from the primary list; 
'[n][1] =option value;
'[n][2] =display 

Response.Write "<SC" & "RIPT LANGUAGE= ""JavaScript"">" & _
	"var " & strArrName & " = new Array(); " & vbCRLF
DIM intRow
intRow = 0
DO UNTIL rsSource.EOF
   Response.Write strArrName & "[" & CStr(intRow) & _
   "] = new Array ('" & _
   SingleQuote(Trim(rsSource (strKey))) & _
   "', '" & Replace(rsSource(strValue), "'", "\'") & _
   "', '" & SingleQuote(Trim(rsSource(strDisplay))) & _
   "');" & vbCRLF
   intRow = intRow + 1
   rsSource.MoveNext
   Loop
   Response.Write " </SC" & "RIPT>"
END SUB

FUNCTION SingleQuote(strTarget)
	dim intpos
   '"Escapes" embedded single quote in a text string.
   intPos = InStr(1, strTarget, "'")
   DO WHILE intPos > 0 
      strTarget = Left(strTarget, intPos - 1) & "\'" & _
      Right(strTarget, Len(strTarget) - intPos)
      'Bump up TWO places because the original single quote has moved.
      intPos = InStr(intPos + 2, strTarget, "'")
   LOOP

   SingleQuote = strTarget
END FUNCTION
</script><script LANGUAGE="Javascript">
function ChangeOptions(lstPrimary, lstSecondary, strArray) 
{
var arrLen = eval(strArray + ".length");
var listLen = 0;
var strKey = eval("document.forms[0]." + lstPrimary + ".options[document.forms[0]." + lstPrimary + ".selectedIndex].value");
<!-- var strKey = eval("document.forms[1]." + lstPrimary + ".options[document.forms[1]." + lstPrimary + ".selectedIndex].value"); -->

eval("document.forms[0]." + lstSecondary + ".options.length = 0");

for (var i = 0; i < arrLen; i++) 
{
   if (eval(strArray + "[i][0] == " + strKey))
      {
<!--      eval("document.forms[1]." + lstSecondary + ".options[listLen] = new Option(" + strArray + "[i][2], " + strArray + "[i][1])"); -->
      eval("document.forms[0]." + lstSecondary + ".options[listLen] = new Option(" + strArray + "[i][2], " + strArray + "[i][1])");
      listLen = listLen + 1;
      }
   }

if (listLen > 0)
<!-- {eval("document.forms[1]." + lstSecondary + ".options[0].selected = true");} -->
     {eval("document.forms[0]." + lstSecondary + ".options[0].selected = true");}
      
   //alert ("document.all is " + document.all);
   //alert ("document.layers is " + document.layers);
   if (document.all == null) //Not using Internet Explorer
      {history.go(0);} 
   }

</script>
