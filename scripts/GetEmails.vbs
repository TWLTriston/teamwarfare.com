Dim oConn, ors, strSQL
SET oConn = CreateObject("ADODB.Connection")
oConn.Open "file name=c:\twl.udl"
Set ors = CreateObject("ADODB.Recordset")


Dim ladderRS 

Set ladderRS = CreateObject("ADODB.RecordSet")

Dim fso, f1, filename, filepath
Set fso = CreateObject("Scripting.FileSystemObject")
filepath = "E:\"
filename="Emails.tsv"
Set f1 = fso.CreateTextFile(filename, True)

Dim strDelimiter
strDelimiter =  vbTab
strSQL = "SELECT PlayerHandle, PlayerEmail FROM tbl_players WHERE PlayerActive = 'Y' ORDER BY PlayerHandle ASC "
oRs.Open strSQL, oConn
If Not (oRs.EOF AND oRs.BOF) Then
	Do While Not(oRs.EOF)
		f1.writeline Trim(oRs.Fields("PlayerHandle").Value) & strDelimiter & Trim(oRs.Fields("PlayerEmail").Value)
		oRs.MoveNext
	Loop
End If
oRs.Close

Set oRs = Nothing

Set f1 = Nothing
Set fso = Nothing
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
