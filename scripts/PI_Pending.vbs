Dim oConn, ors, strSQL
Set oConn = CreateObject("ADODB.Connection")
oConn.Open "file name=c:\twl.udl"
Set oRs = CreateObject("ADODB.Recordset")

Dim ladderRS 

Set ladderRS = CreateObject("ADODB.RecordSet")

Dim fso, f1, filename, filepath
Set fso = CreateObject("Scripting.FileSystemObject")
filepath = "E:\INETPUB\predictit\ladderdata\"

Dim strMatchDate
Dim strDelimiter
strDelimiter =  vbTab & vbTab

strSQL = "SELECT LadderName FROM tbl_Ladders WHERE LadderActive = 1 "
ladderRS.Open strSQL, oConn
If Not (ladderRS.EOF AND ladderRS.BOF) Then
	Do While Not(ladderRS.EOF)
		filename = "twl_" & replace(replace(ladderRS.Fields("LadderName").Value, " ", "_"), ":", "")  & ".txt"
		Set f1 = fso.CreateTextFile(filepath & filename, True)
		strSQL = "Select * from vPending WHERE LadderName = '" & Replace(ladderRS.Fields("LadderName").Value, "'", "''") & "' AND MatchDate IS NOT NULL ORDER BY DEFENDERRANK ASC"
		ors.Open strsql, oconn
		if not(ors.EOF and ors.BOF) THEN
			do while not(ors.EOF)
				f1.write CheckItem(ors.Fields("DefenderRank").Value) & strDelimiter
				f1.write CheckItem(ors.Fields("DefenderName").Value) & strDelimiter
				f1.write CheckItem(ors.Fields("DefenderWins").Value & "/" & ors.Fields("DefenderLosses").Value) & strDelimiter
				f1.write CheckItem(ors.Fields("AttackerRank").Value) & strDelimiter
				f1.write CheckItem(ors.Fields("AttackerName").Value) & strDelimiter
				f1.write CheckItem(ors.Fields("AttackerWins").Value & "/" & ors.Fields("AttackerLosses").Value) & strDelimiter
				strMatchDate = CheckItem(oRS.Fields("MatchDate").value)
				f1.write formatdatetime(ors.Fields("mDate").value & "", 2) & strDelimiter
				f1.write CheckItem(ors.Fields("Map1").Value) & strDelimiter
				f1.write CheckItem(ors.Fields("Map2").Value) & strDelimiter
				f1.write CheckItem(ors.Fields("Map3").Value) & strDelimiter
				f1.write CheckItem(ors.Fields("LadderName").Value)
				f1.writeline("")
				ors.movenext
			loop
		end if
		ors.nextrecordset
		f1.close
		ladderRS.MoveNext
	Loop
End If
ladderRS.Close

Set LadderRS = Nothing
set f1 = nothing
set fso = nothing
SET ors = nothing
oConn.Close
set oconn = nothing

Function CheckItem(byVal strData)
	CheckItem = URLEncode(strData)
	If IsNull(strData) Or Len(Trim(strData)) = 0 Or strData = "" Then
			CheckItem = "?"
	End If
End Function

Function URLEncode(strData)
	Dim I 
	Dim strTemp 
	Dim strChar 
	Dim strOut 
	Dim intAsc 
	strTemp = Trim(strData)
	If IsNull(strTemp) Then
		strTemp = "?"
	End If
	For I = 1 To Len(strTemp)
	   strChar = Mid(strTemp, I, 1)
	   intAsc = Asc(strChar)
	   If (intAsc >= 48 And intAsc <= 57) Or _
	      (intAsc >= 97 And intAsc <= 122) Or _
	      (intAsc >= 65 And intAsc <= 90) Then
	      strOut = strOut & strChar
	   Else
	      strOut = strOut & "%" & Hex(intAsc)
	   End If
	Next
	URLEncode = strOut

End Function


Function URLDecode(strConvert)
	Dim arySplit
	Dim strHex
	Dim strOutput
	
	If IsNull(strConvert) Then
	   URLDecode = ""
	   Exit Function
	End If
	
	' First convert the + to a space
	strOutput = REPLACE(strConvert, "+", " ")
	
	' Then convert the %number to normal code
	arySplit = Split(strOutput, "%")
	strOutput = arySplit(LBound(arySplit))
   For I = LBound(arySplit) to UBound(arySplit) - 1
      strHex = "&H" & Left(arySplit(i+1),2)
      Letter = Chr(strHex)

      strOutput = strOutput & Letter & Right(arySplit(i+1),len(arySplit(i+1))-2)
   Next
	
	URLDecode = strOutput
End Function
