Dim oConn, ors, strSQL
SET oConn = CreateObject("ADODB.Connection")
oConn.Open "file name=c:\twl.udl"
Set ors = CreateObject("ADODB.Recordset")


Dim ladderRS 

Set ladderRS = CreateObject("ADODB.RecordSet")

Dim fso, f1, filename, filepath
Set fso = CreateObject("Scripting.FileSystemObject")
filepath = "E:\INETPUB\predictit\ladderdata\"

Dim strDelimiter
strDelimiter =  vbTab & vbTab
strSQL = "SELECT LadderName FROM tbl_Ladders WHERE LadderActive = 1 "
ladderRS.Open strSQL, oConn
If Not (ladderRS.EOF AND ladderRS.BOF) Then
	Do While Not(ladderRS.EOF)
		filename = "twl_" & replace(replace(ladderRS.Fields("LadderName").Value, " ", "_"),":","") & "_results.txt"
		Set f1 = fso.CreateTextFile(filepath & filename, True)
	
		strSQL = "Select TOP 20 * from vHistoryPredictIt WHERE LadderName = '" & Replace(ladderRS.Fields("LadderName").Value, "'", "''") & "' ORDER BY MATCHDATE DESC"
		ors.Open strsql, oconn
		if not(ors.EOF and ors.BOF) THEN
			do while not(ors.EOF)
				WinnerRank = ors.Fields("WinnerRank").Value
				WinnerName = ors.Fields("WinnerName").Value
				LoserRank = ors.Fields("LoserRank").Value
				LoserName = ors.Fields("LoserName").Value
				MatchForfeit = ors.Fields("MatchForfeit").Value
				MatchUpDate = ors.Fields("MatchUpDate").Value
				Map1 = ors.Fields("Map1").Value
				WinnerMap1Score = ors.Fields("WinnerMap1Score").Value
				LoserMap1Score = ors.Fields("LoserMap1Score").Value
				Map2 = ors.Fields("Map2").Value
				WinnerMap2Score = ors.Fields("WinnerMap2Score").Value
				LoserMap2Score = ors.Fields("LoserMap2Score").Value
				Map3 = ors.Fields("Map3").Value
				WinnerMap3Score = ors.Fields("WinnerMap3Score").Value
				LoserMap3Score = ors.Fields("LoserMap3Score").Value
				LadderName = ors.Fields("LadderName").Value
				
				f1.write CheckItem(WinnerRank) & strDelimiter
				f1.write CheckItem(WinnerName) & strDelimiter
				f1.write CheckItem(LoserRank) & strDelimiter
				f1.write CheckItem(LoserName) & strDelimiter
				f1.write CheckItem(MatchForfeit) & strDelimiter
				f1.write CheckItem(MatchUpDate) & strDelimiter
				f1.write CheckItem(Map1) & strDelimiter
				f1.write CheckItem(WinnerMap1Score) & strDelimiter
				f1.write CheckItem(LoserMap1Score) & strDelimiter
				f1.write CheckItem(Map2) & strDelimiter
				f1.write CheckItem(WinnerMap2Score) & strDelimiter
				f1.write CheckItem(LoserMap2Score) & strDelimiter
				f1.write CheckItem(Map3) & strDelimiter
				f1.write CheckItem(WinnerMap3Score) & strDelimiter
				f1.write CheckItem(LoserMap3Score) & strDelimiter
				f1.write CheckItem(LadderName) & strDelimiter
				f1.writeline("")
				ors.movenext
			loop
		end if
		f1.close
		ors.nextrecordset
		ladderRS.MoveNExt
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
