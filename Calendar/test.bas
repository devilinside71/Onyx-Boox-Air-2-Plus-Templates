Sub WriteLittleCal()
  '
  ' Macro2 Macro
  '
  
  
  
  '
  
  
  Dim tRange As String
  
  Dim sRanges(100) As String
  Dim rCount As Integer
  Dim iWeek As Integer
  Dim iMonth As Integer
  Dim sWeek As String
  Dim ssWeek As String
  Dim sssWeek As String
  
  sRanges(1) = "V1:AC7"
  sRanges(2) = "AE1:AL7"
  sRanges(3) = "AN1:AU7"
  
  sRanges(4) = "V10:AC15"
  sRanges(5) = "AE10:AL15"
  sRanges(6) = "AN10:AU15"
  
  sRanges(7) = "V18:AC24"
  sRanges(8) = "AE18:AL24"
  sRanges(9) = "AN18:AU24"
  
  sRanges(10) = "V27:AC33"
  sRanges(11) = "AE27:AL33"
  sRanges(12) = "AN27:AU33"
  
  
  
  For i = 2 To 54
    
    iWeek = i -1
    sWeek = GetDoubleNum(iWeek)
    'Debug.Print sWeek
    iMonth = Sheets("Sheet3").Cells(i, 13)
    rCount =(iWeek -1) * 38 + 32
    tRange = "O" + Trim(CStr(rCount)) + ":V" + Trim(CStr(rCount + 6))
    Range(tRange).Select
    Application.CutCopyMode = False
    Selection.Delete Shift: = xlToLeft
    
    Sheets("Sheet3").Select
    Range(sRanges(iMonth)).Select
    
    Selection.Copy
    Sheets("Sheet5").Select
    Range(Split(tRange, ":")(0)).Select
    ActiveSheet.Paste
    Selection.PasteSpecial Paste: = xlPasteAllUsingSourceTheme, Operation: = xlNone _
          , SkipBlanks: = False, Transpose: = False
    
    For k = rCount To(rCount + 6)
      ssWeek = Trim(Sheets("Sheet5").Cells(k, 15))
      If ssWeek <> "" Then
        sssWeek = GetDoubleNum(CInt(ssWeek))
        
      Else
        sssWeek = "00"
      End If
      Debug.Print sWeek, ssWeek
      If sssWeek = sWeek Then
        
        Range("O" + Trim(CStr(k))).Select
        With Selection.Interior
          .Pattern = xlSolid
          .PatternColorIndex = xlAutomatic
          .ThemeColor = xlThemeColorLight1
          .TintAndShade = 0.349986266670736
          .PatternTintAndShade = 0
        End With
        With Selection.Font
          .ThemeColor = xlThemeColorDark1
          .TintAndShade = 0
        End With
        Selection.Font.Bold = True
        
      End If
    Next k
    
  Next i
End Sub
  
Function GetDoubleNum(number As Integer) As String
  Dim retVal As String
  
  GetDoubleNum = Trim(CStr(number))
  If number < 10 Then
    
    GetDoubleNum = "0" + GetDoubleNum
  End If
  
  
  
End Function
  
  
  
  
  
  
