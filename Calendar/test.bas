Sub WriteSmallCal(ByVal week As Integer, ByVal SheetName As String)
  Dim smallDays(1 To 6, 1 To 7) As Variant
  Dim dWeeks As Variant
  Dim year As Integer
  Dim baseLine As Integer
  Dim firstMonth As String
  Dim secondMonth As String
  
  year = 2023
  baseLine = GetBaseLine(week)
  
  REM WeekNums and days
  dWeeks = GetDisplayWeeks(week)
  For i = 1 To 6
    Sheets(SheetName).Cells(baseLine + 37 + i, 19) = dWeeks(i)
    datesOfWeek = GetDatesOfWeek(year, dWeeks(i))
    For k = 1 To 7
      If i = 1 And k = 1 Then
        firstMonth = UCase(MonthName(Month(datesOfWeek(1)), True))
      End If
      If i = 1 And k = 7 Then
        secondMonth = UCase(MonthName(Month(datesOfWeek(7)), True))
      End If
      Sheets(SheetName).Cells(baseLine + 37 + i, 19 + k) = Day(datesOfWeek(k))
    Next k
  Next i
  
End Sub
  
  
  
  
  
  
