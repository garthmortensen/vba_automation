Option Explicit

'▄▄▌ ▐ ▄▌ ▄ .▄ ▄▄▄· ▄▄▄▄▄    ▪  .▄▄ ·     ▄▄▄▄▄ ▄ .▄▪  .▄▄ · 
'██· █▌▐███▪▐█▐█ ▀█ •██      ██ ▐█ ▀.     •██  ██▪▐███ ▐█ ▀. 
'██▪▐█▐▐▌██▀▐█▄█▀▀█  ▐█.▪    ▐█·▄▀▀▀█▄     ▐█.▪██▀▐█▐█·▄▀▀▀█▄
'▐█▌██▐█▌██▌▐▀▐█ ▪▐▌ ▐█▌·    ▐█▌▐█▄▪▐█     ▐█▌·██▌▐▀▐█▌▐█▄▪▐█
' ▀▀▀▀ ▀▪▀▀▀ · ▀  ▀  ▀▀▀     ▀▀▀ ▀▀▀▀      ▀▀▀ ▀▀▀ ·▀▀▀ ▀▀▀▀ 
'I believe this script was for consolidating revenue?
'Or it is an attempt at performing a sql inner join in vba?


Sub AlignAndMatch()

    'backup sheet
    ActiveSheet.Copy after:=Sheets(Sheets.Count)
    
    'Insert rows where current cell <> cell above
    Dim i, totalrows As Integer
    Dim strRange As String
    Dim strRange2 As String

    '----------------------------------------
    'Monday sort table
    Range("A2:C65536").Select
    Selection.Sort Key1:=Range("A2:C65536"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal
    
    'Monday insert loop
    totalrows = ActiveSheet.Range("A65536").End(xlUp).Offset(1, 0).Row
    i = 0
    
    Do While i <= totalrows
       i = i + 1
       strRange = "A" & i
       strRange2 = "A" & i + 1
       If Range(strRange).Text <> Range(strRange2).Text Then
           Range(Cells(i + 1, 1), Cells(i + 2, 3)).Insert xlDown 'think cells ~A1:C2 insert
           totalrows = ActiveSheet.Range("A65536").End(xlUp).Offset(1, 0).Row
           i = i + 2 'for insert 2 rows
       End If
    Loop
    
    'Monday footer row loop
    totalrows = ActiveSheet.Range("A65536").End(xlUp).Offset(0, 0).Row
    i = 0
    
    Do While i <= totalrows
       i = i + 1
       If IsEmpty(Range("A" & i).Value) And Not IsEmpty(Range("A" & i + 1).Value) Then
           Range("A" & i).Value = Range("A" & i + 1).Value
           Range("B" & i).Value = "Sum"
       End If
    Loop
    
    '----------------------------------------
    'Tuesday sort table
    Range("E2:G65536").Select
    Selection.Sort Key1:=Range("E2:G65536"), Order1:=xlAscending, Header:=xlGuess, _
    OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
    DataOption1:=xlSortNormal

    'Tuesday insert loop
    totalrows = ActiveSheet.Range("E65536").End(xlUp).Offset(0, 0).Row
    i = 0
    
    Do While i <= totalrows
       i = i + 1
       strRange = "E" & i
       strRange2 = "E" & i + 1
       If Range(strRange).Text <> Range(strRange2).Text Then
           Range(Cells(i + 1, 5), Cells(i + 2, 7)).Insert xlDown 'think cells ~A1:C2 insert
           totalrows = ActiveSheet.Range("A65536").End(xlUp).Offset(1, 0).Row
           i = i + 2 'for insert 2 rows
       End If
    Loop
    
    'Tuesday footer row loop
    totalrows = ActiveSheet.Range("E65536").End(xlUp).Offset(0, 0).Row
    i = 0
    
    Do While i <= totalrows
       i = i + 1
       If IsEmpty(Range("E" & i).Value) And Not IsEmpty(Range("E" & i + 1).Value) Then
           Range("E" & i).Value = Range("E" & i + 1).Value
           Range("F" & i).Value = "Sum"
       End If
    Loop
    
    Dim colStart As Integer
    
    'Sum calculations
    st = Cells(1, colStart).End(xlDown).Row
    Dn = Cells(200, colStart).End(xlUp).Row
    Set myRange = ActiveSheet.Range _
    (Cells(st, colStart), Cells(Dn, colStart))


End Sub


Sub HighlightMatches()

    Dim i, LastRowA, LastRowB
    LastRowA = Range("A" & Rows.Count).End(xlUp).Row
    LastRowB = Range("B" & Rows.Count).End(xlUp).Row
    Columns("A:A").Interior.ColorIndex = xlNone
    Columns("B:B").Interior.ColorIndex = xlNone
    For i = 1 To LastRowA
        If Application.CountIf(Range("B:B"), Cells(i, "A")) > 0 Then
            Cells(i, "A").Interior.ColorIndex = 36
        End If
    Next
    For i = 1 To LastRowB
        If Application.CountIf(Range("A:A"), Cells(i, "B")) > 0 Then
            Cells(i, "B").Interior.ColorIndex = 36
        End If
    Next

End Sub


Sub AlignCustNbr()
' hiker95, 01/10/2011
' http://www.mrexcel.com/forum/showthread.php?t=520077
'
' The macro was modified from code by:
' Krishnakumar, 12/12/2010
' http://www.ozgrid.com/forum/showthread.php?t=148881
'
Dim ws As Worksheet
Dim LR As Long, a As Long
Dim CustNbr As Range
Application.ScreenUpdating = False
Set ws = Worksheets("Sheet1")
LR = ws.Range("E" & ws.Rows.Count).End(xlUp).Row
ws.Range("E3:G" & LR).Sort Key1:=ws.Range("E3"), Order1:=xlAscending, Header:=xlNo, _
  OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
LR = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
ws.Range("A3:C" & LR).Sort Key1:=ws.Range("A3"), Order1:=xlAscending, Header:=xlNo, _
  OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal
Set CustNbr = ws.Range("A2:C" & LR)
a = 2
Do While CustNbr.Cells(a, 1) <> ""
  If CustNbr.Cells(a, 1).Offset(, 4) <> "" Then
    If CustNbr.Cells(a, 1) < CustNbr.Cells(a, 1).Offset(, 4) Then
      CustNbr.Cells(a, 1).Offset(, 4).Resize(, 3).Insert -4121
    ElseIf CustNbr.Cells(a, 1) > CustNbr.Cells(a, 1).Offset(, 4) Then
      CustNbr.Cells(a, 1).Resize(, 3).Insert -4121
      LR = LR + 1
      Set CustNbr = ws.Range("A3:C" & LR)
    End If
  End If
  a = a + 1
Loop
Application.ScreenUpdating = 1
End Sub
