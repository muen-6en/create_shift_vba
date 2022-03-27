Attribute VB_Name = "Module1"
Sub CreateMonthSheet()
    Application.DisplayAlerts = False
    
    'Get WorkBook
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    'Get Value.
    Dim month As String
    Dim macro As Worksheet
    Set macro = Worksheets("マクロ")
    Dim targetMonth As String
    Dim targetTerm As String
    targetMonth = macro.Range("F2").Value
    targetTerm = macro.Range("F3").Value

    If targetMonth = "" Then
        MsgBox "月を選択してください", vbOKOnly + vbCritical
        Exit Sub
    End If

    If targetTerm = "" Then
        MsgBox "期間を選択してください", vbOKOnly + vbCritical
        Exit Sub
    End If

    month = targetMonth & "月 " & targetTerm
    
    'Create new Sheet.
    Worksheets().Add After:=Worksheets(Worksheets.Count)
    Dim MonthSheet As Worksheet
    Set MonthSheet = ActiveSheet
    
    Dim SheetName() As String
    SheetName() = GetSheetName(SheetName, wb)

    Dim isName As Boolean
    isName = ChangeSheetMonth(isName, SheetName, month, MonthSheet)
    If isName Then
        Exit Sub
    End If
    
    ' Export Template.
    Call CreateTemplateMonth(MonthSheet, month)
    
End Sub

Function GetSheetName(SheetName() As String, ByVal wb As Workbook) As String()
    Dim SheetCnt As Long
    SheetCnt = wb.Sheets.Count
    ReDim SheetName(1 To SheetCnt)
    For i = 1 To SheetCnt
        SheetName(i) = Sheets(i).Name
    Next i

    GetSheetName = SheetName
End Function

Function ChangeSheetMonth(isName As Boolean, SheetName() As String, ByVal month As String, MonthSheet As Worksheet) As Boolean
    isName = False
    For Each Name In SheetName()
        If Name = month Then
            isName = True
        End If
    Next Name
    If isName Then
        MonthSheet.Delete
    Else
        MonthSheet.Name = month
    End If
    ChangeSheetMonth = isName
End Function

Function CreateTemplateMonth(MonthSheet As Worksheet, ByVal month As String)
    ' Initialize
    MonthSheet.Cells.Clear
    MonthSheet.Range("A1").Value = month
    MonthSheet.Range("A1").Font.Size = 14

    ' Class
    MonthSheet.Range("C2").Value = "勤務区分"
    MonthSheet.Range("C3").Value = "A"
    MonthSheet.Range("C4").Value = "B"
    MonthSheet.Range("C5").Value = "C"
    MonthSheet.Range("C6").Value = "D"

    ' Start
    MonthSheet.Range("D2").Value = "始業"
    MonthSheet.Range("D3").Value = "7:00"
    MonthSheet.Range("D4").Value = "9:00"
    MonthSheet.Range("D5").Value = "12:00"
    MonthSheet.Range("D6").Value = "14:00"

    ' End
    MonthSheet.Range("E2").Value = "終業"
    MonthSheet.Range("E3").Value = "16:00"
    MonthSheet.Range("E4").Value = "18:00"
    MonthSheet.Range("E5").Value = "21:00"
    MonthSheet.Range("E6").Value = "23:00"

    ' other
    MonthSheet.Range("F2").Value = "その他"
    MonthSheet.Range("F3").Value = "休：休日"
    MonthSheet.Range("F4").Value = "半：半休"

    ' Day
    ' Dim days(15) As Integet
    ' day() = GetDays(days, targetTerm)
    ' GetDays()
End Function

Function GetDays(days() As Integer, ByVal targetTerm As String) As Integer()
    Dim d As Integer

    If targetTerm = "前半" Then
        d = 1
    Else If targetTerm = "後半" Then
        d = 16
    End If

    For i = d To d + 14
        days(i) = d
    Next i

    GetDays = days
End Function
