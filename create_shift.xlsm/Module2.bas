Attribute VB_Name = "Module2"
' シフトを作成する
Sub CreateShift()
    Application.DisplayAlerts = False
    
    'Get WorkBook
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    'Get Value.
    Dim month As String
    Dim settingSheet As Worksheet
    Set settingSheet = Worksheets("マクロ")
    Dim targetSheet As String
    targetSheet = settingSheet.Cells(16, 6).Value

    Dim SheetName() As String
    Dim isName
    SheetName() = GetSheetName(SheetName, wb)
    For Each Name in SheetName()
        IF Name = targetSheet Then
            isName = True
        End IF
    Next Name

    IF isName = False Then
        MsgBox "シートがありません", vbOKOnly + vbCritical
        Exit Sub
    End IF

    ' Dim exportSheet As Worksheet
    ' Set exportSheet = Worksheet(targetSheet)

End Sub

' シート名を取得する
Function GetSheetName(SheetName() As String, ByVal wb As Workbook) As String()
    Dim SheetCnt As Long
    SheetCnt = wb.Sheets.Count
    ReDim SheetName(1 To SheetCnt)
    For i = 1 To SheetCnt
        SheetName(i) = Sheets(i).Name
    Next i

    GetSheetName = SheetName
End Function
