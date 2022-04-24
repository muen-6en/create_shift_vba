Attribute VB_Name = "Module1"

' 入力シート
Dim set_position_col As Integer
Dim set_member_col As Integer
Dim set_work_col As Integer

' 出力シート
Dim exp_position_col As Integer
Dim exp_member_col As Integer
Dim exp_work_col As Integer

' シートを作成する
Sub CreateMonthSheet()
    Application.DisplayAlerts = False
    
    'Get WorkBook
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    'Get Value.
    Dim month As String
    Dim settingSheet As Worksheet
    Set settingSheet = Worksheets("マクロ")
    Dim targetMonth As String
    Dim targetTerm As String
    targetMonth = settingSheet.Range("H7").Value
    targetTerm = settingSheet.Range("I7").Value

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
    Dim exportSheet As Worksheet
    Set exportSheet = ActiveSheet
    
    Dim SheetName() As String
    SheetName() = GetSheetName(SheetName, wb)

    Dim isName As Boolean
    isName = ChangeSheetMonth(isName, SheetName, month, exportSheet)
    If isName Then
        Exit Sub
    End If
    
    ' Export Sheet.
    Call ExportMonthSheet(exportSheet, settingSheet, month)
    
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

' シート名を変更する
Function ChangeSheetMonth(isName As Boolean, SheetName() As String, ByVal month As String, exportSheet As Worksheet) As Boolean
    isName = False
    For Each Name In SheetName()
        If Name = month Then
            isName = True
        End If
    Next Name
    If isName Then
        exportSheet.Delete
    Else
        exportSheet.Name = month
    End If
    ChangeSheetMonth = isName
End Function

' シートに月のシフトを出力する
Function ExportMonthSheet(exportSheet As Worksheet, settingSheet As Worksheet, ByVal month As String)
    set_position_col = 10
    set_member_col = 11
    set_work_col = 12
    exp_position_col = 1
    exp_member_col = 2
    exp_work_col = 3

    ' Initialize
    exportSheet.Cells.Clear
    exportSheet.Range("A1").Value = month
    exportSheet.Range("A1").Font.Size = 14

    ' Class
    exportSheet.Range("C2").Value = "勤務区分"
    exportSheet.Range("C3").Value = "A"
    exportSheet.Range("C4").Value = "B"
    exportSheet.Range("C5").Value = "C"
    exportSheet.Range("C6").Value = "D"
    exportSheet.Range("C7").Value = "休"

    ' Start
    exportSheet.Range("D2").Value = "始業"
    exportSheet.Range("D3").Value = "7:00"
    exportSheet.Range("D4").Value = "9:00"
    exportSheet.Range("D5").Value = "12:00"
    exportSheet.Range("D6").Value = "14:00"
    exportSheet.Range("D7").Value = "休日"

    ' End
    exportSheet.Range("E2").Value = "終業"
    exportSheet.Range("E3").Value = "16:00"
    exportSheet.Range("E4").Value = "18:00"
    exportSheet.Range("E5").Value = "21:00"
    exportSheet.Range("E6").Value = "23:00"

    ' Day
    exportSheet.Cells(10, 4).Value = "日付⇒"
    Call SetDate(exportSheet, settingSheet)

    Dim blankRow As Integer
    blankRow = 6
    Dim cnt As Integer
    cnt = settingSheet.Cells(blankRow, 10).End(xlDown).Row - blankRow

    ' Positon, Member, Workの関数は共通化する（暫定実装）
    ' Position
    exportSheet.Cells(11, 1).Value = "役職"
    exportSheet.Cells(11, 2).Value = "名前"
    exportSheet.Cells(11, 3).Value = "担当"

    For i = 1 To cnt
        exportSheet.Cells(11 + i, exp_position_col).Value = settingSheet.Cells(blankRow + i, set_position_col).Value
        exportSheet.Cells(11 + i, exp_member_col).Value = settingSheet.Cells(blankRow + i, set_member_col).Value
        exportSheet.Cells(11 + i, exp_work_col).Value = settingSheet.Cells(blankRow + i, set_work_col).Value
    Next i

    ' UI
    Dim end_row As Integer
    Dim end_col As Integer
    end_row = exportSheet.Cells(11, 1).End(xlDown).Row
    end_col = exportSheet.Cells(10, 4).End(xlToRight).Column
    exportSheet.Range(Cells(10, 1), Cells(end_row, end_col)).Borders.LineStyle = xlContinuous

End Function

' 日付出力
' exportSheet:出力先、 settingSheet:取得先、
Function SetDate(exportSheet As Worksheet, settingSheet As Worksheet)
    Dim date_row As Integer
    Dim week_row As Integer
    date_row = 10
    week_row = 11

    Dim targetYear As String
    Dim targetMonth As String
    Dim targetTerm As String
    Dim target As String
    Dim lstDate
    Dim day_num
    targetYear = settingSheet.Range("G7").Value
    targetMonth = settingSheet.Range("H7").Value
    targetTerm = settingSheet.Range("I7").Value
    target = STR(targetYear) + "/" + STR(targetMonth) + "/01"
    lstDate = Format(DateSerial(Year(target), month(target) + 1, 0), "d")

    Dim endRoop
    If targetTerm = "前半" Then
        day_num = 1
        endRoop = 14
    Else
        day_num = 16
        endRoop = Val(lstDate) - day_num
    End If

    Dim targetWeek As String
    Dim wk As String
    Dim exp_date_col As Integer
    exp_date_col = 5
    For i = exp_date_col To endRoop + exp_date_col
        targetWeek = Weekday(STR(targetYear) + "/" + STR(targetMonth) + "/" + STR(i + day_num - exp_date_col))
        Select Case targetWeek
            Case 1
                wk = "（日）"
            Case 2
                wk = "（月）"
            Case 3
                wk = "（火）"
            Case 4
                wk = "（水）"
            Case 5
                wk = "（木）"
            Case 6
                wk = "（金）"
            Case 7
                wk = "（土）"
        End Select
        exportSheet.Cells(date_row, i).Value = STR(i + day_num - exp_date_col) + "日"
        exportSheet.Cells(week_row, i).Value = wk
    Next i


End Function
