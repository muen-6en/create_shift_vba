Attribute VB_Name = "Module1"
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
    targetMonth = settingSheet.Range("F3").Value
    targetTerm = settingSheet.Range("F4").Value

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
    
    ' Export Template.
    Call CreateTemplateMonth(exportSheet, settingSheet, month)
    
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

Function CreateTemplateMonth(exportSheet As Worksheet, settingSheet As Worksheet, ByVal month As String)
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

    ' Start
    exportSheet.Range("D2").Value = "始業"
    exportSheet.Range("D3").Value = "7:00"
    exportSheet.Range("D4").Value = "9:00"
    exportSheet.Range("D5").Value = "12:00"
    exportSheet.Range("D6").Value = "14:00"

    ' End
    exportSheet.Range("E2").Value = "終業"
    exportSheet.Range("E3").Value = "16:00"
    exportSheet.Range("E4").Value = "18:00"
    exportSheet.Range("E5").Value = "21:00"
    exportSheet.Range("E6").Value = "23:00"

    ' other
    exportSheet.Range("F2").Value = "その他"
    exportSheet.Range("F3").Value = "休：休日"
    exportSheet.Range("F4").Value = "半：半休"

    ' Day
    exportSheet.Cells(8, 3).Value = "日付⇒"
    Call SetDate(exportSheet, settingSheet)

    ' Positon, Member, Workの関数は共通化する（暫定実装）
    ' Position
    exportSheet.Cells(9, 1).Value = "役職"
    Call SetPosition(exportSheet, settingSheet)

    ' Member
    exportSheet.Cells(9, 2).Value = "名前"
    Call SetMember(exportSheet, settingSheet)

    ' Work
    exportSheet.Cells(9, 3).Value = "担当"
    Call SetWork(exportSheet, settingSheet)

    ' UI
    ' TODO: A8-最終入力セルを取得して罫線を設定する
    exportSheet.Range("A8:S14").Borders.LineStyle = xlContinuous

End Function

' 日付出力
' exportSheet:出力先、 settingSheet:取得先、
Function SetDate(exportSheet As Worksheet, settingSheet As Worksheet)

    Dim targetYear As String
    Dim targetMonth As String
    Dim targetTerm As String
    Dim target As String
    Dim lstDate
    Dim day_num
    targetYear = settingSheet.Range("F2").Value
    targetMonth = settingSheet.Range("F3").Value
    targetTerm = settingSheet.Range("F4").Value
    target = STR(targetYear) + "/" + STR(targetMonth) + "/01"
    lstDate = Format(DateSerial(Year(target), month(target) + 1, 0), "d")

    Dim endRoop
    If targetTerm = "前半" Then
        day_num = 1
        endRoop = 15
    Else
        day_num = 16
        endRoop = Val(lstDate) - day_num
    End If

    Dim targetWeek As String
    Dim wk As String
    For i = 4 To endRoop + 4
        targetWeek = Weekday(STR(targetYear) + "/" + STR(targetMonth) + "/" + STR(i + day_num - 4))
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
        exportSheet.Cells(8, i).Value = STR(i + day_num - 4) + "日"
        exportSheet.Cells(9, i).Value = wk
    Next i


End Function

' 役職出力
' exportSheet:出力先、 settingSheet:取得先、
Function SetPosition(exportSheet As Worksheet, settingSheet As Worksheet)
    ' 最後の入力行を取得
    ' 9行目～最終行の値を取得
    ' 値を設定
    Dim lstRow
    lstRow = exportSheet.Cells(9, 1).End(xlDown).Row

    exportSheet.Cells(10, 1).Value = settingSheet.Cells(7, 5).Value
    exportSheet.Cells(11, 1).Value = settingSheet.Cells(8, 5).Value
    exportSheet.Cells(12, 1).Value = settingSheet.Cells(9, 5).Value
    exportSheet.Cells(13, 1).Value = settingSheet.Cells(10, 5).Value
    exportSheet.Cells(14, 1).Value = settingSheet.Cells(11, 5).Value

End Function

' 名前出力
' exportSheet:出力先、 settingSheet:取得先、
Function SetMember(exportSheet As Worksheet, settingSheet As Worksheet)
    ' 最後の入力行を取得
    ' 9行目～最終行の値を取得
    ' 値を設定
    Dim lstRow
    lstRow = exportSheet.Cells(9, 2).End(xlDown).Row

    exportSheet.Cells(10, 2).Value = settingSheet.Cells(7, 6).Value
    exportSheet.Cells(11, 2).Value = settingSheet.Cells(8, 6).Value
    exportSheet.Cells(12, 2).Value = settingSheet.Cells(9, 6).Value
    exportSheet.Cells(13, 2).Value = settingSheet.Cells(10, 6).Value
    exportSheet.Cells(14, 2).Value = settingSheet.Cells(11, 6).Value

End Function

' 担当出力
' exportSheet:出力先、 settingSheet:取得先、
Function SetWork(exportSheet As Worksheet, settingSheet As Worksheet)
    ' 最後の入力行を取得
    ' 9行目～最終行の値を取得
    ' 値を設定
    Dim lstRow
    lstRow = exportSheet.Cells(9, 2).End(xlDown).Row

    exportSheet.Cells(10, 3).Value = settingSheet.Cells(7, 7).Value
    exportSheet.Cells(11, 3).Value = settingSheet.Cells(8, 7).Value
    exportSheet.Cells(12, 3).Value = settingSheet.Cells(9, 7).Value
    exportSheet.Cells(13, 3).Value = settingSheet.Cells(10, 7).Value
    exportSheet.Cells(14, 3).Value = settingSheet.Cells(11, 7).Value

End Function

