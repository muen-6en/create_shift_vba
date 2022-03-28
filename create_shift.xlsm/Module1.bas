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
    Call CreateTemplateMonth(MonthSheet, macro, month)
    
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

Function CreateTemplateMonth(MonthSheet As Worksheet,macro As Worksheet, ByVal month As String)
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

    ' Position
    ' マクロシートの特定のセルから値を取得する
    Dim positions() As String
    positions() = GetBelongs(positions(), macro, "E")
    Call SetBelongs(positions(), MonthSheet, "A")
    MonthSheet.Range("A8").Value = "役職"
    MonthSheet.Range("A9").Value = "施設長"
    MonthSheet.Range("A10").Value = "社員"
    MonthSheet.Range("A11").Value = "社員"
    MonthSheet.Range("A12").Value = "契約社員"
    MonthSheet.Range("A13").Value = "パート"

    ' Member
    Dim members() As String
    members() = GetBelongs(members(), macro, "F")
    Call SetBelongs(members(), MonthSheet, "B")
    MonthSheet.Range("B8").Value = "名前"
    MonthSheet.Range("B9").Value = "部長薫子"
    MonthSheet.Range("B10").Value = "社員太郎"
    MonthSheet.Range("B11").Value = "社員心太"
    MonthSheet.Range("B12").Value = "契約花子"
    MonthSheet.Range("B13").Value = "仁科仁部"

    ' Work
    MonthSheet.Range("C8").Value = "担当"

    ' Day
    ' Dim days(15) As Integet
    ' day() = GetDays(days, targetTerm)
    ' GetDays()

    ' UI
    ' TODO: A8-最終入力セルを取得して罫線を設定する
    MonthSheet.Range("A8:B13").Borders.LineStyle = xlContinuous

End Function

Function GetBelongs(collections() As String , macro As Worksheet, col As String) As String()
    ' macroシートのRow6以下の値が空白までループ
    ' 空白までの数をカウント
    ' collectionsの要素数をカウントにReDim
    ' collectionsに役職の値を設定する

    GetPosition = collections
End Function

Function SetBelongs(collections() As String , MonthSheet As Worksheet, col As String)
    ' MonthSheetに値を設定する
    Exit Function
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
