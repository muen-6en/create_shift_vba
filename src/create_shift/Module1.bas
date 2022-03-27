Attribute VB_Name = "Module1"
Sub CreateMonthSheet()
    Application.DisplayAlerts = False
    
    'ワークブックの取得
    Dim wb As Workbook
    Set wb = Workbooks("シフト作成ツール.xlsm")
    Debug.Print ("ブック名：" & wb.Name)
    
    '入力値の取得
    Dim month As String
    Dim macro As Worksheet
    Set macro = Worksheets("マクロ")
    month = macro.Range("F2").Value & "月 " & macro.Range("F3").Value
    Debug.Print (month)
    
    '新規シートを作成
    Worksheets().Add After:=Worksheets(Worksheets.Count)
    Dim MonthSheet As Worksheet
    Set MonthSheet = ActiveSheet
    
    'シート名の取得
    Dim SheetCnt As Long
    Dim SheetName() As String
    SheetCnt = wb.Sheets.Count
    ReDim SheetName(1 To SheetCnt)
    For i = 1 To SheetCnt
        SheetName(i) = Sheets(i).Name
    Next i

    'シート名を変更
    Dim isName As Boolean
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

    
    
    '固定出力
    MonthSheet.Cells.Clear
    'MonthSheet.Range("").Value =
    MonthSheet.Range("A1").Value = month
    MonthSheet.Range("A1").Font.Size = 14
    MonthSheet.Range("C2").Value = "勤務区分"
    MonthSheet.Range("D2").Value = "始業"
    MonthSheet.Range("E2").Value = "終業"
    MonthSheet.Range("F2").Value = "その他"
    
End Sub
