Attribute VB_Name = "Module1"

' ���̓V�[�g
Dim set_position_col As Integer
Dim set_member_col As Integer
Dim set_work_col As Integer

' �o�̓V�[�g
Dim exp_position_col As Integer
Dim exp_member_col As Integer
Dim exp_work_col As Integer

' �V�[�g���쐬����
Sub CreateMonthSheet()
    Application.DisplayAlerts = False
    
    'Get WorkBook
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    'Get Value.
    Dim month As String
    Dim settingSheet As Worksheet
    Set settingSheet = Worksheets("�}�N��")
    Dim targetMonth As String
    Dim targetTerm As String
    targetMonth = settingSheet.Range("H7").Value
    targetTerm = settingSheet.Range("I7").Value

    If targetMonth = "" Then
        MsgBox "����I�����Ă�������", vbOKOnly + vbCritical
        Exit Sub
    End If

    If targetTerm = "" Then
        MsgBox "���Ԃ�I�����Ă�������", vbOKOnly + vbCritical
        Exit Sub
    End If

    month = targetMonth & "�� " & targetTerm
    
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

' �V�[�g�����擾����
Function GetSheetName(SheetName() As String, ByVal wb As Workbook) As String()
    Dim SheetCnt As Long
    SheetCnt = wb.Sheets.Count
    ReDim SheetName(1 To SheetCnt)
    For i = 1 To SheetCnt
        SheetName(i) = Sheets(i).Name
    Next i

    GetSheetName = SheetName
End Function

' �V�[�g����ύX����
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

' �V�[�g�Ɍ��̃V�t�g���o�͂���
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
    exportSheet.Range("C2").Value = "�Ζ��敪"
    exportSheet.Range("C3").Value = "A"
    exportSheet.Range("C4").Value = "B"
    exportSheet.Range("C5").Value = "C"
    exportSheet.Range("C6").Value = "D"
    exportSheet.Range("C7").Value = "�x"

    ' Start
    exportSheet.Range("D2").Value = "�n��"
    exportSheet.Range("D3").Value = "7:00"
    exportSheet.Range("D4").Value = "9:00"
    exportSheet.Range("D5").Value = "12:00"
    exportSheet.Range("D6").Value = "14:00"
    exportSheet.Range("D7").Value = "�x��"

    ' End
    exportSheet.Range("E2").Value = "�I��"
    exportSheet.Range("E3").Value = "16:00"
    exportSheet.Range("E4").Value = "18:00"
    exportSheet.Range("E5").Value = "21:00"
    exportSheet.Range("E6").Value = "23:00"

    ' Day
    exportSheet.Cells(10, 4).Value = "���t��"
    Call SetDate(exportSheet, settingSheet)

    Dim blankRow As Integer
    blankRow = 6
    Dim cnt As Integer
    cnt = settingSheet.Cells(blankRow, 10).End(xlDown).Row - blankRow

    ' Positon, Member, Work�̊֐��͋��ʉ�����i�b������j
    ' Position
    exportSheet.Cells(11, 1).Value = "��E"
    exportSheet.Cells(11, 2).Value = "���O"
    exportSheet.Cells(11, 3).Value = "�S��"

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

' ���t�o��
' exportSheet:�o�͐�A settingSheet:�擾��A
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
    If targetTerm = "�O��" Then
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
                wk = "�i���j"
            Case 2
                wk = "�i���j"
            Case 3
                wk = "�i�΁j"
            Case 4
                wk = "�i���j"
            Case 5
                wk = "�i�؁j"
            Case 6
                wk = "�i���j"
            Case 7
                wk = "�i�y�j"
        End Select
        exportSheet.Cells(date_row, i).Value = STR(i + day_num - exp_date_col) + "��"
        exportSheet.Cells(week_row, i).Value = wk
    Next i


End Function
