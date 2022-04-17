Attribute VB_Name = "Module1"
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
    targetMonth = settingSheet.Range("F3").Value
    targetTerm = settingSheet.Range("F4").Value

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
    
    ' Export Template.
    Call CreateTemplateMonth(exportSheet, settingSheet, month)
    
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

Function CreateTemplateMonth(exportSheet As Worksheet, settingSheet As Worksheet, ByVal month As String)
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

    ' Start
    exportSheet.Range("D2").Value = "�n��"
    exportSheet.Range("D3").Value = "7:00"
    exportSheet.Range("D4").Value = "9:00"
    exportSheet.Range("D5").Value = "12:00"
    exportSheet.Range("D6").Value = "14:00"

    ' End
    exportSheet.Range("E2").Value = "�I��"
    exportSheet.Range("E3").Value = "16:00"
    exportSheet.Range("E4").Value = "18:00"
    exportSheet.Range("E5").Value = "21:00"
    exportSheet.Range("E6").Value = "23:00"

    ' other
    exportSheet.Range("F2").Value = "���̑�"
    exportSheet.Range("F3").Value = "�x�F�x��"
    exportSheet.Range("F4").Value = "���F���x"

    ' Day
    exportSheet.Cells(8, 3).Value = "���t��"
    Call SetDate(exportSheet, settingSheet)

    ' Positon, Member, Work�̊֐��͋��ʉ�����i�b������j
    ' Position
    exportSheet.Cells(9, 1).Value = "��E"
    Call SetPosition(exportSheet, settingSheet)

    ' Member
    exportSheet.Cells(9, 2).Value = "���O"
    Call SetMember(exportSheet, settingSheet)

    ' Work
    exportSheet.Cells(9, 3).Value = "�S��"
    Call SetWork(exportSheet, settingSheet)

    ' UI
    ' TODO: A8-�ŏI���̓Z�����擾���Čr����ݒ肷��
    exportSheet.Range("A8:S14").Borders.LineStyle = xlContinuous

End Function

' ���t�o��
' exportSheet:�o�͐�A settingSheet:�擾��A
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
    If targetTerm = "�O��" Then
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
        exportSheet.Cells(8, i).Value = STR(i + day_num - 4) + "��"
        exportSheet.Cells(9, i).Value = wk
    Next i


End Function

' ��E�o��
' exportSheet:�o�͐�A settingSheet:�擾��A
Function SetPosition(exportSheet As Worksheet, settingSheet As Worksheet)
    ' �Ō�̓��͍s���擾
    ' 9�s�ځ`�ŏI�s�̒l���擾
    ' �l��ݒ�
    Dim lstRow
    lstRow = exportSheet.Cells(9, 1).End(xlDown).Row

    exportSheet.Cells(10, 1).Value = settingSheet.Cells(7, 5).Value
    exportSheet.Cells(11, 1).Value = settingSheet.Cells(8, 5).Value
    exportSheet.Cells(12, 1).Value = settingSheet.Cells(9, 5).Value
    exportSheet.Cells(13, 1).Value = settingSheet.Cells(10, 5).Value
    exportSheet.Cells(14, 1).Value = settingSheet.Cells(11, 5).Value

End Function

' ���O�o��
' exportSheet:�o�͐�A settingSheet:�擾��A
Function SetMember(exportSheet As Worksheet, settingSheet As Worksheet)
    ' �Ō�̓��͍s���擾
    ' 9�s�ځ`�ŏI�s�̒l���擾
    ' �l��ݒ�
    Dim lstRow
    lstRow = exportSheet.Cells(9, 2).End(xlDown).Row

    exportSheet.Cells(10, 2).Value = settingSheet.Cells(7, 6).Value
    exportSheet.Cells(11, 2).Value = settingSheet.Cells(8, 6).Value
    exportSheet.Cells(12, 2).Value = settingSheet.Cells(9, 6).Value
    exportSheet.Cells(13, 2).Value = settingSheet.Cells(10, 6).Value
    exportSheet.Cells(14, 2).Value = settingSheet.Cells(11, 6).Value

End Function

' �S���o��
' exportSheet:�o�͐�A settingSheet:�擾��A
Function SetWork(exportSheet As Worksheet, settingSheet As Worksheet)
    ' �Ō�̓��͍s���擾
    ' 9�s�ځ`�ŏI�s�̒l���擾
    ' �l��ݒ�
    Dim lstRow
    lstRow = exportSheet.Cells(9, 2).End(xlDown).Row

    exportSheet.Cells(10, 3).Value = settingSheet.Cells(7, 7).Value
    exportSheet.Cells(11, 3).Value = settingSheet.Cells(8, 7).Value
    exportSheet.Cells(12, 3).Value = settingSheet.Cells(9, 7).Value
    exportSheet.Cells(13, 3).Value = settingSheet.Cells(10, 7).Value
    exportSheet.Cells(14, 3).Value = settingSheet.Cells(11, 7).Value

End Function

