Attribute VB_Name = "Module2"
Dim targetYear As String ' �N
Dim targetMonth As String ' ��
Dim days() As String ' ���t
Dim weeks() As String ' �T

' �V�t�g���쐬����
Sub CreateShift()
    Application.DisplayAlerts = False
    
    'Get WorkBook
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    'Get Value.
    Dim month As String
    Dim settingSheet As Worksheet
    Set settingSheet = Worksheets("�}�N��")
    Dim targetTerm As String
    targetYear = settingSheet.Cells(7, 7).Value 
    targetMonth = settingSheet.Cells(7, 8).Value
    targetTerm = settingSheet.Cells(7, 9).Value
    Dim targetSheetName As String
    targetSheetName = targetMonth & "�� " & targetTerm

    Dim SheetName() As String
    Dim isTargetSheet
    SheetName() = GetSheetName(SheetName, wb)
    For Each Name in SheetName()
        IF Name = targetSheetName Then
            isTargetSheet = True
        End IF
    Next Name

    IF isTargetSheet = False Then
        MsgBox "�V�[�g������܂���", vbOKOnly + vbCritical
        Exit Sub
    End IF

    ' �V�[�g���쐬����
    Worksheets().Add After:=Worksheets(Worksheets.Count)
    Dim exportSheet As Worksheet
    Set exportSheet = ActiveSheet

    Dim isExportSheet As Boolean
    Dim shiftName As String
    shiftName = targetSheetName + " �V�t�g"
    isExportSheet = ChangeSheetName(isExportSheet, SheetName, shiftName, exportSheet)
    IF isExportSheet Then
        MsgBox "�V�[�g�����݂��܂�", vbOKOnly + vbOKOnly
        Exit Sub
    End IF

    Dim inputSheet As Worksheet
    Set inputSheet = Worksheets(targetSheetName)

    ' �V�t�g�̓������擾����
    Dim dateCnt As Integer
    dateCnt = GetDateCount(inputSheet)

    ' �V�[�g��ǂݎ��
    Dim data() As String
    data() = GetInputData(inputSheet, dateCnt)

    Dim header_row As Integer
    Dim day_row As Integer
    header_row = 3
    day_row = UBound(data, 1) + 6
    ' �V�[�g���o�͂���
    For i = 1 To dateCnt
        Dim array_num As Integer
        array_num = i
        Call ExportTemplete(exportSheet, array_num, header_row)
        Call ExportData(exportSheet, data, header_row, 4 + i)
        header_row = header_row + day_row
    Next i

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
Function ChangeSheetName(isName As Boolean, SheetName() As String, ByVal setName As String, exportSheet As Worksheet) As Boolean
    ' ���ɃV�[�g�����݂���Ώ������I������
    isName = False
    For Each Name In SheetName()
        If Name = setName Then
            isName = True
        End If
    Next Name
    If isName Then
        exportSheet.Delete
    Else
        exportSheet.Name = setName
    End If
    ChangeSheetName = isName
End Function

' �o�͓������擾����
Function GetDateCount(inputSheet As Worksheet) As Integer
    Dim start_num As Integer
    Dim end_num As Integer
    Dim date_num As Integer

    end_num = inputSheet.Cells(10, Columns.Count).End(xlToLeft).Column
    start_num = inputSheet.Cells(10, end_num).End(xlToLeft).Column

    date_num = end_num - start_num

    Redim days(1 To date_num)
    Redim weeks(1 To date_num)
    For i = 1 To date_num
        days(i) = inputSheet.Cells(10, start_num + i).Value
        weeks(i) = inputSheet.Cells(11, start_num + i).Value
    Next i

    GetDateCount = date_num
End Function

' ���̓f�[�^�擾
Function GetInputData(inputSheet As Worksheet, dateCnt As Integer) As String()
    Dim position_col As Integer
    Dim member_col As Integer
    Dim work_col As Integer
    position_col = 1
    member_col = 2
    work_col = 3

    Dim start_row As Integer
    DIm end_row As Integer
    Dim row_num As Integer

    start_row = inputSheet.Cells(1, position_col).End(xlDown).Row
    end_row = inputSheet.Cells(start_row, position_col).End(xlDown).Row
    row_num = end_row - start_row

    ReDim data(row_num, 4 + dateCnt) As String
    For i = 1 To row_num
        For j = 1 To 4 + dateCnt
            data(i, j) = inputSheet.Cells(start_row + i, j).Value
        Next j
    Next i

    GetInputData = data()
End Function

' �V�t�g�̂ЂȌ^���o�͂���
Function ExportTemplete(exportSheet As Worksheet, array_num As Integer, header_row)
    exportSheet.Cells(header_row - 2, 1).Value = _
     targetYear + "�N" + targetMonth + "��" + _ 
     days(array_num) + weeks(array_num) + "�V�t�g"
    exportSheet.Cells(header_row - 2, 1).Font.Size = 14

    exportSheet.Cells(header_row, 1).Value = "��E"
    exportSheet.Cells(header_row, 2).Value = "���O"
    exportSheet.Cells(header_row, 3).Value = "�S��"
    exportSheet.Cells(header_row, 4).Value = "�Ζ��敪"

    exportSheet.Cells(header_row, 5).Value = "7:00"
    exportSheet.Cells(header_row, 6).Value = "7:30"
    exportSheet.Cells(header_row, 7).Value = "8:00"
    exportSheet.Cells(header_row, 8).Value = "8:30"
    exportSheet.Cells(header_row, 9).Value = "9:00"
    exportSheet.Cells(header_row, 10).Value = "9:30"
    exportSheet.Cells(header_row, 11).Value = "10:00"
    exportSheet.Cells(header_row, 12).Value = "10:30"
    exportSheet.Cells(header_row, 13).Value = "11:00"
    exportSheet.Cells(header_row, 14).Value = "11:30"
    exportSheet.Cells(header_row, 15).Value = "12:00"
    exportSheet.Cells(header_row, 16).Value = "12:30"
    exportSheet.Cells(header_row, 17).Value = "13:00"
    exportSheet.Cells(header_row, 18).Value = "13:30"
    exportSheet.Cells(header_row, 19).Value = "14:00"
    exportSheet.Cells(header_row, 20).Value = "14:30"
    exportSheet.Cells(header_row, 21).Value = "15:00"
    exportSheet.Cells(header_row, 22).Value = "15:30"
    exportSheet.Cells(header_row, 23).Value = "16:00"
    exportSheet.Cells(header_row, 24).Value = "16:30"
    exportSheet.Cells(header_row, 25).Value = "17:00"
    exportSheet.Cells(header_row, 26).Value = "17:30"
    exportSheet.Cells(header_row, 27).Value = "18:00"
    exportSheet.Cells(header_row, 28).Value = "18:30"
    exportSheet.Cells(header_row, 29).Value = "19:00"
    exportSheet.Cells(header_row, 30).Value = "19:30"
    exportSheet.Cells(header_row, 31).Value = "20:00"
    exportSheet.Cells(header_row, 32).Value = "20:30"
    exportSheet.Cells(header_row, 33).Value = "21:00"
    exportSheet.Cells(header_row, 34).Value = "21:30"
    exportSheet.Cells(header_row, 35).Value = "22:00"
    exportSheet.Cells(header_row, 36).Value = "22:30"

    exportSheet.Range(Cells(header_row, 5), Cells(header_row, 36)).Font.Size = 8
    exportSheet.Range(Cells(header_row, 5), Cells(header_row, 36)).Columns.AutoFit

End Function

' �V�t�g���o�͂���
Function ExportData(exportSheet As Worksheet, data() As String, ByRef header_row As Integer, ByRef array_position As Integer)
    Dim row_num As Integer
    Dim export_num As Integer
    row_num = UBound(data, 1)
    export_num = 33

    ' �s�̐������J��Ԃ�
    For i = 1 To row_num
        Dim export_row As Integer
        export_row = header_row + i
        exportSheet.Cells(export_row, 1).Value = data(i, 1)
        exportSheet.Cells(export_row, 2).Value = data(i, 2)
        exportSheet.Cells(export_row, 3).Value = data(i, 3)
        exportSheet.Cells(export_row, 4).Value = data(i, array_position)

        Select Case data(i, array_position)
            Case "A"
                ' 7:00 - 16:00
                Call ColoringCell(exportSheet, export_row, 5, 22, "Green")
            Case "B"
                ' 9:00 - 18:00
                Call ColoringCell(exportSheet, export_row, 9, 26, "Green")
            Case "C"
                ' 12:00 - 21:00
                Call ColoringCell(exportSheet, export_row, 15, 32, "Green")
            Case "D"
                ' 14:00 - 23:00
                Call ColoringCell(exportSheet, export_row, 19, 36, "Green")
            Case "�x"
                Call ColoringCell(exportSheet, export_row, 5, 36, "Gray")
        End Select
    Next i

    Call WriteLine(exportSheet, header_row, 5, 36)

End Function

' �Z���ɐF��h��
Function ColoringCell(exportSheet As Worksheet, ByRef exp_row As Integer, ByRef start_col As Integer, ByRef end_col As Integer, color As String)
    For i = start_col To end_col
        Select Case color
            Case "Green"
                exportSheet.Cells(exp_row, i).Interior.Color = RGB(60, 179, 113)
            Case "Gray"
                exportSheet.Cells(exp_row, i).Interior.Color = RGB(128, 128, 128)
            Case "White"
                exportSheet.Cells(exp_row, i).Interior.Color = RGB(255, 255, 255)
        End Select
    Next i
End Function

' �r����`��
Function WriteLine(exportSheet As Worksheet, ByRef header_row As Integer, ByRef start_col As Integer, ByRef end_col As Integer)
    Dim end_row As Integer
    end_row = exportSheet.Cells(header_row, 1).End(xlDown).Row

    exportSheet.Range(Cells(header_row, start_col), Cells(end_row, end_col)).Borders.LineStyle = xlContinuous
End Function
