Attribute VB_Name = "Module1"
Sub CreateMonthSheet()
    Application.DisplayAlerts = False
    
    '���[�N�u�b�N�̎擾
    Dim wb As Workbook
    Set wb = Workbooks("�V�t�g�쐬�c�[��.xlsm")
    Debug.Print ("�u�b�N���F" & wb.Name)
    
    '���͒l�̎擾
    Dim month As String
    Dim macro As Worksheet
    Set macro = Worksheets("�}�N��")
    month = macro.Range("F2").Value & "�� " & macro.Range("F3").Value
    Debug.Print (month)
    
    '�V�K�V�[�g���쐬
    Worksheets().Add After:=Worksheets(Worksheets.Count)
    Dim MonthSheet As Worksheet
    Set MonthSheet = ActiveSheet
    
    '�V�[�g���̎擾
    Dim SheetCnt As Long
    Dim SheetName() As String
    SheetCnt = wb.Sheets.Count
    ReDim SheetName(1 To SheetCnt)
    For i = 1 To SheetCnt
        SheetName(i) = Sheets(i).Name
    Next i

    '�V�[�g����ύX
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

    
    
    '�Œ�o��
    MonthSheet.Cells.Clear
    'MonthSheet.Range("").Value =
    MonthSheet.Range("A1").Value = month
    MonthSheet.Range("A1").Font.Size = 14
    MonthSheet.Range("C2").Value = "�Ζ��敪"
    MonthSheet.Range("D2").Value = "�n��"
    MonthSheet.Range("E2").Value = "�I��"
    MonthSheet.Range("F2").Value = "���̑�"
    
End Sub
