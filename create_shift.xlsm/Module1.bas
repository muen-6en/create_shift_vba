Attribute VB_Name = "Module1"
Sub CreateMonthSheet()
    Application.DisplayAlerts = False
    
    'Get WorkBook
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    'Get Value.
    Dim month As String
    Dim macro As Worksheet
    Set macro = Worksheets("�}�N��")
    Dim targetMonth As String
    Dim targetTerm As String
    targetMonth = macro.Range("F2").Value
    targetTerm = macro.Range("F3").Value

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
    MonthSheet.Range("C2").Value = "�Ζ��敪"
    MonthSheet.Range("C3").Value = "A"
    MonthSheet.Range("C4").Value = "B"
    MonthSheet.Range("C5").Value = "C"
    MonthSheet.Range("C6").Value = "D"

    ' Start
    MonthSheet.Range("D2").Value = "�n��"
    MonthSheet.Range("D3").Value = "7:00"
    MonthSheet.Range("D4").Value = "9:00"
    MonthSheet.Range("D5").Value = "12:00"
    MonthSheet.Range("D6").Value = "14:00"

    ' End
    MonthSheet.Range("E2").Value = "�I��"
    MonthSheet.Range("E3").Value = "16:00"
    MonthSheet.Range("E4").Value = "18:00"
    MonthSheet.Range("E5").Value = "21:00"
    MonthSheet.Range("E6").Value = "23:00"

    ' other
    MonthSheet.Range("F2").Value = "���̑�"
    MonthSheet.Range("F3").Value = "�x�F�x��"
    MonthSheet.Range("F4").Value = "���F���x"

    ' Position
    ' �}�N���V�[�g�̓���̃Z������l���擾����
    Dim positions() As String
    positions() = GetBelongs(positions(), macro, "E")
    Call SetBelongs(positions(), MonthSheet, "A")
    MonthSheet.Range("A8").Value = "��E"
    MonthSheet.Range("A9").Value = "�{�ݒ�"
    MonthSheet.Range("A10").Value = "�Ј�"
    MonthSheet.Range("A11").Value = "�Ј�"
    MonthSheet.Range("A12").Value = "�_��Ј�"
    MonthSheet.Range("A13").Value = "�p�[�g"

    ' Member
    Dim members() As String
    members() = GetBelongs(members(), macro, "F")
    Call SetBelongs(members(), MonthSheet, "B")
    MonthSheet.Range("B8").Value = "���O"
    MonthSheet.Range("B9").Value = "�����O�q"
    MonthSheet.Range("B10").Value = "�Ј����Y"
    MonthSheet.Range("B11").Value = "�Ј��S��"
    MonthSheet.Range("B12").Value = "�_��Ԏq"
    MonthSheet.Range("B13").Value = "�m�Ȑm��"

    ' Work
    MonthSheet.Range("C8").Value = "�S��"

    ' Day
    ' Dim days(15) As Integet
    ' day() = GetDays(days, targetTerm)
    ' GetDays()

    ' UI
    ' TODO: A8-�ŏI���̓Z�����擾���Čr����ݒ肷��
    MonthSheet.Range("A8:B13").Borders.LineStyle = xlContinuous

End Function

Function GetBelongs(collections() As String , macro As Worksheet, col As String) As String()
    ' macro�V�[�g��Row6�ȉ��̒l���󔒂܂Ń��[�v
    ' �󔒂܂ł̐����J�E���g
    ' collections�̗v�f�����J�E���g��ReDim
    ' collections�ɖ�E�̒l��ݒ肷��

    GetPosition = collections
End Function

Function SetBelongs(collections() As String , MonthSheet As Worksheet, col As String)
    ' MonthSheet�ɒl��ݒ肷��
    Exit Function
End Function

Function GetDays(days() As Integer, ByVal targetTerm As String) As Integer()
    Dim d As Integer

    If targetTerm = "�O��" Then
        d = 1
    Else If targetTerm = "�㔼" Then
        d = 16
    End If

    For i = d To d + 14
        days(i) = d
    Next i

    GetDays = days
End Function
