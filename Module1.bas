Attribute VB_Name = "Module1"
Option Explicit
Sub �̍وꊇ����()
Attribute �̍وꊇ����.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim �I�s As Long, �s As Long
    Dim �͈� As Range
    With Workbooks("�ėp�ُ̍���.xlsm").Sheets("�ُ̍����ݒ�")
        �I�s = .Cells(Rows.Count, 1).End(xlUp).Row
        For Each �͈� In Selection
            If �͈� = "" Or �͈�.HasFormula = True Then GoTo �L�����Z��
            �͈�.Characters.PhoneticCharacters = ""
            ReDim �u���ݒ�(3 To �I�s, 1 To 2)
            If �I�s > 2 Then
                For �s = 3 To �I�s
                    �u���ݒ�(�s, 1) = .Cells(�s, 1)
                    �u���ݒ�(�s, 2) = .Cells(�s, 2)
                Next
                For �s = 3 To �I�s
                    �͈� = Replace(�͈�, �u���ݒ�(�s, 1), �u���ݒ�(�s, 2))
                Next
            End If
            Select Case .Range("�S���p�L�[")
                Case "�S�p����": �͈� = StrConv(�͈�, vbWide)
                Case "���p����": �͈� = StrConv(�͈�, vbNarrow)
            End Select
�L�����Z��:
        Next
    End With
End Sub
