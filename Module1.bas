Attribute VB_Name = "Module1"
Option Explicit
Sub 体裁一括処理()
Attribute 体裁一括処理.VB_ProcData.VB_Invoke_Func = "q\n14"
    Dim 終行 As Long, 行 As Long
    Dim 範囲 As Range
    With Workbooks("汎用体裁処理.xlsm").Sheets("体裁処理設定")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        For Each 範囲 In Selection
            If 範囲 = "" Or 範囲.HasFormula = True Then GoTo キャンセル
            範囲.Characters.PhoneticCharacters = ""
            ReDim 置換設定(3 To 終行, 1 To 2)
            If 終行 > 2 Then
                For 行 = 3 To 終行
                    置換設定(行, 1) = .Cells(行, 1)
                    置換設定(行, 2) = .Cells(行, 2)
                Next
                For 行 = 3 To 終行
                    範囲 = Replace(範囲, 置換設定(行, 1), 置換設定(行, 2))
                Next
            End If
            Select Case .Range("全半角キー")
                Case "全角統一": 範囲 = StrConv(範囲, vbWide)
                Case "半角統一": 範囲 = StrConv(範囲, vbNarrow)
            End Select
キャンセル:
        Next
    End With
End Sub
