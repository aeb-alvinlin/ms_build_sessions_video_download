Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      If Cells(i, 6).Value Like "*new*" Then
        Cells(i, 3).Value = "New"
      ElseIf Cells(i, 10).Value Like "*new*" Then
        Cells(i, 3).Value = "New"
      End If
    Next i
    Application.ScreenUpdating = True
End Sub


