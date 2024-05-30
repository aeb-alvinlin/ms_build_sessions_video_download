Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      If Cells(i, 5).Value Like "*Security*" Then
        Cells(i, 8).Value = "Security"
      ElseIf Cells(i, 5).Value Like "*OpenAI*" Then
        Cells(i, 8).Value = "OpenAI"
      End If
    Next i
    Application.ScreenUpdating = True
End Sub


