Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      If Cells(i, 3).Value Like "*Security*" Then
        Cells(i, 5).Value = "Security"
      End If
      If Cells(i, 3).Value Like "*AI*" Then
        Cells(i, 6).Value = "AI"
      End If
      If Cells(i, 4).Value Like "*Security*" Then
        Cells(i, 7).Value = "Security"
      End If
    Next i
    Application.ScreenUpdating = True
End Sub


