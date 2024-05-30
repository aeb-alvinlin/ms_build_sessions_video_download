Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      If Cells(i, 2).Value Like "*Will be recorded*" Then
        Cells(i, 3).Value = "Will be recorded"
      End If
      If Cells(i, 2).Value Like "*Will not be recorded*" Then
        Cells(i, 3).Value = "Will not be recorded"
      End If
      If Cells(i, 2).Value Like "*In Seattle only*" Then
        Cells(i, 4).Value = "In Seattle only"
      End If
      If Cells(i, 2).Value Like "*+ Online*" Then
        Cells(i, 4).Value = "Online"
      End If
    Next i
    Application.ScreenUpdating = True
End Sub


