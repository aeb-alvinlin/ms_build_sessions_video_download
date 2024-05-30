Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To 426
      If Cells(i, 5).Value Like "*AI*" Then
        Cells(i, 2).Value = "AI"
      End If
      If Cells(i, 8).Value Like "*AI*" Then
        Cells(i, 2).Value = "AI"
      End If
      If Cells(i, 5).Value Like "*Copilot*" Then
        Cells(i, 2).Value = "Copilot"
      End If
      If Cells(i, 8).Value Like "*Copilot*" Then
        Cells(i, 2).Value = "Copilot"
      End If
      If Cells(i, 5).Value Like "*Security*" Then
        Cells(i, 2).Value = "Security"
      End If
      If Cells(i, 8).Value Like "*Security*" Then
        Cells(i, 2).Value = "Security"
      End If
    Next i
    Application.ScreenUpdating = True
End Sub


