Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      If i > 1 And i Mod 2 = 1 Then
        Range(Cells(i, 1), Cells(i, 13)).Interior.Color = RGB(255, 242, 204)
      End If
      If Cells(i, 2).Text Like "http*" Then
        Cells(i, 2).HorizontalAlignment = xlLeft
        DisplayStr = Cells(i, 1).Text
        Str = Cells(i, 2).Text
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 2), Address:=Str, TextToDisplay:=DisplayStr
      End If
    Next i
    ActiveSheet.UsedRange.Select
    Selection.Font.Name = "Yu Gothic"
    Selection.Font.Size = 10
    Selection.VerticalAlignment = xlCenter
    Application.ScreenUpdating = True
End Sub
