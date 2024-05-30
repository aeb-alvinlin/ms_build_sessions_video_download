Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    j = Columns("D").Column ' 連結開始欄
    k = Columns("C").Column ' 文字結束欄
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      If i > 1 And i Mod 2 = 1 Then
        Range(Cells(i, 1), Cells(i, 13)).Interior.Color = RGB(255, 242, 204)
      End If
      If Cells(i, j).Text Like "http*" Then
        Cells(i, k).HorizontalAlignment = xlLeft
        DisplayStr = Cells(i, k).Text
        AddrHrefLink = Cells(i, j).Text
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, k), Address:=AddrHrefLink, TextToDisplay:=DisplayStr
      End If
    Next i
    ActiveSheet.UsedRange.Select
    Selection.Font.Name = "Yu Gothic"
    Selection.Font.Size = 10
    Selection.VerticalAlignment = xlCenter
    Application.ScreenUpdating = True
End Sub
