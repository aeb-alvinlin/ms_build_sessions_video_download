Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k, l, m As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      k = Columns("X").Column ' 文字開始欄
      l = Columns("AD").Column ' 文字結束欄
      m = Columns("AE").Column ' 連結開始欄
      For j = k To l
        If Cells(i, m + (j - k)).Text Like "http*" Then
          Cells(i, j).HorizontalAlignment = xlLeft
          DisplayStr = Cells(i, j).Text
          Addr = Cells(i, m + (j - k)).Text
          ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, j), Address:=Addr, TextToDisplay:=DisplayStr
        End If
      Next j
    Next i
    ActiveSheet.UsedRange.Select
    Application.ScreenUpdating = True
End Sub
