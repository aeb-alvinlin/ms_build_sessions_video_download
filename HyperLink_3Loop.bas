Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k, l, m As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      k = Columns("J").Column ' ��r�}�l��
      l = Columns("Q").Column ' ��r������
      m = Columns("R").Column ' �s���}�l��
      For j = k To l
        If Cells(i, m + (j - k)).Text Like "http*" Then
          Cells(i, j).HorizontalAlignment = xlLeft
          DisplayStr = Cells(i, j).Text
          Addr = Cells(i, m + (j - k)).Text
          ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, j), Address:=Addr, TextToDisplay:=DisplayStr
        End If
      Next j
      k = Columns("Z").Column ' ��r�}�l��
      l = Columns("AF").Column ' ��r������
      m = Columns("AG").Column ' �s���}�l��
      For j = k To l
        If Cells(i, m + (j - k)).Text Like "http*" Then
          Cells(i, j).HorizontalAlignment = xlLeft
          DisplayStr = Cells(i, j).Text
          Addr = Cells(i, m + (j - k)).Text
          ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, j), Address:=Addr, TextToDisplay:=DisplayStr
        End If
      Next j
      k = Columns("AO").Column ' ��r�}�l��
      l = Columns("AU").Column ' ��r������
      m = Columns("AV").Column ' �s���}�l��
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
