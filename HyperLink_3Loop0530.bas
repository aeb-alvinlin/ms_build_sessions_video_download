Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k, l, m As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      k = Columns("J").Column ' ��r�}�l��
      l = Columns("R").Column ' ��r������
      m = Columns("S").Column ' �s���}�l��
      For j = k To l
        If Cells(i, m + (j - k)).Text Like "http*" Then
          Cells(i, j).HorizontalAlignment = xlLeft
          DisplayStr = Cells(i, j).Text
          Addr = Cells(i, m + (j - k)).Text
          ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, j), Address:=Addr, TextToDisplay:=DisplayStr
        End If
      Next j
      k = Columns("AB").Column ' ��r�}�l��
      l = Columns("AH").Column ' ��r������
      m = Columns("AI").Column ' �s���}�l��
      For j = k To l
        If Cells(i, m + (j - k)).Text Like "http*" Then
          Cells(i, j).HorizontalAlignment = xlLeft
          DisplayStr = Cells(i, j).Text
          Addr = Cells(i, m + (j - k)).Text
          ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, j), Address:=Addr, TextToDisplay:=DisplayStr
        End If
      Next j
      k = Columns("AQ").Column ' ��r�}�l��
      l = Columns("AW").Column ' ��r������
      m = Columns("AX").Column ' �s���}�l��
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
