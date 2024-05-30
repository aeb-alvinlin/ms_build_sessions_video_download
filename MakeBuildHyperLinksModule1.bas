Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim Str As String
    Dim i, j, k As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
      If Cells(i, 33).Text Like "http*" Then
        Cells(i, 26).HorizontalAlignment = xlLeft
        DisplayStr = Cells(i, 26).Text
        Addr = Cells(i, 33).Text
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 26), Address:=Addr, TextToDisplay:=DisplayStr
      End If
      If Cells(i, 34).Text Like "http*" Then
        Cells(i, 27).HorizontalAlignment = xlLeft
        DisplayStr = Cells(i, 27).Text
        Addr = Cells(i, 34).Text
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 27), Address:=Addr, TextToDisplay:=DisplayStr
      End If
      If Cells(i, 35).Text Like "http*" Then
        Cells(i, 28).HorizontalAlignment = xlLeft
        DisplayStr = Cells(i, 28).Text
        Addr = Cells(i, 35).Text
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 28), Address:=Addr, TextToDisplay:=DisplayStr
      End If
      If Cells(i, 36).Text Like "http*" Then
        Cells(i, 29).HorizontalAlignment = xlLeft
        DisplayStr = Cells(i, 29).Text
        Addr = Cells(i, 36).Text
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 29), Address:=Addr, TextToDisplay:=DisplayStr
      End If
      If Cells(i, 37).Text Like "http*" Then
        Cells(i, 30).HorizontalAlignment = xlLeft
        DisplayStr = Cells(i, 30).Text
        Addr = Cells(i, 37).Text
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 30), Address:=Addr, TextToDisplay:=DisplayStr
      End If
      If Cells(i, 38).Text Like "http*" Then
        Cells(i, 31).HorizontalAlignment = xlLeft
        DisplayStr = Cells(i, 31).Text
        Addr = Cells(i, 38).Text
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 31), Address:=Addr, TextToDisplay:=DisplayStr
      End If
      If Cells(i, 39).Text Like "http*" Then
        Cells(i, 32).HorizontalAlignment = xlLeft
        DisplayStr = Cells(i, 32).Text
        Addr = Cells(i, 39).Text
        ActiveSheet.Hyperlinks.Add Anchor:=Cells(i, 32), Address:=Addr, TextToDisplay:=DisplayStr
      End If
    Next i
    ActiveSheet.UsedRange.Select
    Application.ScreenUpdating = True
End Sub
