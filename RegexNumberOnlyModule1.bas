Attribute VB_Name = "Module1"
Option Base 1
Sub Macro1()
    Dim RegEx As RegExp
    Set RegEx = New RegExp
    RegEx.Pattern = "([0-9]{2})"
    
    Dim m As Match, Matches As MatchCollection

    Dim Str, MyString As String
    
    Dim i, j, k As Long
    Dim WidthAlignment As Variant
    Application.ScreenUpdating = False
    For i = 2 To ActiveSheet.UsedRange.Rows.Count
    MyString = Cells(i, 2).Value
    Set Matches = RegEx.Execute(MyString)
    If Matches.Count > 0 Then
      Cells(i, 3).Value = Matches(0).SubMatches(0)
    End If
    Next i
    ActiveSheet.UsedRange.Select
    Selection.Font.Name = "Yu Gothic"
    Selection.Font.Size = 10
    Selection.VerticalAlignment = xlCenter
    Application.ScreenUpdating = True
End Sub
