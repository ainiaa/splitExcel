Function DoSplitByColumn(ccstr As String, savepath As String)
    Dim arr, d As Object, k, t, i&, lc%, rng As Range, c As Long
    c = CLng(ccstr)
    If c = 0 Then Exit Sub
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    arr = [a1].CurrentRegion
    lc = UBound(arr, 2)
    Set rng = [a1].Resize(, lc)
    Set d = CreateObject("scripting.dictionary")
    For i = 2 To UBound(arr)
        If Not d.Exists(arr(i, c)) Then
            Set d(arr(i, c)) = Cells(i, 1).Resize(1, lc)
        Else
            Set d(arr(i, c)) = Union(d(arr(i, c)), Cells(i, 1).Resize(1, lc))
        End If
    Next
    k = d.Keys
    t = d.Items
    For i = 0 To d.Count - 1
        With Workbooks.Add(xlWBATWorksheet)
            rng.Copy .Sheets(1).[a1]
            t(i).Copy .Sheets(1).[a2]
            .SaveAs Filename:=savepath & "\" & k(i) & ".xlsx"
            .Close
        End With
    Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Function

Function fun()
    MsgBox("111")
End Function
