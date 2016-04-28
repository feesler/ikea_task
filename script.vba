Public Function IsInArray(stringToBeFound As Variant, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function


Public Function getInd(ByVal name As String)
Dim c As Range
Set Target = ActiveWorkbook.Worksheets("SHIFTCALC")

ind = 1
For Each c In Target.Range("A1:P1")
    If c = "" Or c = name Then
        c.Value = name
        Exit For
    Else
        ind = ind + 1
    End If
Next c

getInd = ind
End Function

Sub ExtractSM0()

Dim parts As Variant
parts = Array("MH03", "MH06", "MH10", "MH11", "MH12", "MH13", "MH14", "MH15", "MH16", "MH18", "MH19", "MH92", "SR07", "SR09")

Dim c As Range
Dim otd(50)
Dim sum(50)
Dim lastRow As Long
Dim lastCol As Long

Set Source = ActiveWorkbook.Worksheets("PRENOTE")
Set Target = ActiveWorkbook.Worksheets("SHIFTCALC")

Target.Range("A1:R200").ClearContents

otdSize = 0
For Each c In Source.Range("C2:C2000")
    If c = "0" Then
    hVal = Source.Cells(c.Row, 8).Value
    jVal = Source.Cells(c.Row, 10).Value
    lVal = Source.Cells(c.Row, 12).Value

    If hVal = 0 And IsInArray(lVal, parts) Then
    ' Obtain column index for part (HFB)
    col = getInd(lVal)
    ' Increase count of values stored for part
    cnt = otd(col - 1)
    cnt = cnt + 1
    otd(col - 1) = cnt
    ' Save value
    With Target.Cells(cnt + 1, col)
        .NumberFormat = "0.00000"
        .Value = jVal
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    ' Sum calculation here
    cursum = sum(col - 1)
    cursum = cursum + jVal
    sum(col - 1) = cursum
    ' Update sum at table
    Target.Cells(col + 1, 17).Value = lVal
    With Target.Cells(col + 1, 18)
        .NumberFormat = "0.00000"
        .Value = cursum
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    End If
    End If
Next c

Target.Cells.Sort _
Key1:=Range("A1"), Order1:=xlAscending, _
Key2:=Range("L1"), Order2:=xlAscending, _
Orientation:=xlSortRows

Target.Range("Q2:R50").Sort _
Key1:=Range("Q2:Q50"), Order1:=xlAscending, _
Orientation:=xlSortColumns

End Sub

