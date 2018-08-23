Public Sub quicksort_Main()
Dim StartTime As Double
Dim SecondsElapsed As Double

StartTime = Timer

Dim rng As Range
Dim lng As Long
lng = Range("A" & Rows.Count).End(xlUp).Row
Set rng = Range(Cells(1, "A"), Cells(lng, "A"))

quicksort rng, rng.Row, rng.Rows.Count

SecondsElapsed = Round(Timer - StartTime, 2)


Set rng = Nothing


Debug.Print "TIME TO FINISH: " & SecondsElapsed & " SECONDS" & vbNewLine


End Sub

Public Sub quicksort(ByRef rng As Range, left As Integer, right As Integer)

Dim i As Integer: i = left
Dim j As Integer: j = right

Dim tmp As Integer

Dim pivot As Integer
pivot = rng(((left + right) / 2), "A").Value

'Debug.Print "RANGE_C: " & rng.Rows.Count & " LEFT: " & left & " RIGHT: " & right & " I: " & i & " J: " & j & " PIVOT: " & pivot

While i <= j
    
    While rng(i, "A").Value < pivot
        i = i + 1
    Wend

    While rng(j, "A").Value > pivot
        j = j - 1
    Wend

If i <= j Then

        tmp = rng(i, "A").Value
        rng(i, "A").Value = rng(j, "A").Value
        rng(j, "A").Value = tmp
        i = i + 1
        j = j - 1
End If


Wend

If left < j Then
quicksort rng, left, j
End If

If i < right Then
quicksort rng, i, right
End If


'Debug.Print "RANGE_C: " & rng.Rows.Count & " LEFT: " & left & " RIGHT: " & right & " I: " & i & " J: " & j & " PIVOT: " & pivot


End Sub