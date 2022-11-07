Sub CPSet1_Click()
'dynamic copy selection range code
Dim TotalRows, EndPrevSetsRow, CurrentSetRow, NextSetRow, _
BeginCurrentRow, EndCurrentRow, BeginNextSetRow, EndNextSetRow As Long
Dim FindCurrentRow, FindNextRow As Long

Dim CurrentDataRange, NextDataRange, getPrevRow, setResultRow  As Range

'set initial setting of procedure
Application.CutCopyMode = False

'set var from previous selection rang
If Range("N4").Value < 3 Then
    Range("N4").Value = 3
End If
EndPrevSetsRow = Range("N4").Value
BeginCurrentRow = Range("N4").Value + 1
EndCurrentRow = BeginCurrentRow + 28

'set var for relative range
Set getPrevRow = Range("N4")
Set setResultRow = Range("N33")
Set CurrentDataRange = Range("B" & BeginCurrentRow & ":B" & EndCurrentRow)
Set NextDataRange = Range("B" & EndCurrentRow + 1 & ":B" & EndCurrentRow + 29)

Debug.Print "CurrentRange=Row(" & BeginCurrentRow & ":" & EndCurrentRow & ")"
Debug.Print "NextDataRange=Row(" & EndCurrentRow + 1 & ":" & EndCurrentRow + 29 & ")"

With Range("L4:L206")

    TotalRows = NotZero.Row - 1

    For Each CurrentDataRange In CurrentDataRange
        If CurrentDataRange.Value = Range("B" & EndCurrentRow).Value Then
            FindCurrentRow = CurrentDataRange.Row
            Exit For
        End If
    Next CurrentDataRange
              
    'find row of current val in next data set
    For Each NextDataRange In NextDataRange
        If NextDataRange.Value <> Range("B" & EndCurrentRow).Value Then
            FindNextRow = NextDataRange.Row
            Exit For
        Else
            FindNextRow = EndCurrentRow + 28
            
        End If
    Next NextDataRange
'FindCurrentRow = Find row has same value as EndCurrentRow row
'FindNextRow  find row has same value as EndCurrentRow after EndCurrentRow row
'DiffRange = FindNextRow - FindCurrentRow

Debug.Print "TR-BCR:" & TotalRows - BeginCurrentRow
Debug.Print "FNR-FCR:" & FindNextRow - FindCurrentRow
Debug.Print "FCR:" & FindCurrentRow & ",BCR:" & BeginCurrentRow

If EndPrevSetsRow < 3 Then
    setResultRow.Offset(-29, 0).Value = "0"
    Debug.Print "Failed1"
    Exit Sub

ElseIf TotalRows - BeginCurrentRow <= 28 Then
        Range("B" & BeginCurrentRow & ":L" & EndCurrentRow).Copy
        setResultRow.Value = EndCurrentRow
        setResultRow.Offset(0, 1).Value = "Case1"
        setResultRow.Offset(0, 2).Value = "Basic Case"

ElseIf TotalRows - BeginCurrentRow > 28 Then


    If FindCurrentRow = BeginCurrentRow Then

        Range("B" & BeginCurrentRow & ":L" & FindNextRow - 1).Copy
        setResultRow.Value = FindNextRow - 1
        setResultRow.Offset(0, 1).Value = "Case2"
        setResultRow.Offset(0, 2).Value = "Subcase1.1"
        setResultRow.Offset(-29, 0).Value = 3
        Debug.Print "Case2"
        
    ElseIf FindCurrentRow <> BeginCurrentRow Then
        Range("B" & BeginCurrentRow & ":L" & FindCurrentRow - 1).Copy
        setResultRow.Value = FindCurrentRow - 1
        setResultRow.Offset(0, 1).Value = "Case3"
        setResultRow.Offset(0, 2).Value = "Subcase1.1"
        Debug.Print "Case3"

    ElseIf FindCurrentRow >= 18 Then
        Range("B" & BeginCurrentRow & ":L" & FindCurrentRow - 1).Copy
        setResultRow.Offset(-29, 0).Value = FindCurrentRow - 1
        setResultRow.Offset(-29, 1).Value = "ReverseCase1"
        setResultRow.Offset(-29, 2).Value = "Subcase1.1"
        Debug.Print "RCase1"
        
    End If
    
End If

End With

    If Application.CutCopyMode = xlCopy Then
        ''Debug.Print "Copy status:=Done"
        Me.CPSet1.BackColor = vbGreen 'RGB(50, 205, 50)
    Else
        ''Debug.Print "Copy status:=Failed"
        Me.CPSet1.BackColor = vbRed 'RGB(255, 0, 0)
    End If

End Sub