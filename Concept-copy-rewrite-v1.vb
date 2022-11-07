Sub CPSet1_Click()
'dynamic copy selection range code
Dim TotalRows, EndPrevSetsRow, CurrentSetRow, NextSetRow, _
BeginCurrentRow, EndCurrentRow, BeginNextSetRow, EndNextSetRow As Long
dim FindCurrentRow, FindNextRow as long

Dim CurrentDataRange, NextDataRange, getPrevRow, setResultRow  As Range

'set initial setting of procedure
Application.CutCopyMode = False

'set var from previous selection rang
if Range("N4").Value  < 3 then
    Range("N4").Value = 3
end if
EndPrevSetsRow = Range("N4").Value 
BeginCurrentRow = Range("N4").Value+1
EndCurrentRow = BeginCurrentRow + 28

'set var for relative range
set getPrevRow = range("N4")
Set setResultRow = Range("N33")
Set CurrentDataRange = Range("B" & BeginCurrentRow & ":B" & EndCurrentRow)
Set NextDataRange = Range("B" & EndCurrentRow & ":B" & EndCurrentRow + 29)

Debug.Print "CurrentRange=Row(" & BeginCurrentRow & ":" & EndCurrentRow & ")"
Debug.Print "NextDataRange=Row(" & EndCurrentRow & ":" & EndCurrentRow + 29 & ")"

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
            FindNextRow = NextDataRange.Row - 1
            Exit For
        Else
            FindNextRow = EndCurrentRow + 28
            
        End If
    Next NextDataRange
'FindCurrentRow = Find row has same value as EndCurrentRow row 
'FindNextRow  find row has same value as EndCurrentRow after EndCurrentRow row
'DiffRange = FindNextRow - FindCurrentRow 

If EndPrevSetsRow < 3 then
    setResultRow.offset(-29,0).Value = "0"
    Exit Sub

Else
    if TotalRows - BeginCurrentRow <=28 then
        Range("B" & BeginCurrentRow & ":L" & EndCurrentRow).Copy
        setResultRow.Value = EndCurrentRow
        setResultRow.Offset(0, 1).Value = "Case1"
        setResultRow.Offset(0, 2).Value = "Basic Case"  

    ElseIf FindNextRow - FindCurrentRow > 29 then

        if  FindCurrentRow < BeginCurrentRow + 14 then
            Range("B" & BeginCurrentRow & ":L" & FindCurrentRow -1).Copy
            setResultRow.Offset(-29,0 ).Value = FindCurrentRow -1
            setResultRow.Offset(-29, 1).Value = "ReverseCase1"
            setResultRow.Offset(-29, 2).Value = "Subcase1.1"   

        elseif FindCurrentRow > BeginCurrentRow + 14 then
            Range("B" & BeginCurrentRow & ":L" & EndCurrentRow).Copy
            setResultRow.Value = EndCurrentRow
            setResultRow.Offset(0, 1).Value = "ReverseCase1"
            setResultRow.Offset(0, 2).Value = "Subcase1.1"       
        else

        end if
    else 
        
    End If
End If

end with

'-------------------------------------------------------------------
'-------------------------------old---------------------------------
'-------------------------------------------------------------------

    If PrevSetsRow < 4 Then
        setResultRow.offset(-29,0).Value = "0"
        Exit Sub
    Else
        'basic condition
        If LastDataRow - PrevSetsRow <= 29 Then
            Range("B" & PrevSetsRow & ":L" & LastDataRow).Copy
            setResultRow.Value = LastDataRow
            setResultRow.Offset(0, 1).Value = "Case1"
            setResultRow.Offset(0, 2).Value = "Basic Case"
            
        ElseIf LastDataRow - PrevSetsRow > 29 Then
            'find row of work order is not equal as PrevSetsRow to more 29 row from it
            
            
            Debug.Print "FindNext:" & FindDataNextSet & ", 2ndRow:" & SecondDataRowCurrent
            Debug.Print "Diff FindNext-PrevSet:" & FindDataNextSet - PrevSetsRow
            
            'return selection range
            If FindDataNextSet - PrevSetsRow > 29 Then

                Range("B" & PrevSetsRow & ":L" & EndDataSet).Copy
                setResultRow.Value = EndDataSet
                setResultRow.Offset(0, 1).Value = "Case2.3"
                
           ElseIf Range("B" & EndDataSet).Value = Range("B" & EndDataSet + 1).Value Then

                Range("B" & PrevSetsRow & ":L" & SecondDataRowCurrent).Copy
                Range("B" & PrevSetsRow).Select
                setResultRow.Value = SecondDataRowCurrent
                setResultRow.Offset(0, 1).Value = "Case2.1"
                
            ElseIf Range("B" & EndDataSet).Value <> Range("B" & EndDataSet + 1).Value Then
                Range("B" & PrevSetsRow & ":L" & FindDataNextSet).Copy
                setResultRow.Value = FindDataNextSet
                setResultRow.Offset(0, 1).Value = "Case2.2"
                
            Else
                setResultRow.Offset(0, 1) = "Case3:Failed"
                Exit Sub
            End If

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