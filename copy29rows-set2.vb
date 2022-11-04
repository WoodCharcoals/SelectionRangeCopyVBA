Sub CPSet2_Click()
'dynamic copy selection range code
Dim LastDataRow, PrevSetsRow, SecondDataRowCurrent, FindDataNextSet, DataSet As Long
Dim CurrentDataRange, NextDataRange, WOCurrentSet, WONextSet, SetResultRow  As Range

'set initial setting of procedure
Application.CutCopyMode = False
Range("N62:P62").ClearContents
'set var from previous selection range
PrevSetsRow = Range("N33").Value + 1
DataSet = Range("N33").Value + 29
'set var for relative range
Set SetResultRow = Range("N62")
Set CurrentDataRange = Range("B" & PrevSetsRow & ":B" & DataSet)
Set NextDataRange = Range("B" & DataSet & ":B" & DataSet + 29)

Debug.Print "CurrentRange=Row(" & PrevSetsRow & ":" & DataSet & ")"
Debug.Print "NextDataRange=Row(" & DataSet & ":" & DataSet + 29 & ")"

With Range("L4:L206")

    LastDataRow = NotZero.Row - 1

    If Range("B33").Value = 0 Then
        Exit Sub
    ElseIf PrevSetsRow = 0 Then
        Exit Sub
    Else
        'basic condition
        If LastDataRow - PrevSetsRow <= 29 Then
            Range("B" & PrevSetsRow & ":L" & LastDataRow).Copy
            SetResultRow.Value = LastDataRow
            SetResultRow.Offset(0, 1).Value = "Case1"
            SetResultRow.Offset(0, 2).Value = "Basic Case"
     
        ElseIf LastDataRow - PrevSetsRow > 29 Then
            'find row of work order is not equal as PrevSetsRow to more 29 row from it
            For Each WOCurrentSet In CurrentDataRange
                If WOCurrentSet.Value = Range("B" & DataSet).Value Then
                    SecondDataRowCurrent = WOCurrentSet.Row - 1
                    Exit For
                End If
            Next WOCurrentSet
              
            'find row of current val in next data set
            For Each WONextSet In NextDataRange
                If WONextSet.Value <> Range("B" & DataSet).Value Then
                    FindDataNextSet = WONextSet.Row - 1
                    Exit For
                Else
                    FindDataNextSet = DataSet + 29
                    
                End If
            Next WONextSet
            
            Debug.Print "FindNext:" & FindDataNextSet & ", 2ndRow:" & SecondDataRowCurrent
            Debug.Print "Diff FindNext-PrevSet:" & FindDataNextSet - PrevSetsRow
            
            'return selection range
'            If FindDataNextSet - PrevSetsRow > 29 Then
'
'                Range("B" & PrevSetsRow & ":L" & SecondDataRowCurrent).Copy
'                SetResultRow.Value = SecondDataRowCurrent
'                SetResultRow.Offset(0, 1).Value = "Case2.3"
                
            If Range("B" & DataSet).Value = Range("B" & DataSet + 1).Value Then

                Range("B" & PrevSetsRow & ":L" & SecondDataRowCurrent).Copy
                SetResultRow.Value = SecondDataRowCurrent
                SetResultRow.Offset(0, 1).Value = "Case2.1"
                
            ElseIf Range("B" & DataSet).Value <> Range("B" & DataSet + 1).Value Then
                Range("B" & PrevSetsRow & ":L" & FindDataNextSet).Copy
                SetResultRow.Value = FindDataNextSet
                SetResultRow.Offset(0, 1).Value = "Case2.2"
                
            Else
                SetResultRow.Offset(0, 1) = "Case3:Failed"
                Exit Sub
            End If

        End If
    
    End If

End With

    If Application.CutCopyMode = xlCopy Then
        ''Debug.Print "Copy status:=Done"
        Me.CPSet2.BackColor = vbGreen 'RGB(50, 205, 50)
    Else
        ''Debug.Print "Copy status:=Failed"
        Me.CPSet2.BackColor = vbRed 'RGB(255, 0, 0)
    End If
