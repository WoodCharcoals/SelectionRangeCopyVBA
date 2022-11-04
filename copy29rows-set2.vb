Sub CPSet2_Click()
'dynamic copy selection range code
Dim LastDataRow, PrevSetsRow, SecondDataRowCurrent, FindDataNextSet, DataSet As Long
Dim CurrentDataRange, NextDataRange, WOCurrentSet, WONextSet, WOInPrevSet  As Range

Application.CutCopyMode = False
'clear const
Range("N62:P62").ClearContents
'Last row of selection's range in copy button set1
PrevSetsRow = Range("N33").Value + 1
DataSet = Range("N33").Value + 29
'CurrentDataRange = Range("B" & PrevSetsRow & ":B" & DataSet)
'NextDataRange = Range("B" & DataSet & ":B" & DataSet + 29)

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
            Range("N62").Value = LastDataRow
            Range("O62").Value = "Case1"
            Range("P62").Value = "Basic Case"
     
        ElseIf LastDataRow - PrevSetsRow > 29 Then
            'find row of work order is not equal as PrevSetsRow to more 29 row from it
            For Each WOCurrentSet In Range("B" & PrevSetsRow & ":B" & DataSet)
                If WOCurrentSet.Value = Range("B" & DataSet).Value Then
                    SecondDataRowCurrent = WOCurrentSet.Row - 1
                    Exit For
                End If
            Next WOCurrentSet
            
              Debug.Print "2ndDataRow:" & SecondDataRowCurrent
            'find row of current val in next data set
            For Each WONextSet In Range("B" & DataSet & ":B" & DataSet + 29)
                If WONextSet.Value <> Range("B" & DataSet).Value Then
                    FindDataNextSet = WONextSet.Row
                    Range("P62").Value = "Sub_case.01"
                    Exit For
'                Else
'                    FindDataNextSet = DataSet
'                    Range("P62").Value = "Sub_case.02"
'                    Exit For
                End If
            Next WONextSet
            
            Debug.Print "FindDataNextSet:" & FindDataNextSet
            Debug.Print "FindNext:" & FindDataNextSet & ", 2nd:" & SecondDataRowCurrent
            Debug.Print "dif1:" & FindDataNextSet - SecondDataRowCurrent
            
            'return selection range
            If FindDataNextSet - SecondDataRowCurrent > 29 Then

                Range("B" & PrevSetsRow & ":L" & SecondDataRowCurrent).Copy
                Range("N62").Value = SecondDataRowCurrent
                Range("O62").Value = "Case2.3"

            ElseIf Range("B" & DataSet).Value = Range("B" & DataSet + 1).Value Then

                Range("B" & PrevSetsRow & ":L" & SecondDataRowCurrent).Copy
                Range("N62").Value = SecondDataRowCurrent
                Range("O62").Value = "Case2.1"
                
            ElseIf Range("B" & DataSet).Value <> Range("B" & DataSet + 1).Value Then
                Range("B" & PrevSetsRow & ":L" & FindDataNextSet).Copy
                Range("N62").Value = FindDataNextSet
                Range("O62").Value = "Case2.2"
                
            Else
                Exit Sub
            End If

        End If
    
    End If

End With

    If Application.CutCopyMode = xlCopy Then
        'Debug.Print "Copy status:=Done"
        Me.CPSet2.BackColor = vbGreen 'RGB(50, 205, 50)
    Else
        'Debug.Print "Copy status:=Failed"
        Me.CPSet2.BackColor = vbRed 'RGB(255, 0, 0)
    End If

End Sub