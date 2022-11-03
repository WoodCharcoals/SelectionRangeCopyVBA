Sub CPSet2_Click()
'dynamic copy selection range code
Dim LastDataRow, getPreviousSetsRow,SecondDataRowCurrent,FindDataNextSet As Long
Dim WOLastRowOfSet, WOEqPreviousSet, WOCurrentSet, WOInNextSet, WOInPreviousSet  As Range

Application.CutCopyMode = False

'Last row of selection's range in copy button set1
getPreviousSetsRow = Range("N33").Value + 1

With Range("L4:L206")

    LastDataRow = NotZero.Row - 1

    If Range("B33").Value = 0 Then
        Exit Sub
    ElseIf getPreviousSetsRow = 0 Then
        Exit Sub
    Else
        'basic condition
        If LastDataRow - getPreviousSetsRow <= 29 Then
            Range("B" & getPreviousSetsRow & ":L" & LastDataRow).Copy
            Range("N62").Value = LastDataRow
            Range("O62").Value = "Case1"
     
        Elseif  LastDataRow - getPreviousSetsRow > 29 Then
            'find row of work order is not equal as getPreviousSetsRow to more 29 row from it
            For Each WOCurrentSet In Range("B" & getPreviousSetsRow & ":B" & getPreviousSetsRow+29)
                If WOCurrentSet.Value = Range("B" & getPreviousSetsRow+29).Value Then
                    SecondDataRowCurrent = WOCurrentSet.Row
                End If
            Next WOCurrentSet 
            
            if Range("B" & getPreviousSetsRow+29).Value = Range("B" & getPreviousSetsRow+30).Value Then

                Range("B" & getPreviousSetsRow & ":L" & getPreviousSetsRow+29).Copy
                Range("N62").Value = getPreviousSetsRow +29
                Range("O62").Value = "Case2.1"
            Else
                Range("B" & getPreviousSetsRow & ":L" & SecondDataRowCurrent-1).Copy
                Range("N62").Value = SecondDataRowCurrent-1
                Range("O62").Value = "Case2.2"
            end if              

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