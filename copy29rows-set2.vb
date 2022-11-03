Sub CPSet2_Click()
'dynamic copy selection range code
Dim LastDataRow, getPreviousSetsRow,SecondDataRowCurrent As Long
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
     
        Elseif  LastDataRow - getPreviousSetsRow > 29 Then
            'find row of work order is not equal as getPreviousSetsRow to more 29 row from it
            For Each WOCurrentSet In Range("B" & getPreviousSetsRow & ":B" & getPreviousSetsRow+28)
                If WOCurrentSet.Value <> Range("B" & getPreviousSetsRow).Value Then
                    SecondDataRowCurrent = WOCurrentSet.Row
                    Exit For
                Else
                    SecondDataRowCurrent = getPreviousSetsRow+28
                End If
            Next WOCurrentSet      

            Range("B" & getPreviousSetsRow & ":L" & SecondDataRowCurrent).Copy

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