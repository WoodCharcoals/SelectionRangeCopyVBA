Sub CPSet2_Click()
'dynamic copy selection range code
Dim LastDataRow, getPreviousSetsRow As Long
Dim WOLastRowOfSet, WOEqPreviousSet, WOInCurrentSet, WOInNextSet, WOInPreviousSet  As Range

Application.CutCopyMode = False

'Last row of selection's range in copy button set1
getPreviousSetsRow = Range("N33").Value + 1

With Range("L4:L204")

    LastDataRow = NotZero.Row - 1

    If Range("B33").Value = 0 Then
        Exit Sub
    ElseIf getPreviousSetsRow = 0 Then
        Exit Sub
    Else
        If LastDataRow - getPreviousSetsRow <= 29 Then
            Range("B" & getPreviousSetsRow & ":L" & LastDataRow).Copy
            Range("N62").Value = LastDataRow
     
        Elseif  LastDataRow - getPreviousSetsRow > 29 Then
            'find row of work order same as Row32 in Set1
            For Each CellValue In Range("B4:B31")
                If CellValue.Value = Range("B32").Value Then
                    FoundWO32 = CellValue.Row
                    Exit For
                End If
            Next CellValue      

            if Range("B" & getPreviousSetsRow+1 ).value then

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