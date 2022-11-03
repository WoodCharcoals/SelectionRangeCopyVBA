Sub CPSet2_Click()
'dynamic copy selection range code
Dim LastDataRow, PrevSetsRow,SecondDataRowCurrent,FindDataNextSet, DataSet As Long
Dim WOLastRowOfSet, WOEqPrevSet, WOCurrentSet, WONextSet, WOInPrevSet  As Range

Application.CutCopyMode = False

'Last row of selection's range in copy button set1
PrevSetsRow = Range("N33").Value + 1
DataSet = PrevSetsRow + 29
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
     
        Elseif  LastDataRow - PrevSetsRow > 29 Then
            'find row of work order is not equal as PrevSetsRow to more 29 row from it
            For Each WOCurrentSet In Range("B" & PrevSetsRow & ":B" & Dataset)
                If WOCurrentSet.Value = Range("B" & DataSet).Value Then
                    SecondDataRowCurrent = WOCurrentSet.Row
                End If
            Next WOCurrentSet
            'find row of current val in next data set
            for each WONextSet in range ("B" & Dataset &":B" & dataset+29)
                if wonextset.value <> range("B" & dataset).value then 
                    FindDataNextSet = WONextSet.row -1
                else    
                    FindDataNextSet = dateset+29
                end if
            next WONextSet
            
            'return selection range 
            if FindDataNextSet - SecondDataRowCurrent > 29

                Range("B" & PrevSetsRow & ":L" & SecondDataRowCurrent).Copy
                Range("N62").Value = PrevSetsRow +29
                Range("O62").Value = "Case2.3" 

            ElseIf Range("B" & DataSet).Value = Range("B" & PrevSetsRow+30).Value Then

                Range("B" & PrevSetsRow & ":L" & DataSet).Copy
                Range("N62").Value = PrevSetsRow +29
                Range("O62").Value = "Case2.1"
            Elseif
                Range("B" & PrevSetsRow & ":L" & SecondDataRowCurrent-1).Copy
                Range("N62").Value = SecondDataRowCurrent-1
                Range("O62").Value = "Case2.2"
            else

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