Sub CPSet1_Click()
Dim LastDataRow, SameWO As Long
Dim FoundWO32, W32InSet2Row, CellValue, CellValue2  As Range


Application.CutCopyMode = False

With Range("L4:L206")

    LastDataRow = NotZero.Row - 1
    
If Range("B4").Value = 0 Then
    CPSet1.BackColor = RGB(255, 255, 255)
    Exit Sub
    
ElseIf LastDataRow < 33 Then

    Range("B4:L" & LastDataRow).Copy
    Range("N33").Value = LastDataRow
    
Else
    'find row of work order same as Row32 in Set1
    For Each CellValue In Range("B4:B31")
        If CellValue.Value = Range("B32").Value Then
            FoundWO32 = CellValue.Row
            Exit For
        End If
    Next CellValue

    'find row of work order not same as Row32 in Set2
    For Each CellValue2 In Range("B33:B61")
        If CellValue2.Value <> Range("B32").Value Then
            W32InSet2Row = CellValue2.Row
            Exit For
        End If
    Next CellValue2

    If W32InSet2Row - 32 > 1 Then

        Range("B4:L" & FoundWO32 - 1).Copy
        Range("N33").Value = FoundWO32 - 1

    End If
    
End If

    If Application.CutCopyMode = xlCopy Then
        'Debug.Print "Copy status:=Done"
        Me.CPSet1.BackColor = vbGreen 'RGB(50, 205, 50)
    Else
        'Debug.Print "Copy status:=Failed"
        Me.CPSet1.BackColor = vbRed 'RGB(255, 0, 0)
    End If
    
End With

End Sub