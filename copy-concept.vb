'vb 

Private Sub CPSet2_Click()
'Status = Done maybe
'for ROW if last max=32,current=61
Dim LastCopiedRow, FirstRowSet2, SetTwo1stValue, SameWO, lastUsed, Diff2 As Long
Dim Condition1, Found As Long
Dim Found2, Found3, Found4, Found5, _
CellValue, CellValue2, CellValue3, CellValue5, WoInSet2 As Range

Application.CutCopyMode = False
'-----------------------------------------
'Determine constant values
LastCopiedRow = Range("n24").Value + 1
'show 1st row of new copying set
FirstRowSet2 = Range("B" & LastCopiedRow).Row
SetTwo1stValue = Range("B" & FirstRowSet2).Value

'starting of main code
If Range("B24").Value = 0 Then
    'set red bg at set2
    Me.CPSet2.BackColor = RGB(255, 95, 31)
    Exit Sub
ElseIf Range("B33").Value = 0 Then
    Exit Sub
Else
    '1#.Find Row number of WO
    With Range("L4:L204")
        'Find last row with Work Order number
        lastUsed = NotZero.Row 
    'Determine condition for copy range    
    Condition1 = LastUsed -  FirstRowSet2

    '2# start copy with 1st condition
        'verify all row > 29 rows
        If LastUsed < FirstRowSet2 Then
            Exit Sub
        'lastUsed >=24|<= 43 or LastUsed-FirstRowSet2 <=29
        ElseIf Condition1 <= 29 Then
            Range("B" & FirstRowSet2 & ":L" & LastUsed).Copy
        
        ElseIf Condition1 > 29 Then

            For Each LastRowOfLastOrderSet2 in Range("B44:B63")
                If LastRowOfLastOrderSet2.Value <> range("B43").value Then
                LastOrderRow = LastRowOfLastOrderSet2.Row
                Exit for
            Next LastRowOfLastOrderSet2

            If LastOrderRow - FirstRowSet2 <= 29 Then
                Range("B" & FirstRowSet2 & ":L" & LastOrderRow).Copy

            Else
                For Each OrderInSet2 in Range("B24:B42")
                    If OrderInSet2.Value <> range("B43").value Then
                    OrderInSet2Row = OrderInSet2.Row
                    Exit for
                Next OrderInSet2

                Range("B" & FirstRowSet2 & ":L" & OrderInSet2Row).Copy

            End if
        elseif range("B23").value = range("B43").vale Then 
            For Each WoInSet1 in Range("B4:B23")
                If WoInSet1.Value <> range("B23").value Then
                WoSet1EqW43Row = WoInSet1.Row
                Exit for
            Next WoInSet1
            
            if WoInSet1 = 4 Then
                range("B24:L43").copy
            elseif
                range("B" & WoSet1EqW43Row+1 & ":L43").copy
            end if

        End if

    End with
'-------------------------------------------------------------------
'---------- OLD CODE  ----------------------------------------------

    '3. Find last row with relate to set3
    For Each CellValue2 In Range("B" & FirstRowSet2 & ":B63")
        If CellValue2.Value <> SetTwo1stValue Then
            Found2 = CellValue2.Row
            Debug.Print "NewW/O:" & Found2 & ", Value@1stRow:" & SetTwo1stValue
            Exit For
        End If
    Next CellValue2

    '4. Find new WO relate to F2
    For Each CellValue3 In Range(CellValue2.Address & ":B63")
            If Found2 = Found Then
                Found3 = CellValue3.Row
                ''Debug.Print "Found3 New W/O at Row: " & Found3
                Exit For
            ElseIf CellValue3.Value <> CellValue2.Value Then
                Found3 = CellValue3.Row - 1
                'Debug.Print "Found3 New W/O at Row: " & Found3
                Exit For
            Else
                Found3 = Range("B63").Row
            End If
    Next CellValue3
            'Debug.Print "CVv3: " & CellValue3.Value & ", F3:" & Found3
            
    '5. find new-W/O's row
    For Each WoInSet2 In Range("B" & FirstRowSet2 & ":B43")
            If Range("B43").Value = WoInSet2.Value Then
                Found4 = WoInSet2.Row
                'Debug.Print "Found4 New W/O at Row: " & Found4
                Exit For
            End If
    Next WoInSet2
    
    '6. find maximun of row count
    For Each CellValue5 In Range("B44:B63")
        If CellValue5.Value <> Range("B43").Value Then
            Found5 = CellValue5.Row
            'Debug.Print "Found4 New W/O at Row: " & Found4
            Exit For
        End If
    Next CellValue5
    
    Diff2 = Found5 - FirstRowSet2
    Debug.Print "Dff2:" & Diff2 & ", Start/WO1:" & FirstRowSet2
    Debug.Print "WO2:" & Found2 & ", EndWO2:" & Found3
    Debug.Print "WOinSet2FromStart:" & Found4 & ", WOinSet3:" & Found5
    Debug.Print "Diff WOInSet3ToStart:" & Found5 - FirstRowSet2
     
    If Diff2 <= 29 Then
        Range("B" & FirstRowSet2 & ":L" & Found5 - 1).Copy
        Range("N44").Value = Found5 - 1
        Debug.Print "case 1"
        Debug.Print
    ElseIf Diff2 > 29 Then
        Range("B" & FirstRowSet2 & ":L" & Found4 - 1).Copy
        Range("N44").Value = Found4 - 1
        Debug.Print "case2"
    End If
    
    End With
    
End If

    If Application.CutCopyMode = xlCopy Then
        'Debug.Print "Copy status:=Done"
        Me.CPSet2.BackColor = vbGreen 'RGB(50, 205, 50)
    Else
        'Debug.Print "Copy status:=Failed"
        Me.CPSet2.BackColor = vbRed 'RGB(255, 0, 0)
    
    End If

End Sub