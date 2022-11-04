
Option Explicit

Private Sub Worksheet_Activate()
Dim LastRowSet As Long
With Me

    .CPSet1.BackColor = RGB(255, 255, 255)
    .CPSet2.BackColor = RGB(255, 255, 255)
    .CPSet3.BackColor = RGB(255, 255, 255)
    .CPSet4.BackColor = RGB(255, 255, 255)
    
End With

'For LastRowSet = 0 To 6
'    Range("N" & 33 + (29 * LastRowSet)).Value = "0"
'
'Next LastRowSet

End Sub

Public Property Get Ws() As Worksheet
    Set Ws = ActiveWorkbook.Worksheets("To IW44")
End Property

'Status = Done
Public Property Get NotZero() As Range
    Set NotZero = Range("L4:L206").Find("0", LookIn:=xlValues, _
    MatchCase:=True, SearchDirection:=xlNext)
    
End Property

Sub CPSet1_Click()
Dim LastDataRow, PrevSetsRow, SecondDataRowCurrent, FindDataNextSet, DataSet As Long
Dim CurrentDataRange, NextDataRange, WOCurrentSet, WONextSet, SetResultRow  As Range



End Sub


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

    If PrevSetsRow < 4 Then
        SetResultRow.Value = "0"
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
            If FindDataNextSet - PrevSetsRow > 29 Then

                Range("B" & PrevSetsRow & ":L" & DataSet).Copy
                SetResultRow.Value = DataSet
                SetResultRow.Offset(0, 1).Value = "Case2.3"
                
           ElseIf Range("B" & DataSet).Value = Range("B" & DataSet + 1).Value Then

                Range("B" & PrevSetsRow & ":L" & SecondDataRowCurrent).Copy
                Range("B" & PrevSetsRow).Select
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

End Sub

Sub CPSet3_Click()

End Sub

Sub CPSet4_Click()

End Sub


Private Sub CPSet5_Click()

End Sub
Private Sub CPSet6_Click()

End Sub
Private Sub CPSet7_Click()

End Sub
Private Sub CPSet8_Click()

End Sub
Private Sub CPSet9_Click()

End Sub
Private Sub CPSet10_Click()

End Sub
'
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'
'
'
'    If Application.CutCopyMode = xlCopy Then
'        ''Debug.Print "Copy status:=Done"
'        Me.CPSet2.BackColor = vbGreen 'RGB(50, 205, 50)
'
'    ElseIf Application.CutCopyMode = False Then
'        ''Debug.Print "Copy status:=Failed"
'        Me.CPSet2.BackColor = vbWhite 'RGB(255, 0, 0)
'    Else
'        Me.CPSet2.BackColor = vbRed 'RGB(50, 205, 50)
'    End If
'
'
'
'End Sub
