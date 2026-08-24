VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public inputRange As Range
Public table_header As Range
Public table_analytisis As Range
Public DropdownStatus As Range


Sub start()
    Set DropdownStatus = Worksheets("Settings").Range("A2:A7")
    Set table_header = Worksheets("Settings").Range("A11:G11")
    Set table_analytisis = Worksheets("Settings").Range("A15:F32")
    
    table_header.Copy Destination:=Range("A2")
    table_header.Copy Destination:=Range("A10")
    
    Set inputRange = Range("A3:G3")
    insert_dropdown Range("F3")
    insert_dropdown Range("F11:F" & Rows.Count)
    
    Call moveshape("GetJob", Cells(9, "I").top + 2, Cells(9, "I").left + 2)
    Call moveshape("Deletejob", Cells(9, "J").top + 2, Cells(9, "J").left + 2)
    inputRange.Value = Null
End Sub

Private Sub Worksheet_Change(ByVal target As Range)
    If Not Intersect(target, Me.Range("F11:F" & Me.Rows.Count)) Is Nothing Or Not Intersect(target, Me.Range("F3")) Is Nothing Then
        'Ignore if the user cleared multiple cells at once to prevent errors
        If target.Cells.Count > 1 Then Exit Sub
        
        Dim formattedDate As String
        formattedDate = Format(Date, "dd/mm/yyyy")
        
        For Each cell In DropdownStatus
            If cell.Value = target.Value Then
                target.Interior.Color = cell.Interior.Color
                Cells(target.Row, "D").Value = formattedDate
                analytics target
                Exit For
            Else
                If target.Column = 6 And target.Row = 3 Then
                    target.Interior.Color = RGB(255, 255, 255)
                Else
                    target.Interior.Color = Cells(1, 1).Interior.Color
                End If
            End If
        Next cell
        
        If Not Intersect(target, Me.Range("F11:F" & Me.Rows.Count)) Is Nothing Then
            If Cells(target.Row, "F").Value = "Reject" Or Cells(target.Row, "F").Value = "Not Receive Updarted" Then
                Move_To_End target
            Else
                Move_To_Top Range(Cells(target.Row, "A"), Cells(target.Row, "H"))
            End If
        End If
        
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    Dim myShape As Shape
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    If target.Cells.Count > 1 Then Exit Sub
     
    If Not Intersect(target, Me.Range("A11:H" & lastRow)) Is Nothing Then
        targetRow = target.Row
        Call moveshape("GetJob", Me.Cells(targetRow, "I").top, Me.Cells(targetRow, "I").left)
        Call moveshape("Deletejob", Me.Cells(targetRow, "J").top, Me.Cells(targetRow, "J").left)
    End If
    
End Sub

Sub Move_To_End(target As Variant)
    Dim lastRow As Long
    Call moveshape("GetJob", Cells(9, "I").top + 2, Cells(9, "I").left + 2)
    Call moveshape("Deletejob", Cells(9, "J").top + 2, Cells(9, "J").left + 2)
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Set targetDest = Cells(lastRow + 1, "A")
    Range(Cells(target.Row, "A"), Cells(target.Row, "H")).Copy Destination:=targetDest
    analytics Range("F" & targetDest.Row)
    If target.Row = 3 Then
        Cells(target.Row, "G").Value = Null
    Else
        Cells(target.Row, "H").EntireRow.Delete
    End If
    
End Sub

Sub Move_To_Top(target As Range)
    Dim storegetable As String
    storegetable = "A11"
    Call moveshape("GetJob", Cells(9, "I").top + 2, Cells(9, "I").left + 2)
    Call moveshape("Deletejob", Cells(9, "J").top + 2, Cells(9, "J").left + 2)
    Range(storegetable).EntireRow.insert Shift:=xlDown
    target.Copy Destination:=Range(storegetable)
    
    analytics Range("F11")
    
    If target.Row = 3 Then
        Cells(target.Row, "G").Value = Null
    Else
        Cells(target.Row, "H").EntireRow.Delete
    End If
    
End Sub

Sub moveshape(Shapename As String, top As Variant, left As Variant)
    Dim myShape As Shape
    
    On Error Resume Next
    Set myShape = Shapes(Shapename)
    On Error GoTo 0
    
    If Not myShape Is Nothing Then
        myShape.top = top
        myShape.left = left
    End If
    
End Sub

Sub cleanimputrange()
    inputRange.Value = Null
End Sub

Sub insert_dropdown(Ranges As Range)
    
    Ranges.Validation.Delete
    Ranges.Validation.Add Type:=xlValidateList, _
                              AlertStyle:=xlValidAlertStop, _
                              Formula1:="=" & DropdownStatus.Address(External:=True)
End Sub

Sub SaveJob()
    Dim is_regete As Boolean
    is_regete = False
    
    For Each cell In inputRange
        If IsEmpty(cell.Value) And Not Cells(cell.Row - 1, cell.Column).Value = "Note" Then
            MsgBox Cells(cell.Row - 1, cell.Column).Value & " is empty!"
            Exit Sub ' Stops checking once it finds the first empty cell
        End If
        
        If cell.Value = "Reject" Or cell.Value = "Not Receive Updarted" Then
            is_regete = True
        End If
        
    Next cell
    
    If is_regete Then
         Move_To_End inputRange
    Else
        Move_To_Top inputRange
    End If

    inputRange.Value = Null
    inputRange.Interior.Color = RGB(255, 255, 255)
End Sub

Sub Delet_job()
    Dim currentRow As Range
    Dim lastRow As Long
    Dim userResponse As VbMsgBoxResult
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Set currentRow = Cells(ActiveCell.Row, ActiveCell.Column)
    
    If Not Intersect(currentRow, Range("A11:H" & lastRow)) Is Nothing Then
    
        userResponse = MsgBox("Do you want to Delete the Job:" & _
        vbCrLf & Cells(currentRow.Row, "B") & " of " & Cells(currentRow.Row, "A") & "?", vbYesNo + vbQuestion, "Confirmation Required")
        
        If userResponse = vbNo Then
            Exit Sub
        End If
    
        Call moveshape("GetJob", Cells(9, "I").top + 2, Cells(9, "I").left + 2)
        Call moveshape("Deletejob", Cells(9, "J").top + 2, Cells(9, "J").left + 2)
        Rows(currentRow.Row).EntireRow.Delete
        
        MsgBox "Job deleted!", vbExclamation
    Else
        MsgBox "Please Select a Job!", vbExclamation
    End If
End Sub

Sub getjob()
    Dim currentRow As Range
    Dim lastRow As Long
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Set currentRow = Cells(ActiveCell.Row, ActiveCell.Column)
    
    If Not Intersect(currentRow, Range("A11:G" & lastRow)) Is Nothing Then
    
        Call moveshape("GetJob", Cells(9, "I").top + 2, Cells(9, "I").left + 2)
        Call moveshape("Deletejob", Cells(9, "J").top + 2, Cells(9, "J").left + 2)
        Range(Cells(currentRow.Row, "A"), Cells(currentRow.Row, "G")).Copy Destination:=Range("A3")
        Rows(currentRow.Row).EntireRow.Delete
    Else
        MsgBox "Please Select a Job!", vbExclamation
    End If
    
End Sub

Sub analytics(target As Range)
    Dim formattedDate As String
    formattedDate = Format(Date, "dd/mm/yyyy")
        
    For Each cell In table_analytisis.Columns(1).Cells
        If target.Value = cell.Value Then
            If IsEmpty(table_analytisis.Cells(cell.Row, "C").Value) And IsEmpty(table_analytisis.Cells(cell.Row, "E").Value) Then
                Cells(target.Row, "H").Interior.Color = cell.Offset(0, 5).Interior.Color
                Cells(target.Row, "H").Value = cell.Offset(0, 5).Value
                Exit For
            ElseIf Not IsEmpty(table_analytisis.Cells(cell.Row, "C").Value) And IsEmpty(table_analytisis.Cells(cell.Row, "E").Value) Then
                If formattedDate - Cells(target.Row, "D").Value > table_analytisis.Cells(Cells.Row, "C").Value Then
                    Cells(target.Row, "H").Interior.Color = cell.Offset(0, 5).Interior.Color
                    Cells(target.Row, "H").Value = cell.Offset(0, 5).Value
                    Exit For
                End If
            ElseIf Not IsEmpty(table_analytisis.Cells(cell.Row, "C").Value) And Not IsEmpty(table_analytisis.Cells(cell.Row, "E").Value) Then
                If formattedDate - Cells(target.Row, "D").Value > table_analytisis.Cells(Cells.Row, "C").Value And formattedDate - Cells(target.Row, "D").Value < table_analytisis.Cells(Cells.Row, "E").Value Then
                    Cells(target.Row, "H").Interior.Color = cell.Offset(0, 5).Interior.Color
                    Cells(target.Row, "H").Value = cell.Offset(0, 5).Value
                    Exit For
                End If
            End If
            
        End If
    Next cell
    
End Sub

Sub massive_analysis()
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For Each cell In Range("A11:A" & lastRow)
        analytics Range("F" & cell.Row)
    Next cell
    
End Sub
