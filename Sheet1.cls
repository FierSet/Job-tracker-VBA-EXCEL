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
Public DropdownStatus As Range


Sub start()
    Set DropdownStatus = Worksheets("Settings").Range("A2:A7")
    Set table_header = Worksheets("Settings").Range("A11:G11")
    
    table_header.Copy Destination:=Range("A2")
    table_header.Copy Destination:=Range("A10")
    
    Set inputRange = Range("A3:G3")
    insert_dropdown Range("F3")
    insert_dropdown Range("F11:F" & Rows.Count)
    Call moveshape("GetJob", Cells(10, "H").top + 2, Cells(10, "H").left + 2)
    inputRange.Value = Null
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Me.Range("F11:F" & Me.Rows.Count)) Is Nothing Or Not Intersect(Target, Me.Range("F3")) Is Nothing Then
        'Ignore if the user cleared multiple cells at once to prevent errors
        If Target.Cells.Count > 1 Then Exit Sub
        
        Dim formattedDate As String
        formattedDate = Format(Date, "dd/mm/yyyy")
        
        For Each cell In DropdownStatus
            If cell.Value = Target.Value Then
                Target.Interior.Color = cell.Interior.Color
                Cells(Target.Row, Target.Column - 2).Value = formattedDate
                Exit For
            Else
                If Target.Column = 6 And Target.Row = 3 Then
                    Target.Interior.Color = RGB(255, 255, 255)
                Else
                    Target.Interior.Color = Cells(1, 1).Interior.Color
                End If
            End If
        Next cell

        
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim myShape As Shape
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    If Target.Cells.Count > 1 Then Exit Sub
     
    If Not Intersect(Target, Me.Range("A11:G" & lastRow)) Is Nothing Then
        
        On Error Resume Next ' Prevents a crash if the shape name is wrong
        Set myShape = Me.Shapes("GetJob")
        On Error GoTo 0
        
         If Not myShape Is Nothing Then
            targetRow = Target.Row
            
            myShape.top = Me.Cells(targetRow, "H").top
            
            ' Lock the horizontal position to Column H
            myShape.left = Me.Cells(targetRow, "H").left
        End If
        
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
    Dim lastRow As Long
    Dim storegetable As String
    is_regete = False
    
    For Each cell In inputRange
        If IsEmpty(cell.Value) And Not Cells(cell.Row - 1, cell.Column).Value = "Note" Then
            MsgBox Cells(cell.Row - 1, cell.Column).Value & " is empty!"
            Exit Sub ' Stops checking once it finds the first empty cell
        End If
        
        If cell.Value = "Reject" Then
            is_regete = True
        End If
        
    Next cell
    
    storegetable = "A11"
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    If is_regete Then
        Set targetDest = Cells(lastRow + 1, "A")
        inputRange.Copy Destination:=targetDest
    Else
        Range(storegetable).EntireRow.insert Shift:=xlDown
        inputRange.Copy Destination:=Range(storegetable)
    End If

    inputRange.Value = Null
    inputRange.Interior.Color = RGB(255, 255, 255)
End Sub

Sub getjob()
    Dim currentRow As Range
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    Set currentRow = Cells(ActiveCell.Row, ActiveCell.Column)
    
    If Not Intersect(currentRow, Range("A11:G" & lastRow)) Is Nothing Then
    
        Call moveshape("GetJob", Cells(9, "H").top + 2, Cells(9, "H").left + 2)
        
        Range(Cells(currentRow.Row, "A"), Cells(currentRow.Row, "G")).Copy Destination:=Range("A3")
        Rows(currentRow.Row).EntireRow.Delete
    Else
        MsgBox "Please Select a Job!", vbExclamation
    End If
    
End Sub
