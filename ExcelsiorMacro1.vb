' This macro checks user's input for certain fixed columns and sets a predefined values in the Decision Column

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Student Type First Semester Freshman
    If Not Intersect(Target, Range("M:M")) Is Nothing Then
        For Each cell In Target
            If cell.Value = "6-First Term Freshman" Then
                Range("S" & cell.Row).Value = "1-Meets first-term freshman credit requirements"
            ElseIf cell.Value <> "6-First Term Freshman" Then
                Range("S" & cell.Row).ClearContents
            End If
        Next cell
    End If

    ' Break in Study
    If Not Intersect(Target, Range("O:O")) Is Nothing Then
        For Each cell In Target
            If cell.Value = "Y" Or cell.Value = "y" Then
                Range("S" & cell.Row).Value = "8-Student failed due to break in attendance"
            ElseIf cell.Value = "" Or cell.Value = "N" Or cell.Value = "n" Then
                Range("S" & cell.Row).ClearContents
            End If
        Next cell
    End If
   
    ' Credit Requirement Check
    If Not Intersect(Target, Range("Q2:Q" & Cells(Rows.Count, "Q").End(xlUp).Row)) Is Nothing Then
        For Each cell In Target
            Dim valueQ As Variant
            Dim valueP As Variant
            Dim valueO As Variant
            valueQ = cell.Value
            valueP = cell.Offset(0, -1).Value
            valueO = cell.Offset(0, -2).Value
            If IsNumeric(valueQ) And IsNumeric(valueP) And valueO = "N" Then
                'checks if Total Credit divide by Total Num of Terms more than or equal 15
                If valueQ / valueP >= 15 Then
                    Range("S" & cell.Row).Value = "2-Meets regular 2 or 4 year program requirements"
                'checks if Total Credit divide by Total Num of Terms less than 15
                ElseIf valueQ / valueP < 15 Then
                    Range("S" & cell.Row).Value = "7-Student failed due to insufficient credits"
                Else
                    Range("S" & cell.Row).ClearContents
                End If
            End If
        Next cell
    End If
End Sub
