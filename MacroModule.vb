'This macro module serves to automate the process of filtering out pending data in Excel. When the user presses a button represented by a created Shape, it performs the following tasks:
'Filters the pending data and creates a new file called "Pending Verification" to store the filtered data.
'Creates another new file called "Data to Send" for submission. This file retains only the required columns while preserving all the formatting.
'In summary, the macro filters pending data, saves it in a file called "Pending Verification," and creates a submission file named "Data to Send" with the necessary columns and formatting intact.

Sub ExportData_Click()
    Application.ScreenUpdating = False ' Turn off screen updating to speed up the process
    
    ' Define the source data range (all columns to be checked)
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Worksheets("Data Sheet")
    
    ' Create a new workbook and define the destination sheets ("Data to Send" and "Pending Verifications")
    Dim destBook As Workbook
    Set destBook = Workbooks.Add
    Dim destSheet As Worksheet
    Set destSheet = destBook.Worksheets(1)
    destSheet.Name = "Data to send"
    
    Dim pendingBook As Workbook
    Set pendingBook = Workbooks.Add
    Dim pendingSheet As Worksheet
    Set pendingSheet = pendingBook.Worksheets(1)
    pendingSheet.Name = "Pending Verifications"
    
    
    ' Copy entire source sheet to destination sheet
    sourceSheet.Cells.Copy
    destSheet.Cells.PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    
    ' Define the columns to be kept in "Data to Send" sheet
    Dim columnsToKeep As Variant
    columnsToKeep = Array("Student ID", "NAME LAST STDNT", "NAME FIRST STDN", "MIDDLE INITIAL", "HESC DOB", "Federal Code", _
                          "Branch", "HESC Inst Code", "HESC Acad Year", "Applied term", "Student Type", "INFO-PENDING", _
                          "TOTAL-NUM-TERMS", "TOTAL CREDITS-EARNED", "MEETS or FAILS CREDIT-REQUIREM", "Institution", "Max Inbound date")

    ' Initialize the infoPendingIndex variable
    Dim infoPendingIndex As Long
    infoPendingIndex = Application.Match("INFO-PENDING", destSheet.Rows(1), 0)

    ' Copy header row from destination sheet to pending sheet
    destSheet.Rows(1).Copy
    pendingSheet.Rows(1).PasteSpecial Paste:=xlPasteAllUsingSourceTheme
    
    ' Initialize the pendingRow variable
    Dim pendingRow As Long
    pendingRow = 2

    ' Delete rows that have INFO-PENDING = Y from destination sheet and copy them to pending sheet.
    
    Dim lastRow As Long
    lastRow = destSheet.Cells(destSheet.Rows.Count, infoPendingIndex).End(xlUp).Row
    
    Dim rowNum As Long
    For rowNum = lastRow To 2 Step -1 ' Start from last row and go up to row 2 (to skip header row).
        If destSheet.Cells(rowNum, infoPendingIndex).Value = "Y" Then
            
            ' Copy entire row from destination sheet to pending sheet.
            destSheet.Rows(rowNum).Copy
            
            With pendingSheet.Rows(pendingRow)
                .PasteSpecial Paste:=xlPasteAllUsingSourceTheme
                
                pendingRow = pendingRow + 1 ' Move to the next row in the pending sheet.
            End With
            
            destSheet.Rows(rowNum).Delete ' Delete row from destination sheet.
        End If
    Next rowNum
    
    ' Delete columns that are not specified in columnsToKeep array from destination sheet only.
    
    Dim lastCol As Long
    
    lastCol = destSheet.Cells(1, destSheet.Columns.Count).End(xlToLeft).Column
    
    Dim colNum As Long
    For colNum = lastCol To 1 Step -1 ' Start from last column and go up to first column.
        If IsError(Application.Match(destSheet.Cells(1, colNum).Value, columnsToKeep, 0)) Then
            destSheet.Columns(colNum).Delete
        End If
    Next colNum
    
Application.CutCopyMode = False ' Clear the clipboard.
Application.ScreenUpdating = True ' Turn on screen updating again.
MsgBox "Data exported to 'Data to send' and 'Pending Verifications' sheets in new workbook.", vbInformation

End Sub
