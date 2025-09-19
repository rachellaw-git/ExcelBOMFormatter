Sub RemoveColumn()
    '================================================================
    ' MACRO: RemoveColumn
    ' PURPOSE: Copy the active sheet, remove a specified column from
    '          the copy, and optionally version the new sheet name.
    '================================================================

    '-------------------------
    ' VARIABLE DECLARATIONS
    '-------------------------
    Dim wb As Workbook             ' Active workbook reference
    Dim originalSheet As Worksheet ' Sheet to copy and modify
    Dim cadSheet As Worksheet      ' New sheet where column will be removed
    Dim columnToRemove As String   ' Column header text to remove
    Dim newSheetName As String     ' Base name for new sheet
    Dim foundCell As Range         ' Cell where column to remove is found
    Dim ws As Worksheet            ' Temporary worksheet reference
    Dim versionNumber As Integer   ' Counter for versioning sheets
    Dim finalSheetName As String   ' Final name assigned to new sheet
    Dim useVersioning As Boolean   ' Flag to enable versioned sheet names
    Dim activateOriginal As Boolean ' Flag to reactivate original sheet after processing

    '-------------------------
    ' USER CONFIGURATION
    '-------------------------
    columnToRemove = "BPP SKU"    ' Name of the column to remove
    newSheetName = "CAD"          ' Base name for the copied sheet
    useVersioning = False          ' Enable versioned sheet names if base exists
    activateOriginal = True        ' Return focus to original sheet after processing

    On Error GoTo ErrorHandler

    '-------------------------
    ' VALIDATE WORKBOOK
    '-------------------------
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found. Please open the workbook you wish to process.", vbExclamation
        Exit Sub
    End If
    Set wb = ActiveWorkbook
    Set originalSheet = wb.ActiveSheet   ' Work on the currently active sheet

    '-------------------------
    ' FIND COLUMN TO REMOVE
    '-------------------------
    Set foundCell = originalSheet.Cells.Find(What:=columnToRemove, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If foundCell Is Nothing Then
        MsgBox "The column '" & columnToRemove & "' was not found. Process aborted.", vbCritical
        Exit Sub
    End If

    '-------------------------
    ' DETERMINE FINAL SHEET NAME
    '-------------------------
    If useVersioning Then
        ' Generate a unique versioned name if base exists
        If Not SheetExists(wb, newSheetName) Then
            finalSheetName = newSheetName
        Else
            versionNumber = 1
            finalSheetName = newSheetName & "(" & versionNumber & ")"
            Do While SheetExists(wb, finalSheetName)
                versionNumber = versionNumber + 1
                finalSheetName = newSheetName & "(" & versionNumber & ")"
            Loop
        End If
    Else
        ' Remove any existing sheet with the same name
        On Error Resume Next
        Set ws = wb.Sheets(newSheetName)
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
        End If
        On Error GoTo ErrorHandler
        finalSheetName = newSheetName
    End If

    '-------------------------
    ' COPY SHEET AND REMOVE COLUMN
    '-------------------------
    originalSheet.Copy After:=originalSheet
    Set cadSheet = wb.ActiveSheet

    Set foundCell = cadSheet.Cells.Find(What:=columnToRemove, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not foundCell Is Nothing Then foundCell.EntireColumn.Delete

    '-------------------------
    ' RENAME AND CLEANUP
    '-------------------------
    cadSheet.Name = finalSheetName

    If activateOriginal Then originalSheet.Activate

    Exit Sub

'================================================================
' ERROR HANDLER
'================================================================
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Application.DisplayAlerts = True
End Sub

'================================================================
' HELPER FUNCTION: SheetExists
' PURPOSE: Checks if a sheet with a specified name exists in a workbook
'================================================================
Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

