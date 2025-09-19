Sub FormatBomForCadDatalink()
    '================================================================
    ' MACRO: FormatBomForCadDatalink_Optimized
    ' PURPOSE: Copy the active sheet, remove a specified column,
    '          and format a BOM for CAD Datalink.
    '          - Cells containing "BILL OF MATERIAL" or starting with "SECTION" are centered.
    '          - Column alignment is applied based on user-configured headers.
    '          - Fonts, sizes, and boldness are applied according to content.
    '================================================================

    '================================================================
    ' USER CONFIGURATION - modify these to adjust behavior
    '================================================================
    Dim newSheetName As String: newSheetName = "CAD"
        ' Name of the new sheet to create for the CAD BOM
    Dim sheetSuffix As String: sheetSuffix = "Procurement"
        ' Suffix added to the original sheet when renaming it
    Dim columnToRemove As String: columnToRemove = "BPP SKU"
        ' Column to delete from the copied sheet
    Dim activateOriginal As Boolean: activateOriginal = True
        ' If True, original sheet will be reactivated after processing
    Dim useVersioning As Boolean: useVersioning = False
        ' If True, appends version number if the target sheet already exists

    Dim billOfMaterialText As String: billOfMaterialText = "BILL OF MATERIAL"
        ' Text used to identify the main BOM header cell
    Dim sectionText As String: sectionText = "SECTION"
        ' Text used to identify section header cells
    Dim tableHeaders As Variant: tableHeaders = Array("ITEM#", "QTY", "MFR PART #", "MANUFACTURER", "DESCRIPTION")
        ' Table column headers; used for formatting header cells

    Dim centerAlignedColumns As Variant: centerAlignedColumns = Array("ITEM#", "QTY", "MFR PART #")
        ' Columns to be horizontally centered
    Dim leftAlignedColumns As Variant: leftAlignedColumns = Array("MANUFACTURER", "DESCRIPTION")
        ' Columns to be horizontally left-aligned

    '================================================================
    ' INTERNAL VARIABLES
    '================================================================
    Dim wb As Workbook, originalSheet As Worksheet, cadSheet As Worksheet
    Dim finalSheetName As String, baseFileName As String, versionNumber As Long
    Dim cell As Range, cleanedText As String
    Dim centerDict As Object, leftDict As Object
    Dim centerCols As Object, leftCols As Object
    Dim h As Variant, col As Variant
    Dim foundCell As Range

    On Error GoTo ErrorHandler

    '-------------------------
    ' INITIALIZE DICTIONARIES
    '-------------------------
    Set centerDict = CreateObject("Scripting.Dictionary")
    Set leftDict = CreateObject("Scripting.Dictionary")
    Set centerCols = CreateObject("Scripting.Dictionary")
    Set leftCols = CreateObject("Scripting.Dictionary")

    ' Populate dictionaries for quick lookup of column headers
    For Each h In centerAlignedColumns: centerDict(UCase(h)) = True: Next
    For Each h In leftAlignedColumns: leftDict(UCase(h)) = True: Next

    '-------------------------
    ' VALIDATE ACTIVE WORKBOOK
    '-------------------------
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found. Please open the workbook you wish to process.", vbExclamation
        Exit Sub
    End If
    Set wb = ActiveWorkbook
    Set originalSheet = wb.ActiveSheet

    '-------------------------
    ' LOCATE COLUMN TO REMOVE
    '-------------------------
    Set foundCell = originalSheet.Cells.Find(What:=columnToRemove, LookIn:=xlValues, _
                                             LookAt:=xlWhole, MatchCase:=False)
    If foundCell Is Nothing Then
        MsgBox "The column '" & columnToRemove & "' was not found. Process aborted.", vbCritical
        Exit Sub
    End If

    '-------------------------
    ' RENAME ORIGINAL SHEET
    '-------------------------
    baseFileName = wb.Name
    If InStrRev(baseFileName, ".") > 0 Then
        baseFileName = Left(baseFileName, InStrRev(baseFileName, ".") - 1)
    End If
    originalSheet.Name = baseFileName & " " & sheetSuffix

    '-------------------------
    ' DETERMINE NEW SHEET NAME
    '-------------------------
    If useVersioning Then
        ' Append version number if a sheet with the same name exists
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
        ' Delete existing sheet if versioning not used
        On Error Resume Next
        Application.DisplayAlerts = False
        wb.Sheets(newSheetName).Delete
        Application.DisplayAlerts = True
        On Error GoTo ErrorHandler
        finalSheetName = newSheetName
    End If

    '-------------------------
    ' COPY ORIGINAL SHEET AND REMOVE COLUMN
    '-------------------------
    originalSheet.Copy After:=originalSheet
    Set cadSheet = wb.ActiveSheet

    Set foundCell = cadSheet.Cells.Find(What:=columnToRemove, LookIn:=xlValues, _
                                        LookAt:=xlWhole, MatchCase:=False)
    If Not foundCell Is Nothing Then foundCell.EntireColumn.Delete

    '-------------------------
    ' FORMAT CELLS AND TRACK COLUMN ALIGNMENTS
    '-------------------------
    With cadSheet.UsedRange
        For Each cell In .Cells
            ' Convert all text to uppercase for consistency
            If VarType(cell.Value) = vbString Then cell.Value = UCase(cell.Value)
            cell.Font.Name = "Arial"
            cleanedText = Trim(cell.Value)

            ' Apply font size, bold, and horizontal alignment based on content
            Select Case True
                Case StrComp(cleanedText, billOfMaterialText, vbTextCompare) = 0
                    cell.Font.Size = 16
                    cell.Font.Bold = True
                    cell.HorizontalAlignment = xlCenter
                Case UCase(Left(cleanedText, Len(sectionText))) = UCase(sectionText)
                    cell.Font.Size = 14
                    cell.Font.Bold = True
                    cell.HorizontalAlignment = xlCenter
                Case IsInArray(cleanedText, tableHeaders)
                    cell.Font.Size = 11
                    cell.Font.Bold = True
                Case Else
                    cell.Font.Size = 11
                    cell.Font.Bold = False
            End Select

            ' Default vertical alignment for all cells
            cell.VerticalAlignment = xlCenter

            ' Track which columns need center or left alignment
            If centerDict.Exists(cleanedText) Then centerCols(cell.Column) = True
            If leftDict.Exists(cleanedText) Then leftCols(cell.Column) = True
        Next cell
    End With

    '-------------------------
    ' APPLY HORIZONTAL ALIGNMENTS TO COLUMNS
    '-------------------------
    For Each col In centerCols.Keys: cadSheet.Columns(col).HorizontalAlignment = xlCenter: Next
    For Each col In leftCols.Keys: cadSheet.Columns(col).HorizontalAlignment = xlLeft: Next

    '-------------------------
    ' FINALIZE SHEET
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
' Checks whether a sheet exists in the workbook
'================================================================
Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

'================================================================
' HELPER FUNCTION: IsInArray
' Checks if a value exists in an array (case-insensitive)
'================================================================
Function IsInArray(val As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If StrComp(arr(i), val, vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function
