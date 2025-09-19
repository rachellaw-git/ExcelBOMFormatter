Sub FormatBom()
    ' Formats the active sheet and creates a CAD-ready copy.
    ' - Original sheet is formatted and optionally renamed
    ' - CAD sheet is created (uppercase + optional column removed)
    ' - Existing CAD sheet is deleted silently if present

    '=======================
    ' USER CONFIGURATION
    '=======================
    Dim newSheetName As String: newSheetName = "CAD"
    Dim sheetSuffix As String: sheetSuffix = ""
    Dim columnToRemove As String: columnToRemove = "BPP SKU"
    Dim activateOriginal As Boolean: activateOriginal = True

    Dim billOfMaterialText As String: billOfMaterialText = "BILL OF MATERIAL"
    Dim sectionText As String: sectionText = "SECTION"
    Dim tableHeaders As Variant: tableHeaders = Array("ITEM#", "QTY", "BPP SKU", "MFR PART #", "MANUFACTURER", "DESCRIPTION")

    Dim centerAlignedColumns As Variant: centerAlignedColumns = Array("ITEM#", "QTY", "BPP SKU")
    Dim leftAlignedColumns As Variant: leftAlignedColumns = Array("MFR PART #", "MANUFACTURER", "DESCRIPTION")

    '=======================
    ' INTERNAL VARIABLES
    '=======================
    Dim wb As Workbook, originalSheet As Worksheet, cadSheet As Worksheet
    Dim finalSheetName As String, baseFileName As String

    On Error GoTo ErrorHandler

    ' Validate workbook
    If ActiveWorkbook Is Nothing Then
        MsgBox "No active workbook found.", vbExclamation
        Exit Sub
    End If
    Set wb = ActiveWorkbook
    Set originalSheet = wb.ActiveSheet

    ' Format the original sheet
    FormatOriginalSheet originalSheet, billOfMaterialText, sectionText, tableHeaders, centerAlignedColumns, leftAlignedColumns

    ' Determine new name for original sheet
    baseFileName = wb.Name
    If InStrRev(baseFileName, ".") > 0 Then baseFileName = Left(baseFileName, InStrRev(baseFileName, ".") - 1)
    If sheetSuffix = "" Then
        finalSheetName = baseFileName
    Else
        finalSheetName = baseFileName & " " & sheetSuffix
    End If

    ' Rename original sheet if needed
    If originalSheet.Name <> finalSheetName Then originalSheet.Name = finalSheetName

    ' Copy original sheet for CAD
    originalSheet.Copy After:=originalSheet
    Set cadSheet = wb.ActiveSheet

    ' Prepare CAD sheet (uppercase + remove specified column)
    PrepareCadSheet cadSheet, columnToRemove

    ' Delete existing CAD sheet silently if present, then rename
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(newSheetName)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo ErrorHandler
    cadSheet.Name = newSheetName

    ' Optionally activate the original sheet
    If activateOriginal Then originalSheet.Activate

    Exit Sub

'=======================
' ERROR HANDLER
'=======================
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

'=======================
' FORMAT ORIGINAL SHEET
'=======================
Sub FormatOriginalSheet(ws As Worksheet, billText As String, sectionText As String, _
                        tableHeaders As Variant, centerColsArr As Variant, leftColsArr As Variant)
    Dim cell As Range, cleanedText As String
    Dim centerDict As Object, leftDict As Object
    Dim centerCols As Object, leftCols As Object
    Dim h As Variant, col As Variant

    Set centerDict = CreateObject("Scripting.Dictionary")
    Set leftDict = CreateObject("Scripting.Dictionary")
    Set centerCols = CreateObject("Scripting.Dictionary")
    Set leftCols = CreateObject("Scripting.Dictionary")

    For Each h In centerColsArr: centerDict(UCase(h)) = True: Next
    For Each h In leftColsArr: leftDict(UCase(h)) = True: Next

    With ws.UsedRange
        For Each cell In .Cells
            If VarType(cell.Value) = vbString Then cleanedText = Trim(cell.Value) Else cleanedText = ""
            cell.Font.Name = "Arial"

            Select Case True
                Case StrComp(cleanedText, billText, vbTextCompare) = 0
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

            cell.VerticalAlignment = xlCenter

            If centerDict.Exists(cleanedText) Then centerCols(cell.Column) = True
            If leftDict.Exists(cleanedText) Then leftCols(cell.Column) = True
        Next cell
    End With

    For Each col In centerCols.Keys: ws.Columns(col).HorizontalAlignment = xlCenter: Next
    For Each col In leftCols.Keys: ws.Columns(col).HorizontalAlignment = xlLeft: Next
End Sub

'=======================
' PREPARE CAD SHEET
'=======================
Sub PrepareCadSheet(ws As Worksheet, columnToRemove As String)
    Dim cell As Range, foundCell As Range

    ' Remove specified column
    Set foundCell = ws.Cells.Find(What:=columnToRemove, LookIn:=xlValues, _
                                  LookAt:=xlWhole, MatchCase:=False)
    If Not foundCell Is Nothing Then foundCell.EntireColumn.Delete

    ' Convert all text to uppercase
    For Each cell In ws.UsedRange
        If VarType(cell.Value) = vbString Then cell.Value = UCase(cell.Value)
    Next cell
End Sub

'=======================
' HELPER: Check if value exists in array
'=======================
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
