' These variables are declared at the module level so all subs can use them
Private reportSheet As Worksheet
Private reportRow As Long
Private noScanZones As Object 'Will hold a Dictionary of ranges to ignore

Sub GenerateSmarterWorkbookReport()
    ' --- OPTIMIZATIONS START ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo Cleanup ' If something goes wrong, jump to the cleanup section
    
    ' --- SETUP ---
    Dim wb As Workbook
    Set wb = ThisWorkbook
    reportRow = 1
    Set noScanZones = CreateObject("Scripting.Dictionary") ' Create a dictionary to hold ranges we should skip
    
    ' Delete old report sheet if it exists
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets("Workbook_Info").Delete
    On Error GoTo Cleanup ' Reset error handling
    Application.DisplayAlerts = True
    
    ' Create a new sheet for the report
    Set reportSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    reportSheet.Name = "Workbook_Info"
    
    ' --- [The report generation logic is now modular] ---
    
    WriteSectionHeader "WORKBOOK SUMMARY"
    AnalyzeWorkbookProperties wb
    
    WriteSectionHeader "WORKSHEETS"
    AnalyzeWorksheets wb
    
    WriteSectionHeader "NAMED RANGES"
    AnalyzeNamedRanges wb
    
    WriteSectionHeader "DATA CONNECTIONS & QUERIES (TABLES)"
    AnalyzeDataConnections wb
    
    WriteSectionHeader "CELL & FORMULA ANALYSIS (GROUPED)"
    AnalyzeCellContent wb
    
    WriteSectionHeader "VBA CODE"
    AnalyzeVBACode wb

    ' --- FINAL FORMATTING ---
    reportSheet.Columns("A:B").AutoFit
    MsgBox "Smarter workbook report has been generated on the 'Workbook_Info' sheet."

Cleanup:
    ' --- OPTIMIZATIONS END ---
    ' This section runs whether the macro finishes or fails, ensuring Excel is returned to a normal state.
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Set noScanZones = Nothing ' Clean up the dictionary object
    If Err.Number <> 0 Then
        MsgBox "An error occurred: " & Err.Description
    End If

End Sub

' --- HELPER SUBROUTINES FOR WRITING TO THE REPORT ---
Private Sub WriteInfo(header As String, content As String, Optional indentLevel As Integer = 0)
    reportSheet.Cells(reportRow, 1).Value = String(indentLevel * 2, " ") & header
    reportSheet.Cells(reportRow, 2).Value = content
    reportSheet.Cells(reportRow, 1).Font.Bold = (indentLevel = 0)
    reportRow = reportRow + 1
End Sub

Private Sub WriteSectionHeader(title As String)
    reportRow = reportRow + 1
    With reportSheet.Cells(reportRow, 1)
        .Value = "--- " & UCase(title) & " ---"
        .Font.Bold = True
    End With
    reportSheet.Range(reportSheet.Cells(reportRow, 1), reportSheet.Cells(reportRow, 2)).Merge
    reportRow = reportRow + 2
End Sub


' --- ANALYSIS MODULES ---

Private Sub AnalyzeWorkbookProperties(wb As Workbook)
    WriteInfo "File Name:", wb.FullName
    WriteInfo "Last Save Time:", wb.BuiltinDocumentProperties("Last Save Time").Value
End Sub

Private Sub AnalyzeWorksheets(wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        WriteInfo "Sheet Name:", ws.Name
        WriteInfo "Visible:", IIf(ws.Visible = xlSheetVisible, "Yes", "No"), 1
        WriteInfo "Protected:", IIf(ws.ProtectContents, "Yes", "No"), 1
        WriteInfo "Used Range:", ws.UsedRange.Address, 1
    Next ws
End Sub

Private Sub AnalyzeNamedRanges(wb As Workbook)
    Dim nm As Name
    If wb.Names.Count > 0 Then
        For Each nm In wb.Names
            WriteInfo "Name:", nm.Name
            WriteInfo "Refers To:", "'" & nm.RefersTo, 1
            WriteInfo "Scope:", nm.Parent.Name, 1
        Next nm
    Else
        WriteInfo "No Named Ranges found.", ""
    End If
End Sub

' --- NEW, MORE ROBUST VERSION ---
Private Sub AnalyzeDataConnections(wb As Workbook)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim tq As QueryTable
    Dim wq As WorkbookQuery
    Dim addressKey As String
    Dim sheetName As String
    
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            Set tq = Nothing
            On Error Resume Next
            Set tq = lo.QueryTable
            On Error GoTo 0
            
            If Not tq Is Nothing Then
                WriteInfo "Table Name:", lo.Name & " on sheet '" & ws.Name & "'"
                WriteInfo "Range:", lo.Range.Address, 1
                WriteInfo "Source Type:", "Legacy QueryTable", 1
                WriteInfo "Connection:", tq.Connection, 1
                
                '--- FIX: Properly handle sheet names with apostrophes ---
                If Not lo.DataBodyRange Is Nothing Then
                    sheetName = lo.Parent.Name
                    ' Escape any single quotes in the sheet name by doubling them up
                    sheetName = Replace(sheetName, "'", "''")
                    addressKey = "'" & sheetName & "'!" & lo.DataBodyRange.Address
                    noScanZones(addressKey) = True ' Add to no-scan list
                End If
            End If
        Next lo
    Next ws
    
    On Error Resume Next
    For Each wq In wb.Queries
        If wq.ListObject Is Nothing Then
             WriteInfo "Query Name:", wq.Name & " (Connection Only)"
        Else
             WriteInfo "Query Name:", wq.Name & " (Loads to Table '" & wq.ListObject.Name & "')"
             WriteInfo "Table Location:", wq.ListObject.Range.Address(External:=True), 1
             
             '--- FIX: Properly handle sheet names with apostrophes ---
             If Not wq.ListObject.DataBodyRange Is Nothing Then
                Set lo = wq.ListObject
                sheetName = lo.Parent.Name
                ' Escape any single quotes in the sheet name by doubling them up
                sheetName = Replace(sheetName, "'", "''")
                addressKey = "'" & sheetName & "'!" & lo.DataBodyRange.Address
                noScanZones(addressKey) = True ' Add to no-scan list
             End If
        End If
        WriteInfo "Source Type:", "Power Query", 1
        WriteInfo "M Code:", wq.Formula, 1
    Next wq
    On Error GoTo 0
End Sub

Private Sub AnalyzeCellContent(wb As Workbook)
    Dim ws As Worksheet, cell As Range, col As Range, scanArea As Range
    Dim isTrackingFormula As Boolean, currentFormulaR1C1 As String, formulaStartRow As Long
    
    For Each ws In wb.Worksheets
        WriteInfo "Scanning Sheet:", "'" & ws.Name & "'"
        If ws.UsedRange.Cells.Count > 2000000 Then ' Safety check for huge sheets
            WriteInfo "NOTE:", "Sheet is very large, analysis skipped to prevent freezing.", 1
            GoTo NextSheet ' Skip very large sheets
        End If

        For Each col In ws.UsedRange.Columns
            isTrackingFormula = False
            currentFormulaR1C1 = ""
            formulaStartRow = 0
            
            'This loop is inefficient for huge numbers of no-scan zones, but fine for dozens.
            For Each cell In col.Cells
                ' Check if the cell is inside a no-scan zone (like a query result table)
                Dim key As Variant, skipCell As Boolean
                skipCell = False
                For Each key In noScanZones.Keys
                    '--- THE FIX IS ON THE NEXT LINE ---
                    ' Changed ws.Range(key) to Application.Range(key) to handle cross-sheet references correctly
                    If Not Intersect(cell, Application.Range(key)) Is Nothing Then
                        skipCell = True
                        Exit For
                    End If
                Next key

                If Not skipCell Then
                    ' --- Formula Grouping Logic ---
                    If cell.HasFormula Then
                        If isTrackingFormula And cell.FormulaR1C1 = currentFormulaR1C1 Then
                            ' Continue tracking the same formula
                        Else
                            ' End the previous tracking run if there was one
                            If isTrackingFormula Then
                                'Check if the start row is valid before creating a range
                                If cell.Row > formulaStartRow Then
                                    WriteInfo "Formula Range:", ws.Range(ws.Cells(formulaStartRow, col.Column), cell.Offset(-1, 0)).Address, 1
                                Else 'Handle single-cell formula range
                                    WriteInfo "Formula Cell:", ws.Cells(formulaStartRow, col.Column).Address, 1
                                End If
                                WriteInfo "Formula (R1C1):", currentFormulaR1C1, 2
                            End If
                            ' Start a new tracking run
                            isTrackingFormula = True
                            currentFormulaR1C1 = cell.FormulaR1C1
                            formulaStartRow = cell.Row
                        End If
                    Else
                        ' The formula block has ended
                        If isTrackingFormula Then
                            If cell.Row > formulaStartRow Then
                                WriteInfo "Formula Range:", ws.Range(ws.Cells(formulaStartRow, col.Column), cell.Offset(-1, 0)).Address, 1
                            Else 'Handle single-cell formula range
                                WriteInfo "Formula Cell:", ws.Cells(formulaStartRow, col.Column).Address, 1
                            End If
                            WriteInfo "Formula (R1C1):", currentFormulaR1C1, 2
                            isTrackingFormula = False
                        End If
                        
                        ' Report individual values and comments
                        If Not IsEmpty(cell.Value) Then WriteInfo "Cell " & cell.Address, "Value: " & cell.Value, 1
                        If Not cell.Comment Is Nothing Then WriteInfo "Cell " & cell.Address, "Comment: " & cell.Comment.Text, 1
                    End If
                End If
            Next cell
            
            ' After the loop, close any open formula tracking for that column
            If isTrackingFormula Then
                 WriteInfo "Formula Range:", ws.Range(ws.Cells(formulaStartRow, col.Column), ws.Cells(col.Rows.Count + col.Row - 1, col.Column)).Address, 1
                 WriteInfo "Formula (R1C1):", currentFormulaR1C1, 2
            End If
        Next col
NextSheet:
    Next ws
End Sub

Private Sub AnalyzeVBACode(wb As Workbook)
    Dim vbComp As Object 'VBIDE.VBComponent
    On Error Resume Next
    If wb.VBProject.Protection = 1 Then 'vbext_pp_locked
        WriteInfo "VBA Project is Locked", ""
        Exit Sub
    End If
    
    For Each vbComp In wb.VBProject.VBComponents
        WriteInfo "Component:", vbComp.Name & " (" & TypeName(vbComp) & ")"
        If vbComp.CodeModule.CountOfLines > 0 Then
            WriteInfo "Code:", vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines), 1
        Else
            WriteInfo "Code:", "(No code in this component)", 1
        End If
    Next vbComp
End Sub
