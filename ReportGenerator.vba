Sub GenerateWorkbookReport()
    ' --- SETUP ---
    Dim wb As Workbook
    Dim reportSheet As Worksheet
    Dim reportRow As Long
    Dim ws As Worksheet
    Dim cell As Range
    Dim nm As Name
    Dim vbComp As Object 'VBIDE.VBComponent
    
    Set wb = ThisWorkbook
    reportRow = 1
    
    ' Delete old report sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Workbook_Info").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create a new sheet for the report
    Set reportSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    reportSheet.Name = "Workbook_Info"
    
    ' Function to easily write headers and content
    Sub WriteInfo(header As String, content As String)
        reportSheet.Cells(reportRow, 1).Value = header
        reportSheet.Cells(reportRow, 2).Value = content
        reportSheet.Cells(reportRow, 1).Font.Bold = True
        reportRow = reportRow + 1
    End Sub
    
    ' --- 1. WORKBOOK INFORMATION ---
    WriteInfo "--- WORKBOOK SUMMARY ---", ""
    WriteInfo "File Name:", wb.FullName
    WriteInfo "Creation Date:", wb.BuiltinDocumentProperties("Creation Date").Value
    WriteInfo "Last Author:", wb.BuiltinDocumentProperties("Last Author").Value
    WriteInfo "Last Save Time:", wb.BuiltinDocumentProperties("Last Save Time").Value
    reportRow = reportRow + 1
    
    ' --- 2. WORKSHEET INFORMATION ---
    WriteInfo "--- WORKSHEETS ---", ""
    For Each ws In wb.Worksheets
        WriteInfo "Sheet Name:", ws.Name
        WriteInfo "  - Visible:", IIf(ws.Visible = xlSheetVisible, "Yes", "No")
        WriteInfo "  - Protected:", IIf(ws.ProtectContents, "Yes", "No")
        WriteInfo "  - Used Range:", ws.UsedRange.Address
    Next ws
    reportRow = reportRow + 1
    
    ' --- 3. NAMED RANGES ---
    WriteInfo "--- NAMED RANGES ---", ""
    If wb.Names.Count > 0 Then
        For Each nm In wb.Names
            WriteInfo "Name:", nm.Name
            WriteInfo "  - Refers To:", "'" & nm.RefersTo
            WriteInfo "  - Scope:", nm.Parent.Name
        Next nm
    Else
        WriteInfo "No Named Ranges found.", ""
    End If
    reportRow = reportRow + 1
    
    ' --- 4. CELL FORMULAS AND VALUES ---
    WriteInfo "--- CELL DETAILS (FORMULAS, VALUES, COMMENTS) ---", ""
    For Each ws In wb.Worksheets
        WriteInfo "Sheet:", ws.Name
        reportRow = reportRow + 1
        If WorksheetFunction.CountA(ws.Cells) > 0 Then
            For Each cell In ws.UsedRange
                If cell.HasFormula Then
                    WriteInfo "  - Cell " & cell.Address, "Formula: " & cell.Formula
                ElseIf Not IsEmpty(cell.Value) Then
                    WriteInfo "  - Cell " & cell.Address, "Value: " & cell.Value
                End If
                If Not cell.Comment Is Nothing Then
                     WriteInfo "  - Cell " & cell.Address, "Comment: " & cell.Comment.Text
                End If
            Next cell
        Else
             WriteInfo "  - Sheet is empty", ""
        End If
        reportRow = reportRow + 1
    Next ws
    
    ' --- 5. VBA MACRO CODE ---
    WriteInfo "--- VBA CODE ---", ""
    On Error Resume Next ' In case project is not trusted
    Set vbProj = wb.VBProject
    If Err.Number <> 0 Then
        WriteInfo "ERROR:", "Could not access the VBA project. You may need to 'Trust access to the VBA project object model' in File > Options > Trust Center > Trust Center Settings > Macro Settings."
    Else
        For Each vbComp In wb.VBProject.VBComponents
            WriteInfo "Component Name:", vbComp.Name
            WriteInfo "Component Type:", TypeName(vbComp)
            If vbComp.CodeModule.CountOfLines > 0 Then
                WriteInfo "Code:", vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
            Else
                WriteInfo "Code:", "(No code in this component)"
            End If
            reportRow = reportRow + 1
        Next vbComp
    End If
    On Error GoTo 0
    
    ' --- FINAL FORMATTING ---
    reportSheet.Columns("A:B").AutoFit
    MsgBox "Workbook information report has been generated on the 'Workbook_Info' sheet."

End Sub
