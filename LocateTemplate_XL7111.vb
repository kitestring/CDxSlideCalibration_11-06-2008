Attribute VB_Name = "LocateTemplate"
Sub Locate_Template()

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim TemplateLocation As Variant
Dim ColumnNumber As Byte
Dim TemplateFileName As String
Dim StartColumn As Byte
Dim EndColumn As Byte
StartColumn = 3

        Sheets("Macro Controls").Select
        Sheets("Macro Controls").Visible = True
        Sheets("Macro Controls").Unprotect Password:="81643"

        TemplateLocation = Application.GetOpenFilename(FileFilter:="Microsoft Excel Files(*.xls),*.xls", Title:="Template Location", MultiSelect:=False)
        If TemplateLocation = False Then Exit Sub
        
        Range("C35") = TemplateLocation
        
        Range("C35").Select
        Selection.TextToColumns Destination:=Range("C35"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
            :="\"
        
        ColumnNumber = StartColumn
        
        Do While IsEmpty(Cells(35, ColumnNumber))
            ColumnNumber = ColumnNumber + 1
        Loop
        
        Do While Not (IsEmpty(Cells(35, ColumnNumber)))
            ColumnNumber = ColumnNumber + 1
        Loop
        
        ColumnNumber = ColumnNumber - 1
        EndColumn = ColumnNumber - 1
        TemplateFileName = Cells(35, ColumnNumber)
        Range("I21") = TemplateFileName
        
        TemplateLocation = Cells(35, StartColumn)
        StartColumn = StartColumn + 1
        
        For ColumnNumber = StartColumn To EndColumn
            TemplateLocation = TemplateLocation & "\" & Cells(35, ColumnNumber)
        Next
        
        Range("E22") = TemplateLocation
        
        Rows(35).ClearContents
        Range("A1").Select
        Sheets("Macro Controls").Protect Password:="81643"

End Sub

