Attribute VB_Name = "AdvancedOptions"
Sub OpenAdvancedOptions()

Dim PW As String
Dim address As String
Application.ScreenUpdating = False

PW = InputBox("Please enter advanced user options access ID.", "ADVANCED USER OPTIONS ACCESS ID", , 4500, 4425)
    
If PW = "open" Or PW = "Open" Or PW = "OPEN" Or PW = "close" Or PW = "Close" Or PW = "CLOSE" Then
    Dim TemplateFileName As String
    ActiveWorkbook.Unprotect Password:="81643"
    Sheets("Macro Controls").Visible = True
    Sheets("Macro Controls").Select
    TemplateFileName = Range("I21")
End If

If PW = "open" Or PW = "Open" Or PW = "OPEN" Then
    address = Range("E22")
    Sheets("Macro Controls").Visible = False
    ActiveWorkbook.Protect Password:="81643"
    Workbooks.Open (address & "\" & TemplateFileName)
    Sheets("IDEXX").Unprotect Password:="81643"
    Sheets("Results Summary").Unprotect Password:="81643"
    Sheets("Na Cal").Unprotect Password:="81643"
    Sheets("K Cal").Unprotect Password:="81643"
    Sheets("Cl Cal").Unprotect Password:="81643"
    Sheets("Instrument Template").Unprotect Password:="81643"
    Sheets("Instrument Template (2)").Unprotect Password:="81643"
    Sheets("Time Slice Locations").Unprotect Password:="81643"
    ActiveWorkbook.Unprotect Password:="81643"
End If

If PW = "close" Or PW = "Close" Or PW = "CLOSE" Then
    Windows(TemplateFileName).Activate
    Sheets("IDEXX").Protect Password:="81643"
    Sheets("Results Summary").Protect Password:="81643"
    Sheets("Na Cal").Protect Password:="81643"
    Sheets("K Cal").Protect Password:="81643"
    Sheets("Cl Cal").Protect Password:="81643"
    Sheets("Instrument Template").Protect Password:="81643"
    Sheets("Instrument Template (2)").Protect Password:="81643"
    Sheets("Time Slice Locations").Protect Password:="81643"
    ActiveWorkbook.Protect Password:="81643"
    ActiveWorkbook.Save
    Workbooks(TemplateFileName).Close
    Sheets("Macro Controls").Select
    Sheets("Macro Controls").Visible = False
    Sheets("Macro Controls").Protect Password:="81643"
    ActiveWorkbook.Protect Password:="81643"
End If

If PW = "81643" Then
    ActiveWorkbook.Unprotect Password:="81643"
    Sheets("Macro Controls").Visible = True
    Sheets("Macro Controls").Select
End If

If PW = "Kite" Or PW = "kite" Or PW = "KITE" Then
    Sheets("Launch Macro").Unprotect Password:="81643"
    ActiveWorkbook.Unprotect Password:="81643"
    Sheets("Launch Macro").Unprotect Password:="81643"
    Sheets("Macro Controls").Unprotect Password:="81643"
    Sheets("Macro Functions").Unprotect Password:="81643"
    Sheets("Error Worksheet").Unprotect Password:="81643"
    Sheets("Error Worksheet").Visible = True
    Sheets("Macro Controls").Visible = True
    Sheets("Macro Functions").Visible = True
    Sheets("Macro Functions").Select
End If

End Sub

Sub CloseAdvancedOptions()

    Application.ScreenUpdating = False
    Sheets("Macro Functions").Visible = False
    Sheets("Macro Controls").Visible = False
    Sheets("Error Worksheet").Visible = False
    Sheets("Launch Macro").Select
    ActiveWorkbook.Protect Password:="81643"
    Sheets("Launch Macro").Protect Password:="81643"
    Sheets("Macro Controls").Protect Password:="81643"
    Sheets("Macro Functions").Protect Password:="81643"
    Sheets("Error Worksheet").Protect Password:="81643"

End Sub

