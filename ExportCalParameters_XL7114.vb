Attribute VB_Name = "Module1"
Sub IDEXX_CalParameters()

Application.ScreenUpdating = False
Dim LotNo As String
Dim MacroWindow As Workbook
Dim IDEXXWindow As Workbook
Dim LockedUnlocked As String

Set MacroWindow = ActiveWorkbook
Application.ScreenUpdating = False

    Sheets("Results Summary").Unprotect Password:="81643"
    Sheets("Results Summary").Select
    LockedUnlocked = Range("B56")
    If LockedUnlocked = "Unlocked" Then
        Sheets("Results Summary").Unprotect Password:="81643"
        MsgBox ("Cannot generate calibration parameters if workbook is unlocked." _
        & Chr(10) & Chr(10) & "Please relock this workbook then reattempt.")
        Exit Sub
    End If
    LotNo = Range("J2")
    Sheets("Results Summary").Protect Password:="81643"
    Sheets(LotNo & "-Calibration").Unprotect Password:="81643"
    Sheets(LotNo & "-Calibration").Select

    Rows("1:4").Select
    Selection.Copy
    
    Workbooks.Add
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A1").Select
    
    Set IDEXXWindow = ActiveWorkbook
    
    MacroWindow.Activate
    Sheets(LotNo & "-Calibration").Protect Password:="81643"
    
    IDEXXWindow.Activate
    Application.DisplayAlerts = False
    Sheets("Sheet2").Select
    ActiveWindow.SelectedSheets.Delete
    Sheets("Sheet3").Select
    ActiveWindow.SelectedSheets.Delete
    ActiveSheet.Name = LotNo & "-Calibration"
End Sub




