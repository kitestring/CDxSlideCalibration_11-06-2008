Attribute VB_Name = "Module2"
Sub Lock_Unlock()
Attribute Lock_Unlock.VB_ProcData.VB_Invoke_Func = "L\n14"
    
Dim PW As String
Dim LotNo As String
Dim GF1 As Integer
Dim InstrumentName As String
Dim YESNO As Byte
Application.ScreenUpdating = False
YESNO = 0

PW = InputBox("Please enter advanced user options access ID.", "ADVANCED USER OPTIONS ACCESS ID", , 4500, 4425)
  
Sheets("Results Summary").Unprotect Password:="81643"
Sheets("Results Summary").Select
  
Select Case PW
    Case "81643"
        If Range("B56") = "Unlocked" Then GoTo LockDown
        If Range("B56") = "Locked" Then GoTo UnlockWorkbook
    Case "RD81643"
        If Range("C56") = "Links Active" Then GoTo RemoveLinks
        If Range("C56") = "Links Severed" Then GoTo UnlockWorkbook
    Case Else
        Exit Sub
End Select
    
    
UnlockWorkbook:
    
    ActiveWorkbook.Unprotect Password:="81643"
            Range("B56") = "Unlocked"
            LotNo = Range("J2")
        Sheets(LotNo & "-Calibration").Unprotect Password:="81643"
        Sheets("K Cal").Unprotect Password:="81643"
        Sheets("Cl Cal").Unprotect Password:="81643"
        Sheets("Na Cal").Unprotect Password:="81643"
        Sheets("Na Cal").Select
        Range("A1").Select
    
        Cells.Find(What:="H31 Bias2", After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(1, -11) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
        InstrumentName = "Nothing"
        Do While GF1 > 0
            If Range("B" & GF1) = "x" Then
                Range("A1").Select
                Sheets("Results Summary").Select
                GoTo CreateLinks
            End If
            If InstrumentName <> Range("B" & GF1) Then
                InstrumentName = Range("B" & GF1)
                Sheets(InstrumentName).Unprotect Password:="81643"
            End If
            GF1 = GF1 + 1
        Loop
 
 
CreateLinks:
        
        
    If PW = "81643" Then
        Range("C56") = "Links Active"
        Sheets(LotNo & "-Calibration").Select
            Range("B2") = "='Results Summary'!N15"
            Range("B3") = "='Results Summary'!N25"
            Range("B4") = "='Results Summary'!N36"
            Range("C2") = "='Results Summary'!N14"
            Range("C3") = "='Results Summary'!N23"
            Range("C4") = "='Results Summary'!N34"
            Range("D3") = "='Results Summary'!N24"
            Range("E2") = "='Results Summary'!N16"
            Range("E3") = "='Results Summary'!N28"
            Range("F3") = "='Results Summary'!N26"
            Range("H2") = "='Na Cal'!E20"
            Range("H3") = "='K Cal'!E24"
            Range("H4") = "='Cl Cal'!E21"
            Range("I2") = "='Na Cal'!E21"
            Range("I3") = "='K Cal'!E25"
            Range("I4") = "='Cl Cal'!E22"
            Range("J2") = "='Na Cal'!E22"
            Range("J3") = "='K Cal'!E26"
            Range("J4") = "='Cl Cal'!E23"
            Range("K2") = "='Na Cal'!E23"
            Range("K3") = "='K Cal'!E27"
            Range("K4") = "='Cl Cal'!E24"
        Sheets("Results Summary").Select
    End If
    
    MsgBox ("Be sure to lock-up the workbook when you are done.")
    
Exit Sub


LockDown:
    
    ActiveWorkbook.Protect Password:="81643"
            Range("B56") = "Locked"
            LotNo = Range("J2")
        Sheets("Results Summary").Protect Password:="81643"
        Sheets(LotNo & "-Calibration").Protect Password:="81643"
        Sheets("K Cal").Protect Password:="81643"
        Sheets("Cl Cal").Protect Password:="81643"
        Sheets("Na Cal").Select
        Range("A1").Select
    
        Cells.Find(What:="H31 Bias2", After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(1, -11) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
        InstrumentName = "Nothing"
        Do While GF1 > 0
            If Range("B" & GF1) = "x" Then
                Range("A1").Select
                Sheets("Na Cal").Protect Password:="81643"
                Sheets("Results Summary").Select
                Exit Sub
            End If
            If InstrumentName <> Range("B" & GF1) Then
                InstrumentName = Range("B" & GF1)
                Sheets(InstrumentName).Protect Password:="81643"
            End If
            GF1 = GF1 + 1
        Loop
        
RemoveLinks:
    If Range("B56") = "Locked" Then YESNO = 1
    LotNo = Range("J2")
    Range("C56") = "Links Severed"
    Sheets(LotNo & "-Calibration").Unprotect Password:="81643"
    Sheets(LotNo & "-Calibration").Select
    Range("A1:P4").Copy
    Range("A1:P4").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("A1").Select
    
    If YESNO = 1 Then
        Sheets("Results Summary").Protect Password:="81643"
        Sheets(LotNo & "-Calibration").Protect Password:="81643"
    End If
    Sheets("Results Summary").Select
    
End Sub



