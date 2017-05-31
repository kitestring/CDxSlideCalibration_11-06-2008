Attribute VB_Name = "Calibration"
Sub Catalyst_1()

'**********************************************************
'Open .csv files indicated by Macro Functions worksheet.
'Correct any erros in the .csv data base, sorting errors
'or shifting errors, and copy paste data into instruement
'data template.
'***********************************************************

Dim FileName As String
Dim WindowName As String
Dim RunNo As String
Dim LotNo As String
Dim OperatorID As String
Dim Volume As String
Dim ExpDate As String
Dim SourceData As String
Dim DataDump As String
Dim IR As Byte
Dim InstrumentName As String
Dim BR As Byte
Dim BDR As Byte
Dim BufferName As String
Dim MaxBufBias As Double
Dim LargestBias As String
Dim NaCR As Integer
Dim KCR As Integer
Dim ClCR As Integer
Dim EC As Integer
Dim GF1 As Integer
Dim GF2 As Integer
Dim GF3 As Byte
Dim GF4 As Integer
Dim GF5 As Integer
Dim GF6 As Byte
Dim NaBias As Double
Dim KBias As Double
Dim ClBias As Double
Dim SaveDocument As Variant
Dim wb As Workbook
Dim Na_DryS As Byte
Dim Na_DryR As Byte
Dim Na_Dry1 As String
Dim Na_Dry2 As String
Dim Na_WetS As Byte
Dim Na_WetR As Byte
Dim Na_Wet1 As String
Dim Na_Wet2 As String
Dim K_DryS As Byte
Dim K_DryR As Byte
Dim K_Dry1 As String
Dim K_Dry2 As String
Dim K_WetS As Byte
Dim K_WetR As Byte
Dim K_Wet1 As String
Dim K_Wet2 As String
Dim Cl_DryS As Byte
Dim Cl_DryR As Byte
Dim Cl_Dry1 As String
Dim Cl_Dry2 As String
Dim Cl_WetS As Byte
Dim Cl_WetR As Byte
Dim Cl_Wet1  As String
Dim Cl_Wet2 As String
Dim SearchValue As Byte
Dim RangeOffSet As Integer
Dim T1 As String
Dim T2 As String
Dim NaScalar As Double
Dim KScalar As Double
Dim ClScalar As Double
Dim SV As Byte
Dim PanelID As String
Dim DryInt As Double
Dim IntInt As Double
Dim CalibrationWindow As Workbook
Dim MacroWindow As Workbook
Dim TemplateFileName As String
Dim NumberOfInstruments As Byte
Dim InstrumentList() As String
Dim i As Byte
Dim j As Byte
Dim k As Byte
Dim NoOfBuffers As Byte
Dim FlagFormula As String
Dim ColumnLetter As String
Dim RowNumber As String
Dim FlagRow As Integer

On Error GoTo ErrorHandler
Set MacroWindow = ActiveWorkbook

    If IsEmpty(Range("I9")) Then
        MsgBox "You have not entered a .cvs source file directory address.", vbExclamation
        Exit Sub
    End If

    Do While RunNo = ""
        RunNo = InputBox("Please enter the run number.  Cannot be left blank.  Type exit to terminate macro.", "RUN #", , 4500, 2625)
        RunNo = Trim(RunNo)
        If RunNo = "exit" Then Exit Sub
    Loop
    Do While OperatorID = ""
        OperatorID = InputBox("Please enter the operator ID.  Cannot be left blank.  Type exit to terminate macro.", "OPERATOR ID", , 4500, 3825)
        If OperatorID = "exit" Then Exit Sub
    Loop
    Do While LotNo = ""
        LotNo = InputBox("Please enter the lot number.  Cannot be left blank.  Type exit to terminate macro.", "LOT #", , 4500, 5025)
        LotNo = Trim(LotNo)
        If LotNo = "exit" Then Exit Sub
    Loop
    Do While ExpDate = ""
        ExpDate = InputBox("Please enter the experiment date (mm/dd/yyyy).  Cannot be left blank.  Type exit to terminate macro.", "DATE", , 4500, 6225)
        If ExpDate = "exit" Then Exit Sub
    Loop
    Do While Volume = ""
        Volume = InputBox("Please enter the volume used in micro liters.  Cannot be left blank.  Type exit to terminate macro.", "VOLUME", , 4500, 7425)
        If Volume = "exit" Then Exit Sub
    Loop

    SourceData = Trim(Range("I9")) + "\"

ActiveWorkbook.Unprotect Password:="81643"
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Sheets("Macro Functions").Visible = True
Sheets("Macro Controls").Visible = True
Sheets("Error Worksheet").Visible = False

    IR = 10
    BR = 10
    EC = -6
    
    Sheets("Macro Controls").Select
    
        NaBias = 1000000
        KBias = 1000000
        ClBias = 1000000
        
        Na_DryS = Range("E29")
        Na_DryR = Range("H29")
        Na_WetS = Range("K29")
        Na_WetR = Range("M29")
        K_DryS = Range("E30")
        K_DryR = Range("H30")
        K_WetS = Range("K30")
        K_WetR = Range("M30")
        Cl_DryS = Range("E31")
        Cl_DryR = Range("H31")
        Cl_WetS = Range("K31")
        Cl_WetR = Range("M31")
    
        DataDump = Range("E22") + "\"
        TemplateFileName = Range("I21")
    
    Workbooks.Open (DataDump & TemplateFileName), ReadOnly:=True
    Windows(TemplateFileName).Activate
    Set CalibrationWindow = ActiveWorkbook
    
    Sheets("IDEXX").Unprotect Password:="81643"
    Sheets("Results Summary").Unprotect Password:="81643"
    Sheets("Na Cal").Unprotect Password:="81643"
    Sheets("K Cal").Unprotect Password:="81643"
    Sheets("Cl Cal").Unprotect Password:="81643"
    Sheets("Instrument Template").Unprotect Password:="81643"
    Sheets("Instrument Template (2)").Unprotect Password:="81643"
    Sheets("Time Slice Locations").Unprotect Password:="81643"
    ActiveWorkbook.Unprotect Password:="81643"
    
InstrumentLoop:

    MacroWindow.Activate
    Sheets("Macro Functions").Select
    
        GF3 = 0
        SV = 0
        
        Do While IR > 9 And IR < 19
            Select Case Range("C" & IR)
                Case True
                    InstrumentName = Range("D" & IR)
                    CalibrationWindow.Activate
                    Sheets("Instrument Template").Select
                    Sheets("Instrument Template").Copy Before:=Sheets("Instrument Template")
                    ActiveSheet.Name = InstrumentName
                    
                    MacroWindow.Activate
                    Sheets("Macro Controls").Select
                    If Range("T" & (IR - 5)) = "Active" Then
                        NaScalar = Range("Q" & (IR - 5))
                        KScalar = Range("R" & (IR - 5))
                        ClScalar = Range("S" & (IR - 5))
                        SV = 1
                    End If
                    IR = IR + 1
                    GoTo BufferLoop
                Case False
                    IR = IR + 1
            End Select
        Loop

        CalibrationWindow.Activate
            Sheets("Instrument Template").Delete
            Sheets("Instrument Template (2)").Delete
            Sheets("Time Slice Locations").Delete
            
        IR = 10
        BR = 10
            
        GoTo InstrumentConstants

BufferLoop:

    MacroWindow.Activate
    Sheets("Macro Functions").Select

        Do While BR > 9 And BR < 18
            Select Case Range("F" & BR)
                Case True
                    BufferName = Range("G" & BR)
                    BDR = Range("H" & BR)
                    BR = BR + 1
                    GoTo csvFileOpen
                Case False
                    BR = BR + 1
            End Select
        Loop
        
        BR = 10
        GoTo InstrumentLoop
        
csvFileOpen:

    WindowName = LotNo + "_" + InstrumentName + "_" + BufferName + "_" + RunNo + ".csv"
    FileName = SourceData + LotNo + "_" + InstrumentName + "_" + BufferName + "_" + RunNo + ".csv"
    Workbooks.Open FileName:=FileName, ReadOnly:=True
    
    Windows(WindowName).Activate
        
        If Range("A1") = "Slide Position" Then GoTo csvFileOpen2
        
        GF1 = 2
        GF2 = 0
    
        Do While GF1 < 8
            If IsEmpty(Range("BF" & GF1)) Then GF2 = GF2 + 1
            GF1 = GF1 + 1
        Loop
    
        If GF2 = 6 Then
            Range("C2:C19").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("K2:N19").Select
            Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
    
        If Range("B2") = 2 Then
            Range("D1").Select
            Range("A1:BI19").Sort Key1:=Range("D2"), Order1:=xlAscending, Key2:=Range("B2"), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase _
                :=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
                DataOption2:=xlSortNormal
        End If
        
        If SV = 1 Then
            GF6 = 2
            Do While Not (IsEmpty(Range("R" & GF6)))
                Range("R" & GF6) = NaScalar
                Range("R" & (GF6 + 1)) = KScalar
                Range("R" & (GF6 + 2)) = ClScalar
                GF6 = GF6 + 3
            Loop
        End If
    
        Range("A1:BI19").Copy
        
    CalibrationWindow.Activate
        Sheets(InstrumentName).Select
        Range("P" & BDR).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
        
        Range("D" & BDR + 10 & ":N" & BDR + 10).Select
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, address:=FileName, TextToDisplay:=FileName
        With Selection.Font
            .Name = "Calibri"
            .FontStyle = "Regular"
            .Size = 8
            .Underline = xlUnderlineStyleSingle
            .ThemeColor = xlThemeColorHyperlink
        End With
        
    Sheets("Time Slice Locations").Select
    GF4 = 1
    
    Do While GF4 < 7
        Select Case GF4
            Case 1
                SearchValue = Na_DryS
                RangeOffSet = (Na_DryR - 1) * -1
            Case 2
                SearchValue = Na_WetS
                RangeOffSet = (Na_WetR - 1) * -1
            Case 3
                SearchValue = K_DryS
                RangeOffSet = (K_DryR - 1) * -1
            Case 4
                SearchValue = K_WetS
                RangeOffSet = (K_WetR - 1) * -1
            Case 5
                SearchValue = Cl_DryS
                RangeOffSet = (Cl_DryR - 1) * -1
            Case 6
                SearchValue = Cl_WetS
                RangeOffSet = (Cl_WetR - 1) * -1
        End Select
    
        Cells.Find(What:=SearchValue, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(1, 0).Activate
        T2 = ActiveCell
        ActiveCell.Offset(0, RangeOffSet).Activate
        T1 = ActiveCell
        
        Select Case GF4
            Case 1
                Na_Dry1 = Trim(T1)
                Na_Dry2 = Trim(T2)
            Case 2
                Na_Wet1 = Trim(T1)
                Na_Wet2 = Trim(T2)
            Case 3
                K_Dry1 = Trim(T1)
                K_Dry2 = Trim(T2)
            Case 4
                K_Wet1 = Trim(T1)
                K_Wet2 = Trim(T2)
            Case 5
                Cl_Dry1 = Trim(T1)
                Cl_Dry2 = Trim(T2)
            Case 6
                Cl_Wet1 = Trim(T1)
                Cl_Wet2 = Trim(T2)
        End Select
        GF4 = GF4 + 1
    Loop
        
    Sheets(InstrumentName).Select
    GF4 = 3
    GF5 = 1
        
    Do While GF4 < 9
        Range("D" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & Na_Dry1 & BDR + GF5 & ":" & Na_Dry2 & BDR + GF5 & "),0)-AH" & BDR + GF5
        Range("F" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & Na_Wet1 & BDR + GF5 & ":" & Na_Wet2 & BDR + GF5 & "),0)-AH" & BDR + GF5
        Range("H" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & K_Dry1 & BDR + GF5 + 1 & ":" & K_Dry2 & BDR + GF5 + 1 & "),0)-AH" & BDR + GF5 + 1
        Range("J" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & K_Wet1 & BDR + GF5 + 1 & ":" & K_Wet2 & BDR + GF5 + 1 & "),0)-AH" & BDR + GF5 + 1
        Range("L" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & Cl_Dry1 & BDR + GF5 + 2 & ":" & Cl_Dry2 & BDR + GF5 + 2 & "),0)-AH" & BDR + GF5 + 2
        Range("N" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & Cl_Wet1 & BDR + GF5 + 2 & ":" & Cl_Wet2 & BDR + GF5 + 2 & "),0)-AH" & BDR + GF5 + 2
        GF4 = GF4 + 1
        GF5 = GF5 + 3
    Loop

        
        GF1 = 3
        GF2 = 0
        
        Do While GF1 < 9
            Select Case Range("B" & BDR + GF1)
                Case 0
                    Range("B" & BDR + GF1 & ":N" & BDR + GF1).ClearContents
                Case Else
                    GF2 = GF2 + 1
            End Select
            GF1 = GF1 + 1
        Loop
        
        Range("B" & BDR + 11) = GF2
        Range("A1").Select
        
    Workbooks(WindowName).Close saveChanges:=False
    
    GoTo BufferLoop
    
csvFileOpen2:

    If GF3 = 0 Then
        CalibrationWindow.Activate
            Sheets(InstrumentName).Delete
            GF3 = 1
            Sheets("Instrument Template (2)").Select
            Sheets("Instrument Template (2)").Copy Before:=Sheets("Instrument Template (2)")
            ActiveSheet.Name = InstrumentName
    End If
    
    Windows(WindowName).Activate
    
        GF1 = 2
        GF2 = 0
    
        Do While Not IsEmpty(Range("A" & GF1))
            GF1 = GF1 + 1
        Loop
        
        If ((GF1 - 2) / 3) - Int((GF1 - 2) / 3) <> 0 Then
            GF1 = GF1 - 1
            Range("A" & GF1 & ":BM" & GF1).ClearContents
        End If
    
        If Range("B2") = 2 Then
            Range("D1").Select
            Range("A1:BI19").Sort Key1:=Range("D2"), Order1:=xlAscending, Key2:=Range("B2"), Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase _
                :=False, Orientation:=xlTopToBottom, DataOption1:=xlSortNormal, _
                DataOption2:=xlSortNormal
        End If
    
        If SV = 1 Then
            GF6 = 2
            Do While Not (IsEmpty(Range("T" & GF6)))
                Range("T" & GF6) = NaScalar
                Range("T" & (GF6 + 1)) = KScalar
                Range("T" & (GF6 + 2)) = ClScalar
                GF6 = GF6 + 3
            Loop
        End If
    
        Range("A1:BM19").Copy
        
     CalibrationWindow.Activate
        Sheets(InstrumentName).Select
        Range("P" & BDR).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
        
        Range("D" & BDR + 10 & ":N" & BDR + 10).Select
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, address:=FileName, TextToDisplay:=FileName
        With Selection.Font
            .Name = "Calibri"
            .FontStyle = "Regular"
            .Size = 8
            .Underline = xlUnderlineStyleSingle
            .ThemeColor = xlThemeColorHyperlink
        End With
    
    Sheets("Time Slice Locations").Select
    GF4 = 1
    
    Do While GF4 < 7
        Select Case GF4
            Case 1
                SearchValue = Na_DryS
                RangeOffSet = (Na_DryR - 1) * -1
            Case 2
                SearchValue = Na_WetS
                RangeOffSet = (Na_WetR - 1) * -1
            Case 3
                SearchValue = K_DryS
                RangeOffSet = (K_DryR - 1) * -1
            Case 4
                SearchValue = K_WetS
                RangeOffSet = (K_WetR - 1) * -1
            Case 5
                SearchValue = Cl_DryS
                RangeOffSet = (Cl_DryR - 1) * -1
            Case 6
                SearchValue = Cl_WetS
                RangeOffSet = (Cl_WetR - 1) * -1
        End Select
    
        Range("A1").Select
        Cells.Find(What:=SearchValue, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
        ActiveCell.Offset(2, 0).Activate
        T2 = ActiveCell
        ActiveCell.Offset(0, RangeOffSet).Activate
        T1 = ActiveCell
        
        Select Case GF4
            Case 1
                Na_Dry1 = Trim(T1)
                Na_Dry2 = Trim(T2)
            Case 2
                Na_Wet1 = Trim(T1)
                Na_Wet2 = Trim(T2)
            Case 3
                K_Dry1 = Trim(T1)
                K_Dry2 = Trim(T2)
            Case 4
                K_Wet1 = Trim(T1)
                K_Wet2 = Trim(T2)
            Case 5
                Cl_Dry1 = Trim(T1)
                Cl_Dry2 = Trim(T2)
            Case 6
                Cl_Wet1 = Trim(T1)
                Cl_Wet2 = Trim(T2)
        End Select
        
        GF4 = GF4 + 1
    Loop
        
    Sheets(InstrumentName).Select
    GF4 = 3
    GF5 = 1
        
    Do While GF4 < 9
        Range("D" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & Na_Dry1 & BDR + GF5 & ":" & Na_Dry2 & BDR + GF5 & "),0)-AJ" & BDR + GF5
        Range("F" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & Na_Wet1 & BDR + GF5 & ":" & Na_Wet2 & BDR + GF5 & "),0)-AJ" & BDR + GF5
        Range("H" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & K_Dry1 & BDR + GF5 + 1 & ":" & K_Dry2 & BDR + GF5 + 1 & "),0)-AJ" & BDR + GF5 + 1
        Range("J" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & K_Wet1 & BDR + GF5 + 1 & ":" & K_Wet2 & BDR + GF5 + 1 & "),0)-AJ" & BDR + GF5 + 1
        Range("L" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & Cl_Dry1 & BDR + GF5 + 2 & ":" & Cl_Dry2 & BDR + GF5 + 2 & "),0)-AJ" & BDR + GF5 + 2
        Range("N" & BDR + GF4) = "=ROUNDDOWN(AVERAGE(" & Cl_Wet1 & BDR + GF5 + 2 & ":" & Cl_Wet2 & BDR + GF5 + 2 & "),0)-AJ" & BDR + GF5 + 2
        GF4 = GF4 + 1
        GF5 = GF5 + 3
    Loop
    
        GF1 = 3
        GF2 = 0
        
        Do While GF1 < 9
            Select Case Range("B" & BDR + GF1)
                Case 0
                    Range("B" & BDR + GF1 & ":N" & BDR + GF1).ClearContents
                Case Else
                    GF2 = GF2 + 1
            End Select
            GF1 = GF1 + 1
        Loop
        
        Range("B" & BDR + 11) = GF2
        Range("A1").Select
        
    Workbooks(WindowName).Close saveChanges:=False
    
    GoTo BufferLoop
    
    
'******************************************************************
'More instructions
'******************************************************************
    
InstrumentConstants:
    
    MacroWindow.Activate
    Sheets("Macro Functions").Select
        
        Do While IR > 9 And IR < 19
            Select Case Range("C" & IR)
                Case True
                    EC = EC + 6
                    InstrumentName = Range("D" & IR)
                    IR = IR + 1
                    GoTo BufferConstants
                Case False
                    IR = IR + 1
            End Select
        Loop
        
        GF2 = 0
        GoTo NaDataRemoval
    
BufferConstants:
    
    MacroWindow.Activate
    Sheets("Macro Functions").Select

        Do While BR > 9 And BR < 18
            Select Case Range("F" & BR)
                Case True
                    BufferName = Range("G" & BR)
                    BDR = Range("H" & BR)
                    NaCR = EC + Range("I" & BR)
                    KCR = EC + Range("J" & BR)
                    ClCR = EC + Range("K" & BR)
                    BR = BR + 1
                    GoTo NaKDataLinks
                Case False
                    BR = BR + 1
            End Select
        Loop
    
        BR = 10
        GoTo InstrumentConstants
    
NaKDataLinks:

    If BDR = 101 Or BDR = 121 Or BDR = 141 Then GoTo ClDataLinks

    CalibrationWindow.Activate
    Sheets(InstrumentName).Select
    
        GF1 = Range("B" & BDR + 11)
        
    Sheets("Na Cal").Select
    
        Range("B" & NaCR + 1 & ":B" & NaCR + GF1) = InstrumentName
        Range("C" & NaCR + 1) = "=" & InstrumentName & "!B" & BDR + 3
        Range("D" & NaCR + 1) = "=" & InstrumentName & "!E" & BDR + 3
        Range("E" & NaCR + 1) = "=" & InstrumentName & "!F" & BDR + 3
            
        Range("C" & NaCR + 1 & ":E" & NaCR + 1).Copy
        Range("C" & NaCR + 2 & ":E" & NaCR + GF1).Select
            Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
        
    Sheets("K Cal").Select
    
        Range("B" & KCR + 1 & ":B" & KCR + GF1) = InstrumentName
        Range("C" & KCR + 1) = "=" & InstrumentName & "!B" & BDR + 3
        Range("D" & KCR + 1) = "=" & InstrumentName & "!I" & BDR + 3
        Range("E" & KCR + 1) = "=" & InstrumentName & "!J" & BDR + 3
            
        Range("C" & KCR + 1 & ":E" & KCR + 1).Copy
        Range("C" & KCR + 2 & ":E" & KCR + GF1).Select
            Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False

ClDataLinks:

    If BDR = 21 Or BDR = 61 Then GoTo BufferConstants
    
    CalibrationWindow.Activate
    Sheets(InstrumentName).Select
    
        GF1 = Range("B" & BDR + 11)
    
    Sheets("Cl Cal").Select
    
        Range("B" & ClCR + 1 & ":B" & ClCR + GF1) = InstrumentName
        Range("C" & ClCR + 1) = "=" & InstrumentName & "!B" & BDR + 3
        Range("D" & ClCR + 1) = "=" & InstrumentName & "!M" & BDR + 3
        Range("E" & ClCR + 1) = "=" & InstrumentName & "!N" & BDR + 3
            
        Range("C" & ClCR + 1 & ":E" & ClCR + 1).Copy
        Range("C" & ClCR + 2 & ":E" & ClCR + GF1).Select
            Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
    
    GoTo BufferConstants
    
NaDataRemoval:

    Select Case GF2
        Case 0
            BufferName = "H31-H31"
        Case 1
            BufferName = "H22-H22"
        Case 2
            BufferName = "H33-H33"
        Case 3
            BufferName = "H44-H44"
        Case 4
            BufferName = "H35-H35"
        Case 5
            GF2 = 0
            GoTo KDataRemoval
    End Select

    CalibrationWindow.Activate
    Sheets("Na Cal").Select
    
        Range("A1").Select
    
        Cells.Find(What:=BufferName, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(3, -1) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
                    
        Do While GF2 < 5
            Select Case Cells(GF1, 4)
                Case 0
                    Rows(GF1 & ":" & GF1).Delete Shift:=xlUp
                    GF1 = GF1 - 1
                Case "x"
                    GF1 = GF1 - 1
                    Range("B" & GF1 & ":R" & GF1).Activate
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThick
                        End With
                    Range("A1").Select
                    GF2 = GF2 + 1
                    GoTo NaDataRemoval
            End Select
            GF1 = GF1 + 1
        Loop
        
KDataRemoval:

    Select Case GF2
        Case 0
            BufferName = "H31-H31"
        Case 1
            BufferName = "H22-H22"
        Case 2
            BufferName = "H33-H33"
        Case 3
            BufferName = "H44-H44"
        Case 4
            BufferName = "H35-H35"
        Case 5
            GF2 = 0
            GoTo ClDataRemoval
    End Select

    CalibrationWindow.Activate
    Sheets("K Cal").Select
    
        Range("A1").Select
    
        Cells.Find(What:=BufferName, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(3, -1) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
                    
        Do While GF2 < 5
            Select Case Cells(GF1, 4)
                Case 0
                    Rows(GF1 & ":" & GF1).Delete Shift:=xlUp
                    GF1 = GF1 - 1
                Case "x"
                    GF1 = GF1 - 1
                    Range("B" & GF1 & ":U" & GF1).Activate
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThick
                        End With
                    Range("A1").Select
                    GF2 = GF2 + 1
                    GoTo KDataRemoval
            End Select
            GF1 = GF1 + 1
        Loop

ClDataRemoval:

    Select Case GF2
        Case 0
            BufferName = "PL1-PL1"
        Case 1
            BufferName = "PL2-PL2"
        Case 2
            BufferName = "PL3-PL3"
        Case 3
            BufferName = "H22-H22"
        Case 4
            BufferName = "H33-H33"
        Case 5
            BufferName = "H44-H44"
        Case 6
            GF2 = 0
            NaCR = 0
            KCR = 0
            ClCR = 0
            GoTo NaSolver
    End Select

    CalibrationWindow.Activate
    Sheets("Cl Cal").Select
    
        Range("A1").Select
    
        Cells.Find(What:=BufferName, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(3, -1) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
                    
        Do While GF2 < 6
            Select Case Cells(GF1, 4)
                Case 0
                    Rows(GF1 & ":" & GF1).Delete Shift:=xlUp
                    GF1 = GF1 - 1
                Case "x"
                    GF1 = GF1 - 1
                    Range("B" & GF1 & ":Q" & GF1).Activate
                        With Selection.Borders(xlEdgeBottom)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .TintAndShade = 0
                            .Weight = xlThick
                        End With
                    Range("A1").Select
                    GF2 = GF2 + 1
                    GoTo ClDataRemoval
            End Select
            GF1 = GF1 + 1
        Loop

NaSolver:

    Sheets("Na Cal").Select
    

        Select Case GF2
            Case 0
                BufferName = "H31 Bias2"
            Case 1
                BufferName = "H22 Bias2"
            Case 2
                BufferName = "H33 Bias2"
            Case 3
                BufferName = "H44 Bias2"
            Case 4
                BufferName = "H35 Bias2"
            Case 5
                Application.ScreenUpdating = True
                Sheets("Na Cal").Select
                SolverSolve userFinish:=True
                Sheets("Na Cal").Select
                SolverSolve userFinish:=True
                GF2 = 0
                Range("A1").Select
                Application.ScreenUpdating = False
                GoTo NaCleanup
        End Select
        
        Range("A1").Select
    
        Cells.Find(What:=BufferName, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(1, -11) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
            
        Do While GF2 < 5
            If Range("L" & GF1) = "x" Then
                Sheets("Results Summary").Select
                    NaCR = 0
                Sheets("Na Cal").Select
                    GF2 = GF2 + 1
                    GoTo NaSolver
            End If
            InstrumentName = Range("B" & GF1)
            PanelID = Range("C" & GF1)
            DryInt = Range("D" & GF1)
            
            If DryInt <> 0 Then
            
                Sheets(InstrumentName).Select
                Range("A1").Select
                    Cells.Find(What:=PanelID, After:=ActiveCell, LookIn:= _
                        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                        xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    ActiveCell.Offset(0, 74).Select
                    IntInt = ActiveCell
                    Range("A1").Select
            
                Sheets("Na Cal").Select
                    If IntInt > (DryInt * 0.9) Then
                        Range("A1").Select
                    End If
            End If
            GF1 = GF1 + 1
        Loop
        
NaCleanup:

    Application.ScreenUpdating = False

        Select Case GF2
            Case 0
                BufferName = "H31 Bias2"
            Case 1
                BufferName = "H22 Bias2"
            Case 2
                BufferName = "H33 Bias2"
            Case 3
                BufferName = "H44 Bias2"
            Case 4
                BufferName = "H35 Bias2"
            Case 5
                SolverSolve userFinish:=True
                GF2 = 0
                Range("A1").Select
                GoTo NaBiasFineTune
        End Select
        
        Range("A1").Select
    
        Cells.Find(What:=BufferName, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(1, -11) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
            
        Do While GF2 < 5
            Select Case Range("L" & GF1)
                Case "No Data"
                    Range("J" & GF1 & ":K" & GF1 & ":L" & GF1 & ":M" & GF1).ClearContents
                    Range("N" & GF1) = "Undefined"
                Case "x"
                    Sheets("Results Summary").Select
                        Range("H" & (14 + GF2)) = NaCR
                    Sheets("Na Cal").Select
                        NaCR = 0
                    GF2 = GF2 + 1
                    GoTo NaCleanup
                Case Is >= NaBias
                    If Range("N" & GF1) <> "No Dispense" And Range("N" & GF1) <> "Undefined" Then
                        Range("J" & GF1 & ":K" & GF1 & ":L" & GF1 & ":M" & GF1).ClearContents
                        Range("N" & GF1) = "Outlier"
                        Range("B" & GF1 & ":N" & GF1).Font.Bold = True
                        Range("B" & GF1 & ":N" & GF1).Select
                            With Selection.Font
                                .Color = -16776961
                            End With
                        Range("A1").Select
                        NaCR = NaCR + 1
                    End If
            End Select
            GF1 = GF1 + 1
        Loop
        
NaBiasFineTune:
    Sheets("Results Summary").Select
    GF1 = 14

    Do While GF1 < 19
        If Range("H" & GF1) = 0 And Range("O" & GF1) = 1 Or Range("P" & GF1) = 1 Or Range("Q" & GF1) = 1 Then
            BufferName = Range("B" & GF1)
            LargestBias = BufferName + " Bias2"
            Range("H" & GF1) = 1
            Sheets("Na Cal").Select
            
            Range("A1").Select
    
            Cells.Find(What:=LargestBias, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            
            ActiveCell.Offset(1, -11) = "=ROW()"
            GF2 = Range("B1")
            Columns("A:A").ClearContents
            MaxBufBias = Range("W" & GF2 - 1)
    
            GF3 = 0
            
            Do While GF3 = 0
                Select Case Range("V" & GF2)
                    Case 1
                        Range("J" & GF2 & ":M" & GF2).ClearContents
                        Range("N" & GF2) = "Bias Outlier"
                        Range("B" & GF2 & ":N" & GF2).Font.Bold = True
                        Range("B" & GF2 & ":N" & GF2).Select
                            With Selection.Font
                                .Color = -16776961
                            End With
                        GF3 = 1
                End Select
                GF2 = GF2 + 1
            Loop
                    
            Sheets("Results Summary").Select
        End If
        GF1 = GF1 + 1
    Loop

    GF2 = 0
    Sheets("Na Cal").Select
    SolverSolve userFinish:=True
    Range("A1").Select

KSolver:

    Sheets("K Cal").Select
    

        Select Case GF2
            Case 0
                BufferName = "H31 Bias2"
            Case 1
                BufferName = "H22 Bias2"
            Case 2
                BufferName = "H33 Bias2"
            Case 3
                BufferName = "H44 Bias2"
            Case 4
                BufferName = "H35 Bias2"
            Case 5
                Sheets("K Cal").Select
                SolverSolve userFinish:=True
                Sheets("K Cal").Select
                SolverSolve userFinish:=True
                GF2 = 0
                Range("A1").Select
                GoTo KCleanup
        End Select
        
        Range("A1").Select
    
        Cells.Find(What:=BufferName, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(1, -14) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
            
        Do While GF2 < 5
            If Range("B" & GF1) = "x" Then
                Sheets("Results Summary").Select
                    KCR = 0
                Sheets("Na Cal").Select
                    GF2 = GF2 + 1
                    GoTo KSolver
            End If
            InstrumentName = Range("B" & GF1)
            PanelID = Range("C" & GF1)
            DryInt = Range("D" & GF1)
            
            If DryInt <> 0 Then
            
                Sheets(InstrumentName).Select
                Range("A1").Select
                    Cells.Find(What:=PanelID, After:=ActiveCell, LookIn:= _
                        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                        xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    Cells.Find(What:=PanelID, After:=ActiveCell, LookIn:= _
                        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                        xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    ActiveCell.Offset(0, 74).Select
                    IntInt = ActiveCell
                    Range("A1").Select
            
                Sheets("K Cal").Select
                    If IntInt > (DryInt * 0.8) Then
                        Range("A1").Select
                    End If
            End If
            GF1 = GF1 + 1
        Loop

KCleanup:

    Application.ScreenUpdating = False

        Select Case GF2
            Case 0
                BufferName = "H31 Bias2"
            Case 1
                BufferName = "H22 Bias2"
            Case 2
                BufferName = "H33 Bias2"
            Case 3
                BufferName = "H44 Bias2"
            Case 4
                BufferName = "H35 Bias2"
            Case 5
                SolverSolve userFinish:=True
                GF2 = 0
                Range("A1").Select
                GoTo KBiasFineTune
        End Select
        
        Range("A1").Select
    
        Cells.Find(What:=BufferName, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(1, -14) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
            
        Do While GF2 < 5
            Select Case Range("O" & GF1)
                Case "No Data"
                    Range("M" & GF1 & ",N" & GF1 & ",O" & GF1 & ",P" & GF1).ClearContents
                    Range("Q" & GF1) = "Undefined"
                Case "x"
                    Sheets("Results Summary").Select
                        Range("H" & (24 + GF2)) = Range("H" & (24 + GF2)) + KCR
                        KCR = 0
                    Sheets("K Cal").Select
                    GF2 = GF2 + 1
                    GoTo KCleanup
                Case Is >= KBias
                    If Range("Q" & GF1) <> "No Dispense" And Range("Q" & GF1) <> "Undefined" Then
                        Range("M" & GF1 & ",N" & GF1 & ",O" & GF1 & ",P" & GF1).ClearContents
                        Range("Q" & GF1) = "Outlier"
                        Range("B" & GF1 & ":Q" & GF1).Font.Bold = True
                        Range("B" & GF1 & ":Q" & GF1).Select
                            With Selection.Font
                                .Color = -16776961
                            End With
                        Range("A1").Select
                        KCR = KCR + 1
                    End If
            End Select
            GF1 = GF1 + 1
        Loop

KBiasFineTune:
    Sheets("Results Summary").Select
    GF1 = 24
    
    Do While GF1 < 29
        If Range("H" & GF1) = 0 And Range("O" & GF1) = 1 Or Range("P" & GF1) = 1 Or Range("Q" & GF1) = 1 Then
            BufferName = Range("B" & GF1)
            LargestBias = BufferName + " Bias2"
            Range("H" & GF1) = 1
            Sheets("K Cal").Select
            
            Range("A1").Select
            
            Cells.Find(What:=LargestBias, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            
            ActiveCell.Offset(1, -14) = "=ROW()"
            GF2 = Range("B1")
            Columns("A:A").ClearContents
            MaxBufBias = Range("Z" & GF2 - 1)
            
            GF3 = 0
            
            Do While GF3 = 0
                Select Case Range("Y" & GF2)
                    Case 1
                        Range("M" & GF2 & ":P" & GF2).ClearContents
                        Range("Q" & GF2) = "Bias Outlier"
                        Range("B" & GF2 & ":Q" & GF2).Font.Bold = True
                        Range("B" & GF2 & ":Q" & GF2).Select
                            With Selection.Font
                                .Color = -16776961
                            End With
                       GF3 = 1
                End Select
                GF2 = GF2 + 1
            Loop
            Sheets("Results Summary").Select
        End If
        GF1 = GF1 + 1
    Loop
    
    GF2 = 0
    Sheets("K Cal").Select
    SolverSolve userFinish:=True
    Range("A1").Select
    
ClSolver:

    Sheets("Cl Cal").Select

        Select Case GF2
            Case 0
                BufferName = "PL1 Bias2"
            Case 1
                BufferName = "PL2 Bias2"
            Case 2
                BufferName = "PL3 Bias2"
            Case 3
                BufferName = "H22 Bias2"
            Case 4
                BufferName = "H33 Bias2"
            Case 5
                BufferName = "H44 Bias2"
            Case 6
                Sheets("Cl Cal").Select
                SolverSolve userFinish:=True
                Sheets("Cl Cal").Select
                SolverSolve userFinish:=True
                GF2 = 0
                Range("A1").Select
                GoTo ClCleanup
        End Select
        
        Range("A1").Select
    
        Cells.Find(What:=BufferName, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(1, -10) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
            
        Do While GF2 < 6
            If Range("B" & GF1) = "x" Then
                Sheets("Results Summary").Select
                    ClCR = 0
                Sheets("Cl Cal").Select
                    GF2 = GF2 + 1
                    GoTo ClSolver
            End If
            InstrumentName = Range("B" & GF1)
            PanelID = Range("C" & GF1)
            DryInt = Range("D" & GF1)
            
            If Range("D" & GF1) <> 0 Then
            
            
                Sheets(InstrumentName).Select
                Range("A1").Select
                    Cells.Find(What:=PanelID, After:=ActiveCell, LookIn:= _
                        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                        xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    Cells.Find(What:=PanelID, After:=ActiveCell, LookIn:= _
                        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                        xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    Cells.Find(What:=PanelID, After:=ActiveCell, LookIn:= _
                        xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                        xlNext, MatchCase:=False, SearchFormat:=False).Activate
                    ActiveCell.Offset(0, 74).Select
                    IntInt = ActiveCell
                    Range("A1").Select
            
                Sheets("Cl Cal").Select
                    If IntInt > (DryInt * 0.8) Then
                        Range("A1").Select
                    End If
            End If
            GF1 = GF1 + 1
        Loop
        
ClCleanup:

    Application.ScreenUpdating = False

        Select Case GF2
            Case 0
                BufferName = "PL1 Bias2"
            Case 1
                BufferName = "PL2 Bias2"
            Case 2
                BufferName = "PL3 Bias2"
            Case 3
                BufferName = "H22 Bias2"
                ClBias = ClBias * 2
            Case 4
                BufferName = "H33 Bias2"
            Case 5
                BufferName = "H44 Bias2"
            Case 6
                Range("A1").Select
                SolverSolve userFinish:=True
                GF2 = 0
                GoTo ClBiasFineTune
        End Select
        
        Range("A1").Select
    
        Cells.Find(What:=BufferName, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            ActiveCell.Offset(1, -10) = "=ROW()"
            GF1 = Range("B1")
            Columns("A:A").ClearContents
            
        Do While GF2 < 6
            Select Case Range("K" & GF1)
                Case "No Data"
                    Range("I" & GF1 & ",J" & GF1 & ",K" & GF1 & ",L" & GF1).ClearContents
                    Range("M" & GF1) = "Undefined"
                Case "x"
                    Sheets("Results Summary").Select
                        Range("H" & (34 + GF2)) = Range("H" & (34 + GF2)) + ClCR
                        ClCR = 0
                    Sheets("Cl Cal").Select
                    GF2 = GF2 + 1
                    GoTo ClCleanup
                Case Is >= ClBias
                    If Range("M" & GF1) <> "No Dispense" And Range("M" & GF1) <> "Undefined" Then
                        Range("I" & GF1 & ",J" & GF1 & ",K" & GF1 & ",L" & GF1).ClearContents
                        Range("M" & GF1) = "Outlier"
                        Range("B" & GF1 & ":M" & GF1).Font.Bold = True
                        Range("B" & GF1 & ":M" & GF1).Select
                            With Selection.Font
                                .Color = -16776961
                            End With
                        Range("A1").Select
                        ClCR = ClCR + 1
                    End If
            End Select
            GF1 = GF1 + 1
        Loop
        
ClBiasFineTune:
    Sheets("Results Summary").Select
    GF1 = 34
    
    Do While GF1 < 37
        If Range("H" & GF1) = 0 And Range("O" & GF1) = 1 Or Range("P" & GF1) = 1 Or Range("Q" & GF1) = 1 Then
            BufferName = Range("B" & GF1)
            LargestBias = BufferName + " Bias2"
            Range("H" & GF1) = 1
            Sheets("Cl Cal").Select
            
            Range("A1").Select
    
            Cells.Find(What:=LargestBias, After:=ActiveCell, LookIn:= _
                    xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:= _
                    xlNext, MatchCase:=False, SearchFormat:=False).Activate
            
            ActiveCell.Offset(1, -10) = "=ROW()"
            GF2 = Range("B1")
            Columns("A:A").ClearContents
            MaxBufBias = Range("V" & GF2 - 1)
            
            GF3 = 0
            
            Do While GF3 = 0
                Select Case Range("U" & GF2)
                    Case 1
                        Range("I" & GF2 & ":L" & GF2).ClearContents
                        Range("M" & GF2) = "Bias Outlier"
                        Range("B" & GF2 & ":M" & GF2).Font.Bold = True
                        Range("B" & GF2 & ":M" & GF2).Select
                            With Selection.Font
                                .Color = -16776961
                            End With
                        GF3 = 1
                End Select
                GF2 = GF2 + 1
            Loop
            Sheets("Results Summary").Select
        End If
        GF1 = GF1 + 1
    Loop
    
    GF2 = 0
    Sheets("Cl Cal").Select
    SolverSolve userFinish:=True
    Range("A1").Select
        
IDEX_Cal:

    Sheets("Results Summary").Select
        Range("D2") = OperatorID
        Range("G2") = ExpDate
        Range("J2") = LotNo
        Range("M5") = Volume
    
    MacroWindow.Activate
    Sheets("Launch Macro").Select
        ActiveWorkbook.Unprotect Password:="81643"
        Sheets("Launch Macro").Unprotect Password:="81643"
        Range("H14:J17").Copy
        
    CalibrationWindow.Activate
    Sheets("Results Summary").Select
        Range("F42:H45").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
        
        Range("A1").Select
        Range("B4:K7").Copy
        
    Sheets("IDEXX").Select
        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
        Range("K1") = "'Lot No"
        Range("K2") = LotNo
        Range("A1").Select
        Range("A1") = "'Sensor Type"
        Sheets("IDEXX").Name = LotNo & "-Calibration"
        
        Columns("B:C").Select
        Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        Columns("B:B").Select
        Selection.ColumnWidth = 16.57
        Range("B1") = "f_aged_dry_ratio"
        Range("C1") = "fsensor_slope"
        Range("B2") = "='Na Cal'!D5"
        Range("C2") = "='Na Cal'!D4"
        Range("B3") = "='K Cal'!D6"
        Range("C3") = "='K Cal'!D4"
        Range("B4") = "='Cl Cal'!D6"
        Range("C4") = "='Cl Cal'!D4"
        Range("B2:C4").Copy
        Range("B2:C4").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
        Range("A1").Select
        
        Range("N2").Copy
        Range("N2").Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:= _
            xlNone, SkipBlanks:=False, Transpose:=False
                Application.CutCopyMode = False
        Range("A1").Select
        
        
        Range("O2") = (Na_WetS - 1)
        Range("O3") = (K_WetS - 1)
        Range("O4") = (Cl_WetS - 1)
    
    CalibrationWindow.Activate
    Sheets("Results Summary").Select
    
    SaveDocument = False
    
    Do While SaveDocument = False
        SaveDocument = Application.GetSaveAsFilename(LotNo & "-Calibration.xls", , , "SAVE Catalyst Dx Analysis; LOT: " & LotNo)
    Loop
    
    ActiveWorkbook.SaveAs FileName:=SaveDocument
    
    MacroWindow.Activate
    NumberOfInstruments = GrabNumberOfInstruments
    ReDim InstrumentList(NumberOfInstruments - 1) As String
    For i = 0 To (NumberOfInstruments - 1) 'cycle through instruments
        InstrumentList(i) = GrabInstrumentName(i)
    Next
    
    CalibrationWindow.Activate
    Sheets("Results Summary").Select
    
    For k = 0 To 2 'cycle through sensors
        If k = 0 Or k = 1 Then
            NoOfBuffers = 4
        ElseIf k = 2 Then
            NoOfBuffers = 2
        End If
        ColumnLetter = ExportAnalyteColumn(k)
        FlagRow = ExportAnalyteFlagRow(k)
        
        For j = 0 To NoOfBuffers 'cycle through buffers
            FlagFormula = "=SUM("
            RowNumber = ExportBufferRow(k, j)
            
            For i = 0 To (NumberOfInstruments - 1) 'cycle through instruments
                FlagFormula = FlagFormula & InstrumentList(i) & "!" _
                    & ColumnLetter & RowNumber
                
                If i = (NumberOfInstruments - 1) Then
                    FlagFormula = FlagFormula & ")"
                Else
                    FlagFormula = FlagFormula & ","
                End If
                
            Next 'cycle through instruments
            Range("I" & (FlagRow + j)) = FlagFormula
            
        Next 'cycle through buffers
    Next 'cycle through sensors
    
    
'Stop check
'j cycle through buffers
'k cycle through sensors
'=SUM(CDX1149!C35,CDX1225!C35,CDX1312!C35)
    
    MacroWindow.Activate
    Sheets("Macro Functions").Visible = False
    Sheets("Macro Controls").Visible = False
    ActiveWorkbook.Protect Password:="81643"
    Sheets("Launch Macro").Protect Password:="81643"
    
    CalibrationWindow.Activate
        ActiveWorkbook.Protect Password:="81643"
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
                CalibrationWindow.Save
                MsgBox "Catalyst Dx Analysis Complete"
                Exit Sub 'Macro is finished with no errors
            End If
            If InstrumentName <> Range("B" & GF1) Then
                InstrumentName = Range("B" & GF1)
                Sheets(InstrumentName).Protect Password:="81643"
            End If
            GF1 = GF1 + 1
        Loop
        
ErrorHandler:

    For Each wb In Application.Workbooks
        If wb.Name = TemplateFileName Then
            Workbooks(TemplateFileName).Close saveChanges:=False
        End If
    Next
    For Each wb In Application.Workbooks
        If wb.Name = LotNo & "_Barcode_" & RunNo & ".xls" Then
            Workbooks(LotNo & "_Barcode_" & RunNo & ".xls").Close saveChanges:=False
            Kill (DataDump & LotNo & "_Barcode_" & RunNo & ".xls")
        End If
    Next
    
    MacroWindow.Activate
    Sheets("Macro Functions").Visible = True
    Sheets("Macro Controls").Visible = True
    Sheets("Error Worksheet").Visible = True
    Sheets("Error Worksheet").Select
    
    Range("C3:C9,D3:D10").ClearContents
    
    Sheets("Macro Functions").Select
    
    IR = 10
    EC = 3
    
        Do While IR > 9 And IR < 19
            Select Case Range("C" & IR)
                Case True
                    InstrumentName = Range("D" & IR)
                    IR = IR + 1
                    Sheets("Error Worksheet").Select
                    Range("C" & EC) = InstrumentName
                    EC = EC + 1
                    Sheets("Macro Functions").Select
                Case False
                    IR = IR + 1
            End Select
        Loop
        
    BR = 10
    EC = 3
        
        Do While BR > 9 And BR < 18
            Select Case Range("F" & BR)
                Case True
                    BufferName = Range("G" & BR)
                    BR = BR + 1
                    Sheets("Error Worksheet").Select
                    Range("D" & EC) = BufferName
                    EC = EC + 1
                    Sheets("Macro Functions").Select
                Case False
                    BR = BR + 1
            End Select
        Loop
        
    Sheets("Error Worksheet").Select
    
    Range("E3") = RunNo
    Range("E6") = LotNo
    Range("E9") = OperatorID
    Range("G3") = ExpDate
    Range("E12") = SourceData
    Range("E13") = DataDump
    Range("E14") = FileName
    Range("F3") = NaBias
    Range("F6") = KBias
    Range("F9") = ClBias
    
        
    Sheets("Macro Functions").Visible = False
    Sheets("Macro Controls").Visible = False
    Sheets("Error Worksheet").Select
    ActiveWorkbook.Protect Password:="81643"
        
    
    MsgBox "An error has occured.  Please review your inputs as disagreement between .cvs file names or .cvs file locations and user inputs are the likely cause of the error." _
    , vbCritical
    

End Sub

Private Function GrabNumberOfInstruments() As Byte
    Sheets("Macro Functions").Select
    GrabNumberOfInstruments = Range("NumberOfInstruments").Value
End Function

Private Function GrabInstrumentName(ByVal InstrumentReferenceNumber As Byte) As String
Dim CountTrues As Byte
Dim TrueTarget As Byte
Dim Continue As Boolean
Dim i As Byte
    
    CountTrues = 0
    i = 10
    Continue = True
    TrueTarget = InstrumentReferenceNumber + 1
        
    Do While Continue = True
        If Range("C" & i).Value = True Then
            CountTrues = CountTrues + 1
        End If
        If CountTrues = TrueTarget Then
            Continue = False
            GrabInstrumentName = Range("D" & i).Value
        End If
        i = i + 1
    Loop

End Function

Private Function ExportBufferRow(ByVal AnalyteReferenceNumber As Byte, _
    ByVal BufferReferenceNumber As Byte) As String

    If AnalyteReferenceNumber = 0 Or AnalyteReferenceNumber = 1 Then
        Select Case BufferReferenceNumber
            Case 0
                ExportBufferRow = "35"
            Case 1
                ExportBufferRow = "15"
            Case 2
                ExportBufferRow = "55"
            Case 3
                ExportBufferRow = "95"
            Case 4
                ExportBufferRow = "75"
        End Select
    ElseIf AnalyteReferenceNumber = 2 Then
        Select Case BufferReferenceNumber
            Case 0
                ExportBufferRow = "115"
            Case 1
                ExportBufferRow = "135"
            Case 2
                ExportBufferRow = "155"
        End Select
    End If
End Function

Private Function ExportAnalyteColumn(ByVal AnalyteReferenceNumber As Byte) As String
    Select Case AnalyteReferenceNumber
        Case 0
            ExportAnalyteColumn = "B"
        Case 1
            ExportAnalyteColumn = "C"
        Case 2
            ExportAnalyteColumn = "D"
    End Select
End Function

Private Function ExportAnalyteFlagRow(ByVal AnalyteReferenceNumber As Byte) As Integer
    Select Case AnalyteReferenceNumber
        Case 0
            ExportAnalyteFlagRow = "14"
        Case 1
            ExportAnalyteFlagRow = "24"
        Case 2
            ExportAnalyteFlagRow = "34"
    End Select
End Function
