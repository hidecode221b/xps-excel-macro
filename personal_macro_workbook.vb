Option Explicit
    
    Dim highpe() As Variant, ratio() As Variant, bediff() As Variant, strl() As Variant, backSlash As String, wbpath As String, numRun As Integer
    
    Dim j As Integer, k As Integer, q As Integer, p As Integer, n As Integer, iRow As Integer, iCol As Integer, ns As Integer, fileNum As Integer
    Dim startR As Integer, endR As Integer, g As Integer, cae As Integer, ncomp As Integer, numXPSFactors As Integer, numAESFactors As Integer
    Dim numMajorUnit As Integer, modex As Integer, para As Integer, graphexist As Integer, numData As Integer, numChemFactors As Integer
    Dim idebug As Integer, spacer As Integer, sftfit As Integer, sftfit2 As Integer, cmp As Integer, scanNum As Integer, numGrant As Integer
    
    Dim wb As String, ver As String, TimeCheck As String, strAna As String, direc As String, ElemD As String, Results As String, testMacro As String
    Dim strSheetDataName As String, strSheetGraphName As String, strSheetFitName As String, strSheetAnaName As String, strBG1 As String, strBG2 As String
    Dim strMode As String, strLabel As String, strCasa As String, strAES As String, strErr As String, strErrX As String, ElemX As String, strBG3 As String
    
    Dim sheetData As Worksheet, sheetGraph As Worksheet, sheetFit As Worksheet, sheetAna As Worksheet
    Dim dataData As Range, dataKeData As Range, dataIntData As Range, dataBGraph As Range, dataKGraph, dataKeGraph As Range, dataBeGraph As Range
    
    Dim pe As Single, wf As Single, char As Single, off As Single, multi As Single, windowSize As Single, windowRatio As Single
    Dim startEk As Single, endEk As Single, startEb As Single, endEb As Single, stepEk As Single, dblMax As Single, dblMin As Single
    Dim chkMax As Single, chkMin As Single, gamma As Single, lambda As Single, maxXPSFactor As Single, maxAESFactor As Single
    Dim a0 As Single, a1 As Single, a2 As Single, fitLimit As Single, mfp As Single, peX As Single
    
Sub CLAM2()
    ver = "8.31p"                             ' Version of this code.
    backSlash = Application.PathSeparator ' Mac = "/", Win = "\"
    If backSlash = "/" Then    ' location of directory for database
        direc = backSlash + "Users" + backSlash + "apple" + backSlash + "Library" + backSlash + "Group Containers" + backSlash + "UBF8T346G9.Office" + backSlash + "MyExcelFolder" + backSlash + "XPS" + backSlash
        'direc = backSlash + "Users" + backSlash + "apple" + backSlash + "Documents" + backSlash + "XPS" + backSlash
    Else
        direc = "E:" + backSlash + "Data" + backSlash + "Hideki" + backSlash + "XPS" + backSlash
        ' "E:\Data\Hideki\XPS\"
    End If
    windowSize = 1.5          ' 1 for large, 2 for small display, and so on. Larger number, smaller graph plot.
    windowRatio = 4 / 3     ' window width / height, "2/1" for eyes or "4/3" for ppt
    ElemD = "C,O"           ' Default elements to be shown up in the element analysis.
    TimeCheck = "No"        ' "No" only iteration results in fitting, numeric value to suppress any display.
    
    a0 = -0.00044463        ' Undulator parameters for harmonics or
    a1 = 1.0975             ' B vs gap equation
    a2 = -0.02624           ' B = A0 + A1 * Exp(A2 * gap)
    gamma = 1.2             ' An electron energy: GeV
    lambda = 6              ' A magnetic period: cm
    fitLimit = 250          ' Maximum fit range: eV
    mfp = 0.6               ' Inelastic mean free path formula: E^(mfp), and mfp can be from 0.5 to 0.9.
    para = 100              ' position of parameters in the graph sheet with higher version of 6.56.
                            ' the limit of compared spectra depends on (para/3).
    spacer = 4              ' spacer between data tables for each parameter in FitRatioAnalysis, but it should be more than 3
    sftfit = 10
    sftfit2 = 5
    
    Call SheetNameAnalysis
    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    
    Call TargetDataAnalysis
End Sub

Sub SheetNameAnalysis()
    Dim FSO As Object, dt As Integer, C1 As Variant, rng As Range, sh As String, flag As Boolean, strTest As String

    If mid$(direc, Len(direc), 1) <> backSlash Then direc = direc & backSlash
    direc = Replace(direc, "/", backSlash)
    direc = Replace(direc, "*", "")
    
    If backSlash = "/" Then GoTo DeadInTheWater3

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.DriveExists(mid$(direc, 1, 2)) = False Then
        TimeCheck = MsgBox("Drive Not Found in " + mid$(direc, 1, 2) + " !" + vbCrLf + "Change a drive in direc", 0, "Database error")
        End
    End If
    
    If FSO.FolderExists(direc) Then
        If Len(Dir(direc + "UD.xlsx")) = 0 Then
            Application.DisplayAlerts = False
            Workbooks.Add
            Call UDsamples
            ActiveWorkbook.SaveAs fileName:=direc & "UD.xlsx", FileFormat:=xlOpenXMLWorkbook
            ActiveWorkbook.Close
            Application.DisplayAlerts = True
        End If
    Else
        If InStr(1, ActiveSheet.Name, "Fit_") > 0 Then
            TimeCheck = MsgBox("Database Not Found in " + direc + "!" + vbCrLf + "Would you like to continue?", 4, "Database error")
            If TimeCheck = 6 Then
                testMacro = "debug"
                ElemX = ""
                On Error GoTo DeadInTheWater1
                C1 = Split(direc, "\")
                
                For q = 1 To UBound(C1) - 1
                    C1(q) = C1(q - 1) & "\" & C1(q)
                    FSO.CreateFolder C1(q)
                Next q
                
                Workbooks.Add
                Call UDsamples
                ActiveWorkbook.SaveAs fileName:=direc & "UD.xlsx", FileFormat:=51
                ActiveWorkbook.Close
            Else
                End 'Call GetOut
DeadInTheWater1:
                MsgBox "A folder could not be created in the following path: " & direc & "." & vbCrLf & "Create directory manually and try again."
                End
            End If
        Else
            TimeCheck = MsgBox("Database Not Found in " + direc + "!" + vbCrLf + "Would you like to continue and create directory?", 4, "Database error")
            
            If FSO.DriveExists(mid$(direc, 1, 2)) = False Then
                End 'Call GetOut
            ElseIf TimeCheck = 6 Then
                On Error GoTo DeadInTheWater2
                C1 = Split(direc, "\")
                For q = 1 To UBound(C1) - 1
                    C1(q) = C1(q - 1) & "\" & C1(q)
                    FSO.CreateFolder C1(q)
                Next q
                
                Workbooks.Add
                Call UDsamples
                ActiveWorkbook.SaveAs fileName:=direc & "UD.xlsx", FileFormat:=xlOpenXMLWorkbook
                ActiveWorkbook.Close
            Else
                End 'Call GetOut
DeadInTheWater2:
                MsgBox "A folder could not be created in the following path: " & direc & "." & vbCrLf & "Create directory manually and try again."
                End
            End If
        End If
    End If
    Set FSO = Nothing
DeadInTheWater3:


    If StrComp(testMacro, "debug", 1) = 0 Then
        TimeCheck = 0
    End If
    
    Call Initial
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationAutomatic    ' revised for Office 2010
    graphexist = 0
    sh = ActiveSheet.Name
    wbpath = ActiveWorkbook.Path
    
    If InStr(1, sh, "Graph_") > 0 Then
        If InStr(1, sh, "Graph_Norm_") > 0 Then
            strSheetDataName = "Norm_" & mid$(sh, 12, (Len(sh) - 11))
        Else
            strSheetDataName = mid$(sh, 7, (Len(sh) - 6))
        End If
        graphexist = 1       ' for trigger for Graph sheet
        
        If IsEmpty(Cells(1, 2).Value) = False Then
            If IsNumeric(Cells(1, 2).Value) = True Then
                
            Else
                Exit Sub
            End If
        End If
        
        If IsEmpty(Cells(2, 2).Value) Then
            Cells(2, 2).Value = 40
        Else
            If IsNumeric(Cells(2, 2).Value) = False Then
                If StrComp(Cells(2, 2).Value, "MgKa", 1) = 0 Then
                    Cells(2, 2).Value = 1253.6
                ElseIf StrComp(Cells(2, 2).Value, "AlKa", 1) = 0 Then
                    Cells(2, 2).Value = 1486.6
                Else
                    Cells(2, 2).Value = 40
                End If
            Else
                If Cells(2, 2).Value > 1500 And (Cells(10, 2).Value = "Ke" Or Cells(10, 2).Value = "Be") Then
                    Cells(2, 2).Value = 1500
                End If
            End If
        End If
        
        pe = Cells(2, 2).Value
        highpe(0) = pe
        
        If IsEmpty(Cells(2, 3).Value) Then
            Cells(2, 3).Value = "eV"
        Else
            If StrComp(Cells(2, 3).Value, "eV", 1) <> 0 And StrComp(Cells(2, 1).Value, "PE", 1) = 0 Then
                Call HigherOrderCheck           ' Formula ";79;118.5;158 eV" in C2 cell
            End If
        End If
        
        If Cells(2, 1).Value = "PE" Then
            If IsEmpty(Cells(3, 2).Value) Then
                Cells(3, 2).Value = 4
            Else
                If IsNumeric(Cells(3, 2).Value) = False Then
                    Cells(3, 2).Value = 4
                End If
            End If
            wf = Cells(3, 2).Value
            
            If IsEmpty(Cells(4, 2).Value) Then
                Cells(4, 2).Value = 0
            Else
                If IsNumeric(Cells(4, 2).Value) = False Then
                    Cells(4, 2).Value = 0
                End If
            End If
            char = Cells(4, 2).Value
        ElseIf Cells(2, 1).Value = "KE shifts" Then ' AES mode
            wf = Cells(3, 2).Value
        Else
            wf = Cells(3, 2).Value
        End If
        
        If IsEmpty(Cells(9, 2).Value) Then
            Cells(9, 2).Value = 0
        Else
            If IsNumeric(Cells(9, 2).Value) = False Then
                Cells(9, 2).Value = 0
            End If
        End If
        off = Cells(9, 2).Value
        
        If IsEmpty(Cells(9, 3).Value) Then
            Cells(9, 3).Value = 1
        Else
            If IsNumeric(Cells(9, 3).Value) = False Then
                Cells(9, 3).Value = 1
            End If
        End If
        multi = Cells(9, 3).Value
        strAna = Cells(10, 3).Value
        
        If Cells(40, para + 9).Value = "Ver." Then
        Else
            For q = 1 To 500
                If StrComp(Cells(40, q + 9).Value, "Ver.", 1) = 0 Then Exit For
            Next
            para = q
        End If
        
        If IsEmpty(Cells(41, para + 12).Value) Then
            Cells(41, para + 12).Value = ((Cells(6, 2).Value - Cells(5, 2).Value) / Cells(7, 2).Value) + 1
        End If
        numData = Cells(41, para + 12).Value
        
        If IsEmpty(Cells(45, para + 10).Value) Then
            Cells(45, para + 10).Value = 0
        End If
        ncomp = Cells(45, para + 10).Value
        
        If StrComp(Cells(51, para + 9).Value, "FALSE", 1) = 0 Then
            Cells(51, para + 9).Value = "C,O"
        Else
            ElemD = Cells(51, para + 9).Value
        End If
        
        If StrComp(LCase(Cells(1, 1).Value), "exp", 1) = 0 Then
            strSheetAnaName = "Exp_" + strSheetDataName
            strSheetGraphName = "Graph_" + strSheetDataName
            Call ExportCmp("")
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf StrComp(LCase(Cells(1, 1).Value), "norm", 1) = 0 Or StrComp(LCase(Cells(1, 1).Value), "diff", 1) = 0 Then
            Call GetNormalize
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf StrComp(LCase(mid$(Cells(1, 1).Value, 1, 4)), "auto", 1) = 0 Then
            strSheetGraphName = "Graph_" + strSheetDataName
            Call GetAutoScale
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf StrComp(LCase(mid$(Cells(1, 1).Value, 1, 3)), "leg", 1) = 0 Then
            strSheetGraphName = "Graph_" + strSheetDataName
            Results = vbNullString
            Call CombineLegend
            End
        ElseIf StrComp(LCase(Cells(1, 1).Value), "debug", 1) = 0 Then
            Cells(1, 1).Value = "Grating"
            testMacro = "debugGraph"
            Call debugAll
            End
        ElseIf StrComp(LCase(Cells(1, 1).Value), "debugn", 1) = 0 Then
            Cells(1, 1).Value = "Grating"
            testMacro = "debugGraphn"
            Call debugAll
            End
        End If
        
        For k = 0 To CInt(para / 3)
            If StrComp(Cells(1, (4 + (3 * k))).Value, "comp", 1) = 0 Then Exit For
        Next
        
        If k >= CInt(para / 3) Then
            cmp = -1
        Else
            cmp = k     ' position of comp if cmp < ncomp
        End If          ' "cmp" should not be used because it preserves the starting point of comp function!
        
        g = 0
        If StrComp(strAna, "ana", 1) = 0 And StrComp(TimeCheck, "yes", 1) = 0 Then TimeCheck = vbNullString
    ElseIf InStr(1, sh, "Cmp_") > 0 Then
        strSheetDataName = mid$(sh, 5, (Len(sh) - 4))

        If StrComp(LCase(Cells(10, 3).Value), "chem", 1) = 0 Then
            Cells(10, 3).Value = "In-BG"
            strAna = "FitComp"
            
            strSheetAnaName = "Cmp_" + strSheetDataName
            strSheetGraphName = "Graph_" + strSheetDataName
            Set sheetGraph = Worksheets(strSheetGraphName)
            Set sheetAna = Worksheets(strSheetAnaName)
            
            sheetGraph.Activate
            numXPSFactors = Cells(43, para + 12).Value
            numChemFactors = Cells(42, para + 12).Value

            If IsEmpty(Cells(51, para + 10)) = False Then
                sheetGraph.Range(Cells(40, para + 9), Cells((Cells(51, para + 10).End(xlDown).Row), para + 30)).Copy Destination:=sheetAna.Cells(40, para + 9)
            End If
            
            Set sheetGraph = Worksheets(strSheetAnaName)
            sheetGraph.Activate
            If Cells(43, para + 12).Value <> numXPSFactors Then Call PlotElem
            If Cells(42, para + 12).Value <> numChemFactors Then Call PlotChem
            strErr = "skip"
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf StrComp(LCase(mid$(Cells(1, 1).Value, 1, 3)), "leg", 1) = 0 Then
            strSheetGraphName = "Cmp_" + strSheetDataName
            Results = vbNullString
            Call CombineLegend
            End
        ElseIf StrComp(LCase(mid$(Cells(1, 1).Value, 1, 4)), "auto", 1) = 0 Then
            strSheetGraphName = "Cmp_" + strSheetDataName
            ncomp = Cells(45, para + 10).Value
            Call GetAutoScale
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        Else
            strSheetAnaName = "Exc_" + strSheetDataName
            strSheetGraphName = "Cmp_" + strSheetDataName
            ncomp = Range(Cells(10, 1), Cells(10, 1).End(xlToRight)).Columns.Count / 3
            Call ExportCmp("")
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
        
        For k = 0 To CInt(para / 3)
            If StrComp(Cells(1, (4 + (3 * k))).Value, "comp", 1) = 0 Then Exit For
        Next
        
        If k >= CInt(para / 3) Then
            cmp = -1
        Else
            cmp = k     ' position of comp if cmp < ncomp
        End If          ' "cmp" should not be used because it preserves the starting point of comp function!
    ElseIf InStr(1, sh, "Fit_") > 0 Then
        If InStr(1, sh, "Fit_BE") > 0 Then
            strSheetDataName = Cells(1, 101).Value
        Else
            strSheetDataName = mid$(sh, 5, (Len(sh) - 4))
        End If
        
        wb = ActiveWorkbook.Name
        If Not ExistSheet("Graph_" + strSheetDataName) Then
            TimeCheck = MsgBox("Graph sheet " & "Graph_" + strSheetDataName & " is not found.", vbExclamation)
            End
        End If
        
        If Workbooks(wb).Sheets("Graph_" + strSheetDataName).Cells(40, para + 9).Value = "Ver." Then
        Else
            For q = 1 To 500
                If StrComp(Workbooks(wb).Sheets("Graph_" + strSheetDataName).Cells(40, q + 9).Value, "Ver.", 1) = 0 Then Exit For
            Next
            para = q
        End If
        
        If LCase(Cells(1, 4).Value) = "ana" And Cells(1, 1).Value <> "EF" Then
            Cells(1, 4).Value = "Name"
            Set rng = [A:A]
            numData = Application.CountA(rng) - 19
            startEb = Cells(6, 101).Value
            endEb = Cells(7, 101).Value
            dblMax = Cells(3, 101).Value
            dblMin = Cells(2, 101).Value
            Application.Calculation = xlCalculationManual
            Call FitAnalysis
            Application.Calculation = xlCalculationAutomatic
            Application.CutCopyMode = False
            Cells(1, 1).Select
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf LCase(Cells(1, 4).Value) = "debug" Then
            Cells(1, 4).Value = "Name"
            testMacro = "debugFit"
            Call debugAll
            Application.CutCopyMode = False
            End
        ElseIf LCase(Cells(1, 4).Value) = "debuga" Then
            Cells(1, 4).Value = "Name"
            testMacro = "debugShift"
            Call debugAll
            Application.CutCopyMode = False
            End
        ElseIf LCase(Cells(1, 4).Value) = "debugf" Then
            Cells(1, 4).Value = "Name"
            testMacro = "debugPara"
            Call debugAll
            Application.CutCopyMode = False
            End
        Else
            If InStr(1, sh, "Fit_BE") > 0 Then
                strMode = "Do fit range"
            Else
                strMode = "Do fit"
            End If
            Call FitCurve
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
    ElseIf InStr(1, sh, "Ana_") > 0 Then
        strSheetDataName = mid$(sh, 5, (Len(sh) - 4))
        wb = ActiveWorkbook.Name
        
        If StrComp(Cells(1, para + 1).Value, "Parameters", 1) = 0 Then
        Else
            For q = 1 To 500
                If Cells(1, q).Value = "Parameters" Then
                    Exit For
                ElseIf q = 500 Then
                    MsgBox "Ana sheet has no parameters to be compared."
                    End
                End If
            Next
            para = q
        End If

        Call FitRatioAnalysis
        
        Application.CutCopyMode = False
        End
    Else
        If InStr(ActiveWorkbook.Name, ".") < 1 Then
            Application.Dialogs(xlDialogSaveAs).Show
        End If
		
		strTest = mid$(ActiveWorkbook.Name, 1, InStrRev(ActiveWorkbook.Name, ".") - 1)
		strTest = mid$(strTest, 1, 25)
        
        flag = False
        For Each sheetData In Worksheets
            If sheetData.Name = strTest Then flag = True
        Next sheetData
        
        If flag = True Then
            ActiveSheet.Name = mid$(sh, 1, 25)
            strSheetDataName = mid$(sh, 1, 25)
        Else
            ActiveSheet.Name = strTest
            strSheetDataName = strTest
        End If
        
        strCasa = "User Defined"  ' default database for XPS  "VG Avt"/"SC CXRO"
        strAES = "User Defined"   ' default database for AES  "VG qnt"/"Mrocz"
    End If
    
    strSheetGraphName = "Graph_" + strSheetDataName
    strSheetFitName = "Fit_" + strSheetDataName
    
    If Not ExistSheet(strSheetDataName) Then
        TimeCheck = MsgBox("Data sheet " & strSheetDataName & " is not found.", vbExclamation)
        End
    End If

    Set sheetData = Worksheets(strSheetDataName)
    Worksheets(strSheetDataName).Activate
    
	Application.DisplayAlerts = False
	If Len(ActiveWorkbook.Path) < 2 Then
		Application.Dialogs(xlDialogSaveAs).Show
	Else
		On Error GoTo Error1
'            If backSlash = "/" And numRun = 1 Then
'                filePermissionCandidates = Array(wbpath, ActiveWorkbook.FullName, wbpath & backSlash & wb)
'                fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
'            End If
		ActiveWorkbook.SaveAs fileName:=wbpath + backSlash + wb, FileFormat:=xlOpenXMLWorkbook
	End If
	Application.DisplayAlerts = True
	Exit Sub
Error1:
    Err.Clear
End Sub

Sub TargetDataAnalysis()
    strMode = Cells(1, 1).Value

    If InStr(strMode, "E/eV") > 0 Then          ' Manually imported data analsysis
        Do
            If InStr(strMode, "'") > 0 Then     ' remove "'" generated in Igor produced text
                q = InStr(strMode, "'")
                strMode = Left$(strMode, q - 1) + mid$(strMode, q + 1)	
            Else
                Cells(1, 1).Value = strMode
                Exit Do
            End If
        Loop
        
        If InStr(Cells(1, 3).Value, "E/eV") > 0 Then
            Call Convert2Txt("")
            TimeCheck = MsgBox("Data were exported in the text files.", vbExclamation)
        End If
        
        If cmp >= 0 Then
            Call GetCompare
        ElseIf StrComp(strAna, "ana", 1) = 0 Then
            Call FitCurve
        ElseIf StrComp(strAna, "chem", 1) = 0 Then
            Call PlotChem
        Else
            Call KeBL            ' KE, BE, PE, GE, QE, AE, ME/eV data setup
            
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            
            If StrComp(strMode, "GE/eV", 1) = 0 Then        ' Grating scan with fixed gap
                Call EngBL
                Call descriptHidden1
                Call GetOut
            Else
                Call PlotCLAM2
                If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
                Call ElemXPS
                If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
                Call PlotElem
                If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
                Call FitCurve
            End If
        End If
    Else
        strMode = mid$(Cells(2, 1).Value, 1, 5)
        If cmp >= 0 Then
            Call GetCompare
        Else
            Call FormatData
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call PlotCLAM2
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call ElemXPS
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call PlotElem
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call FitCurve
        End If

        If StrComp(TimeCheck, "yes", 1) = 0 Then TimeCheck = "yes2"
        Call GetOut
    End If
End Sub

Sub PlotCLAM2()
    Dim C1 As Variant, C2 As Variant, C3 As Variant, C4 As Variant, imax As Integer, sig As Integer, SourceRangeColor1 As Long, SourceRangeColor2 As Long, strTest As String
    
    sig = 1
    imax = numData + 10
    If ExistSheet(strSheetGraphName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetGraphName).Delete
        Application.DisplayAlerts = True
    End If

    If ExistSheet(strSheetFitName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetFitName).Delete
        Application.DisplayAlerts = True
    End If
    
    Worksheets.Add().Name = strSheetGraphName
    Set sheetGraph = Worksheets(strSheetGraphName)
    sheetGraph.Activate
    
    Set dataBGraph = Range(Cells(11, 2), Cells(11, 2).Offset(numData - 1, 1))
    Set dataKGraph = Union(Range(Cells(11, 1), Cells(11, 1).Offset(numData - 1, 0)), Range(Cells(11, 3), Cells(11, 3).Offset(numData - 1, 0)))
    Set dataKeGraph = Range(Cells(11, 1), Cells(11, 1).Offset(numData - 1, 0))
    Set dataBeGraph = dataKeGraph.Offset(, 1)
    dataKeGraph.Value = dataKeData.Value
    C1 = dataKeData      ' C first column
    C2 = dataIntData     ' U second column
    C3 = dataKeGraph.Offset(, 2)    ' dataIntGraph    ' A
    
    If StrComp(strMode, "AE/eV", 1) = 0 Then
        C4 = Range(Cells(1, para + 1), Cells(1, para + 5))
        For n = 1 To numData
            startR = n - 1
            If (n - 1) < 1 Then startR = 1
            endR = n + wf
            If (n + wf) > numData Then endR = numData
            For j = 1 To 5
                C4(1, j) = 0
            Next
            
            For j = startR To endR
                C4(1, 1) = C4(1, 1) + C1(j, 1) * C1(j, 1)
                C4(1, 2) = C4(1, 2) + C1(j, 1)
                C4(1, 3) = C4(1, 3) + 1
                C4(1, 4) = C4(1, 4) + C1(j, 1) * C2(j, 1)
                C4(1, 5) = C4(1, 5) + C2(j, 1)
            Next
            
            C3(n, 1) = (C4(1, 3) * C4(1, 4) - C4(1, 2) * C4(1, 5)) / (C4(1, 1) * C4(1, 3) - C4(1, 2) * C4(1, 2))
            Range(Cells(11, 2), Cells((numData + 10), 2)) = C2
        Next
    ElseIf InStr(strMode, "E/eV") > 0 Then
        If StrComp(Cells(1, 3).Value, "Ip", 1) = 0 Or StrComp(Cells(1, 3).Value, "Ie", 1) = 0 Then
            C4 = dataKeData.Offset(, 2)
        Else
            C4 = dataKeData.Offset(, para + 30)      ' Empty Ip
        End If
        
        For n = 1 To numData
            If IsEmpty(C4(n, 1)) Then
                C4(n, 1) = 1
            Else
                If IsNumeric(C3(n, 1)) = False Then
                    C4(n, 1) = 1
                Else
                    If C4(n, 1) <= 0 Then
                        C4(n, 1) = 1
                    End If
                End If
            End If
            
            C3(n, 1) = (C2(n, 1) / C4(n, 1))
        Next
    Else
        If WorksheetFunction.Average(dataIntData) < 0 Then sig = -1
        For n = 1 To numData
            If IsNumeric(C2(n, 1)) = False Then Exit For
            C3(n, 1) = C2(n, 1) * sig * 1
        Next
    End If

    Range(Cells(11, 3), Cells((numData + 10), 3)) = C3
    
    Call descriptGraph
    Call scalecheck
    
    If strMode = "ME/eV" Then Call SheetCheckGenerator      ' Check Sheet for "ME/eV"

    If numMajorUnit > 0 Then
        If startEk > 0 Then
            startEk = Application.Floor(startEk, numMajorUnit)
        Else
            startEk = Application.Ceiling(startEk, (-1 * numMajorUnit))
        End If
    
        If endEk > 0 Then
            endEk = Application.Ceiling(endEk, numMajorUnit)
        Else
            endEk = Application.Floor(endEk, (-1 * numMajorUnit))
        End If
    End If
    
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers 'xlXYScatterSmoothNoMarkers
    ActiveChart.SetSourceData Source:=dataBGraph, PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetGraphName
    ActiveChart.SeriesCollection(1).Name = ActiveWorkbook.Name  '"BE graph"
    ActiveChart.ChartTitle.Delete
    
    With ActiveChart.Axes(xlCategory, xlPrimary)
        If StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(3), "De", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then
            .MinimumScale = startEb
            .MaximumScale = endEb
            .Crosses = xlMinimum
        Else
            .MinimumScale = endEb
            .MaximumScale = startEb
            .ReversePlotOrder = True
            .Crosses = xlMaximum
        End If
        .HasTitle = True
        .AxisTitle.Text = strl(0)
    End With
    
    SourceRangeColor1 = ActiveChart.SeriesCollection(1).Border.Color
    
    With ActiveSheet.ChartObjects(1)
        .Top = 20
    End With

    If StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(1), "Be", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then GoTo SkipGraph2
    
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers 'xlXYScatterSmoothNoMarkers
    ActiveChart.SetSourceData Source:=dataKGraph, PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetGraphName
    ActiveChart.SeriesCollection(1).Name = ActiveWorkbook.Name  '"KE graph"
    ActiveChart.ChartTitle.Delete

    With ActiveChart.Axes(xlCategory, xlPrimary)
        .MinimumScale = startEk
        .MaximumScale = endEk
        .HasTitle = True
        .AxisTitle.Text = "Kinetic energy (eV)"
    End With

    ActiveChart.SeriesCollection(1).Border.ColorIndex = 22
    SourceRangeColor2 = ActiveChart.SeriesCollection(1).Border.Color

    Range(Cells(10, 1), Cells(10, 1)).Interior.Color = SourceRangeColor2
    Range(Cells(9 + (imax), 1), Cells(9 + (imax), 1)).Interior.Color = SourceRangeColor2
            
    With ActiveSheet.ChartObjects(2)
        .Top = 1 * (500 / windowSize) + 20
    End With
    
SkipGraph2:
    
    Dim myChartOBJ As ChartObject
    For Each myChartOBJ In ActiveSheet.ChartObjects
        With myChartOBJ
            .Left = 200
            .Width = (550 * windowRatio) / windowSize
            .Height = 500 / windowSize
            '.Chart.Legend.Delete
        End With
        With myChartOBJ.Chart.Axes(xlCategory, xlPrimary)
            .MinorTickMark = xlOutside
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .HasMajorGridlines = True
            If numMajorUnit <> 0 Then
                .MajorUnit = numMajorUnit
            Else
                .MinimumScaleIsAuto = True
                .MaximumScaleIsAuto = True
            End If
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With myChartOBJ.Chart.Axes(xlValue)
            If StrComp(strl(3), "De", 1) = 0 Then
                .HasTitle = True
                .AxisTitle.Text = "Intensity (arb. units)"
                .Crosses = xlMinimum
            Else
                .HasTitle = True
                If InStr(strMode, "E/eV") > 0 Then
                    If sheetData.Cells(1, 2).Value = "AlKa" Then
                        .AxisTitle.Text = "K counts per sec."
                    Else
                        .AxisTitle.Text = "Intensity (arb. units)"
                    End If
                Else
                    .AxisTitle.Text = "Intensity normalized by Ip (arb. units)"
                End If
            End If
            If dblMin <> dblMax Then
                .MinimumScale = dblMin
                .MaximumScale = dblMax
            Else
                .MinimumScaleIsAuto = True
                .MaximumScaleIsAuto = True
            End If
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With myChartOBJ.Chart.Legend
            .Position = xlLegendPositionRight
            .IncludeInLayout = True
            .Left = (850 / windowSize)
            '.Width = 100
            '.Height = 100
            .Top = (50 / windowSize)
            With .Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 255, 255)
                .ForeColor.TintAndShade = 0.1
            End With
        End With
        With myChartOBJ.Chart
            '.PlotArea.Height = ((500 - 40) / windowSize)
            .PlotArea.Width = (((550 * windowRatio) - 40) / windowSize)
            .ChartArea.Border.LineStyle = 0
            '.ChartArea.Interior.ColorIndex = xlNone    'transparent plot
        End With
    Next
    
    If StrComp(strl(3), "De", 1) = 0 Then
        ActiveSheet.ChartObjects(2).Activate
        With ActiveChart.Axes(xlValue)
            .MinimumScale = chkMin
            .MaximumScale = chkMax
        End With
    End If
    
    Range(Cells(10, 2), Cells(10, 2)).Interior.Color = SourceRangeColor1

    If StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(1), "Be", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then
        Range(Cells(10, 1), Cells(10, 1)).Interior.Color = SourceRangeColor1
    End If

    Range(Cells(9 + (imax), 2), Cells(9 + (imax), 2)).Interior.Color = SourceRangeColor1
    strTest = mid$(strSheetGraphName, InStr(strSheetGraphName, "_") + 1, Len(strSheetGraphName) - 6)
    Cells(8 + (imax), 2).Value = strTest + ".xlsx"
    Cells(9 + (imax), 1).Value = strl(1) + strTest
    Cells(9 + (imax), 2).Value = strl(2) + strTest
    Cells(9 + (imax), 3).Value = strl(3) + strTest
    
    If ExistSheet("Sort_" & strSheetDataName) Then
        Application.DisplayAlerts = False
        Worksheets("Sort_" & strSheetDataName).Delete
        Application.DisplayAlerts = True
    End If
    
    If strl(3) = "Pp" Then strErr = "skip"
End Sub

Sub ElemXPS()
    Dim xpsoffset As Integer, aesoffset As Integer, asf As String, oriXPSFactors As Integer, rtoe As Single
    Dim Fname As Variant, Record As Variant, C1 As Variant, C2 As Variant, C3 As Variant, Elem As String, strTest As String
    
    xpsoffset = 0
    
CheckElemAgain:

    If StrComp(testMacro, "debug", 1) = 0 Then
        ElemD = ElemX
    Else
        ElemD = Application.InputBox(Title:="Input atomic elements", Prompt:="Example:C,O,Co,etc ... without space!", Default:=ElemD, Type:=2)
    End If
    
    If ElemD <> "False" Then
        If ElemD = "" Then  ' when you click "OK" without any element in box
			Call descriptHidden2
            Call FitCurve
            strErr = "skip"
            Exit Sub
        End If
    Else        ' when you click "cancel"
        strErr = "skip"
        Call GetOut
        Exit Sub
    End If
    
    n = 0
    j = 0
    k = 0
    q = 0
    
    Fname = direc + "UD.xlsx"
    xpsoffset = 2
    strCasa = "User Defined"
    
    If Not WorkbookOpen("UD.xlsx") Then
        graphexist = 0
        Workbooks.Open Fname
        Workbooks("UD.xlsx").Activate
        If Err.Number > 0 Then
            MsgBox "Error in " & Fname, vbOKOnly, "Error code: " & Err.Number
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf StrComp(ActiveWorkbook.Name, "UD.xlsx", 1) <> 0 Then
            MsgBox "Error in " & Fname
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
    Else
        Workbooks("UD.xlsx").Activate
        graphexist = 1
    End If

    If ExistSheet("XPS") Then
        Workbooks("UD.xlsx").Sheets("XPS").Activate
        iRow = ActiveSheet.UsedRange.Rows.Count
        If iRow = 0 Then iRow = 1
        
        C2 = Range(Cells(1, 1), Cells(1, 1).Offset(iRow - 1, 3)) '
        
        If mid$(Cells(1, 4).Value, 1, 1) = "R" Then
            asf = "RSF"  ' Relative Sensitivity factors
        ElseIf mid$(Cells(1, 4).Value, 1, 1) = "A" Then
            asf = "ASF"  ' Absolute Sensitivity factors: no PI cross-section normalization
        ElseIf mid$(Cells(1, 4).Value, 1, 1) = "P" Then
            asf = "PSF"  ' Photo-ionization Sensitivity factors : ignore database, use WebCross data only
        Else
            asf = "ASF"
        End If
        
        If graphexist = 0 Then
            Workbooks("UD.xlsx").Close False
        End If
        sheetGraph.Activate
    Else
        If graphexist = 0 Then
            Workbooks("UD.xlsx").Close False
        End If
        Call GetOut
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    
    If iRow < 2 Then
        numXPSFactors = 0
        strErrX = "skip"
        Exit Sub
    End If
    
    C1 = C2
    ReDim C2(1 To iRow, 1 To 11)
    k = 0
    C3 = Split(ElemD, ",")
    
    For n = 0 To UBound(C3)
        Elem = C3(n)

        If Elem = "AL" Then
            Elem = "Na,K,Rb,Cs"
        ElseIf Elem = "EA" Then
            Elem = "Be,Mg,Ca,Sr,Ba,Ra"
        ElseIf Elem = "TM" Then
            Elem = "Sc,Ti,V,Cr,Mn,Fe,Co,Ni,Cu,Zn,Y,Zr,Nb,Mo,Tc,Ru,Rh,Pd,Ag,Cd,Lu,Hf,Ta,W,Re,Os,Ir,Pt,Au,Hg"
        ElseIf Elem = "3d" Then
            Elem = "Sc,Ti,V,Cr,Mn,Fe,Co,Ni,Cu,Zn"
        ElseIf Elem = "4d" Then
            Elem = "Y,Zr,Nb,Mo,Tc,Ru,Rh,Pd,Ag,Cd"
        ElseIf Elem = "5d" Then
            Elem = "Lu,Hf,Ta,W,Re,Os,Ir,Pt,Au,Hg"
        ElseIf Elem = "SM" Then
            Elem = "B,Si,Ge,As,Sb,Te"
        ElseIf Elem = "NM" Then
            Elem = "C,N,O,P,S,Se"
        ElseIf Elem = "BM" Then
            Elem = "Al,Ga,In,Sn,Tl,Pb,Bi"
        ElseIf Elem = "HA" Then
            Elem = "F,Cl,Br,I,At"
        ElseIf Elem = "NG" Then
            Elem = "Ne,Ar,Kr,Xe,Rn"
        ElseIf Elem = "RM" Then
            Elem = "La,Ce,Nd,Sm,Eu,Gd,Tb,Er,Tm,Yb,Th,U"
        ElseIf Elem = "LA" Then
            Elem = "La,Ce,Nd,Sm,Eu,Gd,Tb,Er,Tm,Yb"
        ElseIf Elem = "AC" Then
            Elem = "Th,U"
        Else
            k = 1
        End If
        
        If k = 0 Then
            ElemD = Replace(ElemD, C3(n), Elem)
        End If
        k = 0
    Next

    C3 = Split(ElemD, ",")
    
    k = 0
    For n = 0 To UBound(C3)
        Elem = C3(n)
        For p = 1 To Len(Elem)
            If IsNumeric(mid$(Elem, p, 1)) Then
                If IsNumeric(mid$(Elem, p, Len(Elem))) Then
                    rtoe = mid$(Elem, p, Len(Elem))
                Else
                    If StrComp(testMacro, "debug", 1) = 0 Then  ' debugAll code needs this
                        Call GetOut
                        strErrX = "skip"
                        Exit Sub
                    Else
                        TimeCheck = MsgBox(Elem + " : No such an element in database!", vbExclamation, "Input error")
                        GoTo CheckElemAgain
                    End If
                End If
                Elem = mid$(Elem, 1, p - 1)
                Exit For
            Else
                rtoe = 1
            End If
        Next
        j = 1 + k
        For q = 1 To (iRow)
            If C1(q, 1) = Elem Then
                C2(j, 1) = C1(q, 1)   ' Elem
                C2(j, 2) = C1(q, 2)   ' orbit
                C2(j, 3) = C1(q, 3)   ' BE
                C2(j, 7) = C1(q, 6 - xpsoffset) ' RSF
                C2(j, 11) = rtoe                ' atomic ratio
                j = j + 1
            ElseIf LCase(Elem) = "all" And q > 1 Then
                C2(j, 1) = C1(q, 1)   ' Elem
                C2(j, 2) = C1(q, 2)   ' orbit
                C2(j, 3) = C1(q, 3)   ' BE
                C2(j, 7) = C1(q, 6 - xpsoffset) ' RSF
                C2(j, 11) = rtoe                ' atomic ratio
                j = j + 1
            End If
        Next
        
        If j = 1 + k Then
            If Elem = vbNullString Then
            Else
                If StrComp(testMacro, "debug", 1) = 0 Then  ' debugAll code needs this
                    Call GetOut
                    strErrX = "skip"
                    Exit Sub
                Else
                    TimeCheck = MsgBox(Elem + " : No such an element in database!", vbExclamation, "Input error")
                    GoTo CheckElemAgain
                End If
            End If
        End If
        
        k = j - 1
    Next
    
    numXPSFactors = k
    If numXPSFactors = 0 Then GoTo SkipXPSnumZero
    
    maxXPSFactor = 0
    ReDim C3(1 To numXPSFactors, 1 To 7)
    
    For n = 1 To numXPSFactors
        strTest = C2(n, 1) + Left$(C2(n, 2), 2)
        C3(n, 1) = strTest
        
        If Dir(direc + "webCross" + backSlash) = vbNullString Then
            q = 0
            GoTo SkipElem
        End If
        
        Fname = direc + "webCross" + backSlash + LCase(strTest) + ".txt"
        
        If Dir(Fname) = vbNullString Then
            TimeCheck = MsgBox("File Not Found in " + Fname + "!", vbExclamation, "Database error")
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
        
        If Fname = False Then Exit Sub
        
        fileNum = FreeFile(0)
        Open Fname For Input As #fileNum
        iRow = 1
        q = 0
        
        Do
            Line Input #fileNum, Record
            C1 = Split(Record, vbTab)

            If strl(1) = "Pe" Then         ' XAS mode
                If C2(n, 3) < 15 Then    ' if PE < 15 eV, ignore it.
                    
                ElseIf C2(n, 3) > 15 And C1(0) >= C2(n, 3) And q = 0 And C2(n, 3) <> 1486.6 Then
                    C3(n, 2) = C1(0)      ' PE
                    C3(n, 3) = C1(1)      ' Cross section at PE
                    C3(n, 6) = C1(4)  ' asymmetric parameter
                    q = 1
                ElseIf C2(n, 3) > 15 And C1(0) >= C2(n, 3) And q = 0 And C2(n, 3) = 1486.6 Then
                    C3(n, 2) = C1(0)      ' PE
                    C3(n, 3) = C1(1)      ' Cross section at PE
                    C3(n, 4) = C1(0)      ' Al Ka PE
                    C3(n, 5) = C1(1)      ' Cross section at Al Ka
                    C3(n, 6) = C1(4)  ' asymmetric parameter
                    q = 1
                ElseIf C1(0) = 1486.6 Then
                    C3(n, 4) = C1(0)      ' Al Ka PE
                    C3(n, 5) = C1(1)      ' Cross section at Al Ka
                End If
            Else
                If C1(0) >= pe And q = 0 And C1(0) <> 1486.6 Then
                    C3(n, 2) = C1(0)
                    C3(n, 3) = C1(1)
                    C3(n, 6) = C1(4)  ' asymmetric parameter
                    q = 1
                ElseIf C1(0) >= pe And q = 0 And C1(0) = 1486.6 Then
                    C3(n, 2) = C1(0)
                    C3(n, 3) = C1(1)
                    C3(n, 4) = C1(0)
                    C3(n, 5) = C1(1)
                    C3(n, 6) = C1(4)  ' asymmetric parameter
                    q = 1
                ElseIf C1(0) = 1486.6 Then
                    C3(n, 4) = C1(0)
                    C3(n, 5) = C1(1)
                End If
            End If
            
            iRow = iRow + 1
        Loop Until EOF(fileNum)
        
        Close #fileNum
        
SkipElem:
       
        If q = 0 Or StrComp(asf, "ASF", 1) = 0 Then
            C3(n, 2) = 0
            C3(n, 3) = 1        ' if no data in webcross, multiply this factor !
            C3(n, 4) = 0
            C3(n, 5) = 1
            C3(n, 6) = 1
        End If
    Next
    
    For n = 1 To numXPSFactors
        C2(n, 2) = C2(n, 1) + C2(n, 2)
        If C2(n, 7) = "NaN" Or C2(n, 7) = vbNullString Then
            If q = 0 Then
                C2(n, 7) = 0
            Else
                C2(n, 7) = C3(n, 3)
            End If
        ElseIf StrComp(asf, "PSF", 1) = 0 Then
            C2(n, 7) = C3(n, 3)       ' if no RSF available, use cross section as a RSF.
        Else
            C2(n, 7) = C2(n, 7) * C3(n, 3) / C3(n, 5)
        End If
        
        C2(n, 10) = C3(n, 6)
    Next
    
    For n = 1 To numXPSFactors
        If C2(n, 7) >= maxXPSFactor Then maxXPSFactor = C2(n, 7) Else maxXPSFactor = maxXPSFactor
    Next
    
    If Abs(startEb - endEb) > fitLimit Then
        maxXPSFactor = maxXPSFactor * 2
    Else
        maxXPSFactor = maxXPSFactor * 1.2
    End If
    
    For n = 1 To numXPSFactors
        C2(n, 8) = dblMin + (C2(n, 11) * C2(n, 7) * ((dblMax - dblMin) / (maxXPSFactor)))
        If C2(n, 7) = 0 Then
            C2(n, 8) = vbNullString
        End If
    Next
    
    Range(Cells(51, para + 10), Cells((numXPSFactors + 50), para + 20)) = C2
            
    If StrComp(Cells(2, 1).Value, "PE", 1) = 0 Then
        If UBound(highpe) > 0 Then      ' higher order or ghost effects
            For n = 1 To UBound(highpe)
                Range(Cells(51 + numXPSFactors * (n), para + 10), Cells((50 + numXPSFactors * (n + 1)), para + 19)) = C2
                Cells(40 + n, para + 13).Value = "pe" & n
                Cells(40 + n, para + 14).Value = highpe(n)
            Next
            oriXPSFactors = numXPSFactors
            numXPSFactors = (UBound(highpe) + 1) * numXPSFactors
        End If
    End If
            
SkipXPSnumZero:
    
    If strl(1) = "Pe" Then Exit Sub

    aesoffset = 0
    
    Fname = direc + "UD.xlsx"
    strAES = "User Defined"
  
    If Not WorkbookOpen("UD.xlsx") Then
        graphexist = 0
        Workbooks.Open Fname
        Workbooks("UD.xlsx").Activate
        If Err.Number > 0 Then
            MsgBox "Error in " & Fname, vbOKOnly, "Error code: " & Err.Number
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf StrComp(ActiveWorkbook.Name, "UD.xlsx", 1) <> 0 Then
            MsgBox "Error in " & Fname
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
    Else
        Workbooks("UD.xlsx").Activate
        graphexist = 1
    End If
    
    If ExistSheet("AES") Then
        Workbooks("UD.xlsx").Sheets("AES").Activate
        iRow = ActiveSheet.UsedRange.Rows.Count
        If iRow = 0 Then iRow = 1
        C2 = Range(Cells(1, 1), Cells(1, 1).Offset(iRow - 1, 3 + aesoffset))
        
        If graphexist = 0 Then
            Workbooks("UD.xlsx").Close False
        End If
        sheetGraph.Activate
    Else
        If graphexist = 0 Then
            Workbooks("UD.xlsx").Close False
        End If
        Call GetOut
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    
    If iRow < 2 Then
        numAESFactors = 0
        strErrX = "skip"
        Exit Sub
    End If
    
    C1 = C2
    ReDim C2(1 To iRow, 1 To 11)
    C3 = Split(ElemD, ",")
    k = 0
    
    For n = 0 To UBound(C3)
        Elem = C3(n)
        For p = 1 To Len(Elem)
            If IsNumeric(mid$(Elem, p, 1)) Then
                If IsNumeric(mid$(Elem, p, Len(Elem))) Then
                    rtoe = mid$(Elem, p, Len(Elem))
                Else
                    If StrComp(testMacro, "debug", 1) = 0 Then  ' debugAll code needs this
                        Call GetOut
                        strErrX = "skip"
                        Exit Sub
                    Else
                        TimeCheck = MsgBox(Elem + " : No such an element in database!", vbExclamation, "Input error")
                        GoTo CheckElemAgain
                    End If
                End If
                Elem = mid$(Elem, 1, p - 1)
                Exit For
            Else
                rtoe = 1
            End If
        Next
        j = 1 + k
        For q = 1 To (iRow)
            If C1(q, 1) = Elem Then
                C2(j, 1) = C1(q, 1)       ' Element
                C2(j, 2) = C1(q, 2)       ' Transition
                C2(j, 4) = C1(q, 3)       ' KE
                C2(j, 7) = C1(q, 4 + aesoffset)       ' AES RSF
                C2(j, 11) = rtoe          ' atomic element ratio
                j = j + 1
            ElseIf LCase(Elem) = "all" And q > 1 Then
                C2(j, 1) = C1(q, 1)       ' Element
                C2(j, 2) = C1(q, 2)       ' Transition
                C2(j, 4) = C1(q, 3)       ' KE
                C2(j, 7) = C1(q, 4 + aesoffset)       ' AES RSF
                C2(j, 11) = rtoe          ' atomic element ratio
                j = j + 1
            End If
        Next
        k = j - 1
    Next
    
    numAESFactors = k
    maxAESFactor = 0
    
    If numAESFactors = 0 Then Exit Sub
    
    For n = 1 To k
        If C2(n, 7) = "NaN" Then C2(n, 7) = 0
        If C2(n, 7) >= maxAESFactor Then maxAESFactor = C2(n, 7) Else maxAESFactor = maxAESFactor
    Next
    
    If Abs(startEb - endEb) > fitLimit Then
        maxAESFactor = maxAESFactor * 4
    End If
    
    For n = 1 To numAESFactors
        C2(n, 8) = dblMin + (C2(n, 11) * C2(n, 7) * ((dblMax - dblMin) / (maxAESFactor)))
        C2(n, 2) = C2(n, 1) + C2(n, 2)
        C2(n, 9) = (C2(n, 11) * C2(n, 7) * ((chkMin) / (maxAESFactor)))
    Next
    
    Range(Cells((numXPSFactors + 51), para + 10), Cells((numXPSFactors + numAESFactors + 50), para + 20)) = C2
End Sub

Sub PlotElem()
    Dim oriXPSFactors As Integer, rngElemBeX As Range, rngElemBeA As Range, numFinal As Integer, pts As Points, pt As Point
    
    oriXPSFactors = numXPSFactors / (UBound(highpe) + 1)
    sheetGraph.Activate
    
    If strAna = "FitComp" Then
        maxXPSFactor = Cells(43, para + 10).Value
        maxAESFactor = Cells(44, para + 10).Value
        numChemFactors = Cells(42, para + 12).Value
        numXPSFactors = Cells(43, para + 12).Value
        numAESFactors = Cells(44, para + 12).Value
        
        If numXPSFactors = 0 And numAESFactors = 0 Then Exit Sub
        
        With ActiveSheet.ChartObjects(1).Chart
            For n = .SeriesCollection.Count To 1 Step -1
                If .SeriesCollection(n).Name = "XPS peaks in BE" Or .SeriesCollection(n).Name = "AES peaks in BE" Then
                    .SeriesCollection(n).Delete
                End If
            Next n
        End With
        
        If ActiveSheet.ChartObjects.Count > 1 Then
            With ActiveSheet.ChartObjects(2).Chart
                For n = .SeriesCollection.Count To 1 Step -1
                    If .SeriesCollection(n).Name = "XPS peaks in KE" Or .SeriesCollection(n).Name = "AES peaks in KE" Then
                        .SeriesCollection(n).Delete
                    End If
                Next n
            End With
        End If
    Else
        Call descriptHidden2
    End If

    numFinal = numXPSFactors + numAESFactors + 50
    Set rngElemBeX = Range(Cells(51, para + 14), Cells((50 + numXPSFactors), para + 14))
    Set rngElemBeA = Range(Cells((numXPSFactors + 51), para + 14), Cells(numFinal, para + 14))
    
    If numXPSFactors + numAESFactors = 0 Then
        Exit Sub
    ElseIf numXPSFactors = 0 And numAESFactors > 0 Then
        Cells((51 + numXPSFactors), para + 15).FormulaR1C1 = "=RC[-2] - R3C2 - R4C2"        ' KE char from KE
        Cells((51 + numXPSFactors), para + 14).FormulaR1C1 = "=R2C2 - RC[-1]"      ' BE char from KE
        Cells((51 + numXPSFactors), para + 17).FormulaR1C1 = "=R9C3 * ((R41C" & (para + 10) & " + (RC[3] * RC[-1] * (R42C" & (para + 10) & " - R41C" & (para + 10) & ")/R44C" & (para + 10) & ")) - R9C2)"
        Cells((51 + numXPSFactors), para + 18).FormulaR1C1 = "= (RC[-2] * " & (chkMin) & "/R44C" & (para + 10) & ") * R9C3"     ' Sens automatic update
    ElseIf numXPSFactors > 0 And numAESFactors = 0 Then
        Cells(51, para + 15).FormulaR1C1 = "=R2C2 - R3C2 - R4C2 - RC[-3]"     ' KE char from BE
        Cells(51, para + 14).FormulaR1C1 = "=RC[-2]"      ' BE char from BE
        Cells(51, para + 17).FormulaR1C1 = "=R9C3 * ((R41C" & (para + 10) & " + (RC[3] * RC[-1] * (R42C" & (para + 10) & " - R41C" & (para + 10) & ")/R43C" & (para + 10) & ")) - R9C2)"
    Else
        Cells(51, para + 15).FormulaR1C1 = "=R2C2 - R3C2 - R4C2 - RC[-3]"     ' KE char from BE
        Cells((51 + numXPSFactors), para + 15).FormulaR1C1 = "=RC[-2] - R3C2 - R4C2"        ' KE char from KE
        Cells(51, para + 14).FormulaR1C1 = "=RC[-2]"      ' BE char from BE
        Cells((51 + numXPSFactors), para + 14).FormulaR1C1 = "=R2C2 - RC[-1]"      ' BE char from KE
        Cells(51, para + 17).FormulaR1C1 = "=R9C3 * ((R41C" & (para + 10) & " + (RC[3] * RC[-1] * (R42C" & (para + 10) & " - R41C" & (para + 10) & ")/R43C" & (para + 10) & ")) - R9C2)"
        Cells((51 + numXPSFactors), para + 17).FormulaR1C1 = "=R9C3 * ((R41C" & (para + 10) & " + (RC[3] * RC[-1] * (R42C" & (para + 10) & " - R41C" & (para + 10) & ")/R44C" & (para + 10) & ")) - R9C2)"
        Cells((51 + numXPSFactors), para + 18).FormulaR1C1 = "= (RC[-2] * " & (chkMin) & "/R44C" & (para + 10) & ") * R9C3"     ' Sens automatic update
    End If
    
    If (numAESFactors > 1) Then
        Range(Cells((51 + numXPSFactors), para + 15), Cells(numFinal, para + 15)).FillDown
        Range(Cells((51 + numXPSFactors), para + 14), Cells(numFinal, para + 14)).FillDown
        Range(Cells((51 + numXPSFactors), para + 17), Cells(numFinal, para + 17)).FillDown
        Range(Cells((51 + numXPSFactors), para + 18), Cells(numFinal, para + 18)).FillDown
    End If
    
    If (numXPSFactors > 1) Then
        Range(Cells(51, para + 15), Cells((50 + numXPSFactors), para + 15)).FillDown
        Range(Cells(51, para + 14), Cells((50 + numXPSFactors), para + 14)).FillDown
        Range(Cells(51, para + 17), Cells((50 + numXPSFactors), para + 17)).FillDown
    End If
    
    If StrComp(Cells(2, 1).Value, "PE", 1) = 0 Then
        If UBound(highpe) > 0 Then
            For n = 1 To UBound(highpe)
                For q = 0 To oriXPSFactors - 1
                    Cells(51 + q + oriXPSFactors * n, para + 11) = Cells(51 + q + oriXPSFactors * n, para + 11).Value & "_" & Cells(40 + n, para + 13).Value
                Next
                
                Cells(51 + oriXPSFactors * n, para + 14).FormulaR1C1 = "=R2C2 - R" & (40 + n) & "C" & (para + 14) & " + RC[-2]"     ' BE higher order from BE
                Cells(51 + oriXPSFactors * n, para + 15).FormulaR1C1 = "=R" & (40 + n) & "C" & (para + 14) & " - R3C2 - R4C2 - RC[-3]"     ' KE char higher order from BE
                Cells(51 + oriXPSFactors * n, para + 17).FormulaR1C1 = "=R9C3 * (R41C" & (para + 10) & " + (RC[3] * RC[-1] * (R42C" & (para + 10) & " - R41C" & (para + 10) & ")/(R43C" & (para + 10) & " * " & (n + 1) & ")))"
                
                If (oriXPSFactors > 1) Then
                    Range(Cells(51 + oriXPSFactors * n, para + 14), Cells((50 + oriXPSFactors * (n + 1)), para + 14)).FillDown
                    Range(Cells(51 + oriXPSFactors * n, para + 15), Cells((50 + oriXPSFactors * (n + 1)), para + 15)).FillDown
                    Range(Cells(51 + oriXPSFactors * n, para + 17), Cells((50 + oriXPSFactors * (n + 1)), para + 17)).FillDown
                End If
            Next
        End If
    End If
    
    ActiveSheet.ChartObjects(1).Activate
    
    If StrComp(strl(3), "De", 1) = 0 Then
        j = 1
        GoTo AESmode1
    Else
        j = 0
    End If
    
    If numXPSFactors > 0 Then
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(2)   '2
            .ChartType = xlXYScatter
            .XValues = rngElemBeX
            .Values = rngElemBeX.Offset(0, 3)
            .MarkerStyle = 2
            .MarkerSize = 10 / Sqr(windowSize)
            .HasDataLabels = True
            .Name = "XPS peaks in BE"
        n = 0
        Set pts = .Points
        For Each pt In pts
            n = n + 1
            With pt.DataLabel
                .Text = rngElemBeX.Offset(0, -3).Cells(n).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 12 / Sqr(windowSize)
            End With
        Next
        
        End With
    End If
    
    If numAESFactors > 0 Then
AESmode1:
        
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(j * (-1) + 3)
            .ChartType = xlXYScatter
            .XValues = rngElemBeA.Offset(0, j)
            .Values = rngElemBeA.Offset(0, 3)
            .MarkerStyle = 9
            .MarkerSize = 10 / Sqr(windowSize)
            .HasDataLabels = True
            .Name = "AES peaks in BE"
        n = 0
        Set pts = .Points
        For Each pt In pts
            n = n + 1
            With pt.DataLabel
                .Text = rngElemBeA.Offset(0, -3).Cells(n).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 12 / Sqr(windowSize)
            End With
        Next
        
        End With
    End If
    
    If ActiveChart.HasLegend = True Then
        With ActiveSheet.ChartObjects(1).Chart
            For n = .SeriesCollection.Count To 1 Step -1
                If .SeriesCollection(n).Name = "XPS peaks in BE" Or .SeriesCollection(n).Name = "AES peaks in BE" Then
                    .Legend.LegendEntries(n).Delete
                End If
            Next n
        End With
    End If
    
    If StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(1), "Be", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then Exit Sub
    
    If ActiveSheet.ChartObjects.Count = 1 Then Exit Sub
    
    ActiveSheet.ChartObjects(2).Activate
    
    If StrComp(strl(3), "De", 1) = 0 Then GoTo AESmode2
    
    If numXPSFactors > 0 Then
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(2)
            .ChartType = xlXYScatter
            .XValues = rngElemBeX.Offset(0, 1)
            .Values = rngElemBeX.Offset(0, 3)
            .MarkerStyle = 2
            .MarkerSize = 10 / Sqr(windowSize)
            .HasDataLabels = True
            .Name = "XPS peaks in KE"
        n = 0
        Set pts = .Points
        For Each pt In pts
            n = n + 1
            With pt.DataLabel
                .Text = rngElemBeX.Offset(0, -3).Cells(n).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 12 / Sqr(windowSize)
            End With
        Next
    
        End With
    End If
    
    If numAESFactors > 0 Then
AESmode2:

        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(j * (-1) + 3)
            .ChartType = xlXYScatter
            .XValues = rngElemBeA.Offset(0, 1)
            .Values = rngElemBeA.Offset(0, 3 + j)
            .MarkerStyle = 9
            .MarkerSize = 10 / Sqr(windowSize)
            .HasDataLabels = True
            .Name = "AES peaks in KE"
        n = 0
        Set pts = .Points
        For Each pt In pts
            n = n + 1
            With pt.DataLabel
                .Text = rngElemBeA.Offset(0, -3).Cells(n).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 12 / Sqr(windowSize)
            End With
        Next
    
        End With
    End If
    
    If ActiveChart.HasLegend = True Then
        With ActiveSheet.ChartObjects(2).Chart
            For n = .SeriesCollection.Count To 1 Step -1
                If .SeriesCollection(n).Name = "XPS peaks in KE" Or .SeriesCollection(n).Name = "AES peaks in KE" Then
                    .Legend.LegendEntries(n).Delete
                End If
            Next n
        End With
    End If
End Sub

Sub PlotChem()
    Dim Fname As Variant, Record As Variant, C1 As Variant, C2 As Variant, C3 As Variant, rngElemBeC As Range, pts As Points, pt As Point, strTest As String
    
    If Dir(direc + "Chem" + backSlash) = vbNullString Then
        Set sheetGraph = Worksheets("Graph_" + strSheetDataName)
        sheetGraph.Activate
        If LCase(Cells(10, 1).Value) = "pe" Then
            Cells(10, 3).Value = "Ab"   'strl(3)
        Else
            Cells(10, 3).Value = "In"   'strl(3)
        End If
        
        Call GetOut
        End
    End If
    
    If strAna = "FitComp" Then
        numChemFactors = Cells(42, para + 12).Value
        
        If numChemFactors > 0 Then
            GoTo SkipChemLoad
        Else
            sheetGraph.Activate
            Exit Sub
        End If
    End If
    
    Set sheetGraph = Worksheets("Graph_" + strSheetDataName)
    sheetGraph.Activate
    numChemFactors = 0
    
    If StrComp(strAna, "chem", 1) = 0 Then
        If LCase(Cells(10, 1).Value) = "pe" Then
            Cells(10, 3).Value = "Ab"   'strl(3)
        Else
            Cells(10, 3).Value = "In"   'strl(3)
        End If
    Else
        Cells(42, para + 12).Value = 0
        Exit Sub
    End If
    
    C3 = Split(ElemD, ",")
    ReDim C2(1 To 101, 1 To 6)
    iRow = 1
    
    For n = 0 To UBound(C3)
        strTest = C3(n)
        q = 0
        Fname = direc + "Chem" + backSlash + strTest + "_ch"
        
        If Dir(Fname) = vbNullString Then
            TimeCheck = MsgBox("File Not Found in " + Fname + "!", vbExclamation, "Database error")
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
        
        If Fname = False Then Exit Sub
        
        fileNum = FreeFile(0)
        Open Fname For Input As fileNum

        Do
            Line Input #fileNum, Record
            C1 = Split(Record, vbTab)
            If q > 0 Then
                For iCol = 1 To 4
                    C2(iRow, iCol) = C1(iCol - 1)
                Next iCol
                iRow = iRow + 1
            End If
            q = 1
        Loop Until EOF(fileNum)
        
        Close #fileNum
        iRow = iRow - 1
    Next
    
    numChemFactors = iRow
    
    Cells(42, para + 12).Value = numChemFactors
    numXPSFactors = Cells(43, para + 12).Value
    C1 = Range(Cells(51, para + 10), Cells(50 + numXPSFactors, para + 11))
    
    If numChemFactors = 0 Then
        TimeCheck = MsgBox("No data in " + Fname + "!", vbExclamation, "Database error")
        Call GetOut
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    
    iRow = 0
    iCol = 0

    For q = 1 To numChemFactors
        For k = 1 To numXPSFactors
            If C2(q, 1) = C1(k, 2) And k = iRow Then
                'B(q, 6) = A(k, 8) - (iCol * A(k, 8) / 15)
                C2(q, 6) = Cells(50 + k, para + 17).Value - (iCol * Cells(9, 3).Value * (Cells(42, para + 10).Value - Cells(41, para + 10).Value) / 15)
                iCol = iCol + 1
            ElseIf C2(q, 1) = C1(k, 2) Then
                'B(q, 6) = A(k, 8)
                C2(q, 6) = Cells(50 + k, para + 17).Value
                iRow = k
                iCol = 1
            End If
        Next k
    Next q
    
    Range(Cells(51, para + 24), Cells((numChemFactors + 50), para + 29)) = C2
    Cells(51, para + 28).FormulaR1C1 = "=R2C2 - R3C2 - R4C2 - RC[-2]"     ' KE char from BE
    
    If numChemFactors > 1 Then
        Range(Cells(51, para + 28), Cells((50 + numChemFactors), para + 28)).FillDown
    End If

    Cells(50, para + 24).Value = "Chem"
    Cells(50, para + 25).Value = "Mater"
    Cells(50, para + 26).Value = "Shifts"
    Cells(50, para + 27).Value = "Errors"
    Cells(50, para + 28).Value = "KEshts"
    Cells(50, para + 29).Value = "R.Int"
    
SkipChemLoad:

    Set rngElemBeC = Range(Cells(51, para + 22), Cells((50 + numChemFactors), para + 22))
    
    ActiveSheet.ChartObjects(1).Activate
    For n = 1 To ActiveChart.SeriesCollection.Count
        If StrComp(ActiveChart.SeriesCollection(n).Name, "Chem shft in BE", 1) = 0 Then
            ActiveChart.SeriesCollection(n).Delete
            Exit For
        End If
    Next
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
        .ChartType = xlXYScatter
        .XValues = rngElemBeC
        .Values = rngElemBeC.Offset(0, 3)
        .MarkerStyle = xlMarkerStylePlus
        .MarkerSize = 3 / Sqr(windowSize)
        .HasDataLabels = True
        .ErrorBar Direction:=xlX, Include:=xlBoth, Type:=xlCustom, Amount:=rngElemBeC.Offset(0, 1), MinusValues:=rngElemBeC.Offset(0, 1)
        .Name = "Chem shft in BE"
        n = 0
        Set pts = .Points
        For Each pt In pts
            n = n + 1
            With pt.DataLabel
                .Text = rngElemBeC.Offset(0, -1).Cells(n).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 8 / Sqr(windowSize)
            End With
        Next
    End With
    
    If ActiveChart.HasLegend = True Then
        With ActiveSheet.ChartObjects(1).Chart
            n = .Legend.LegendEntries.Count
            .Legend.LegendEntries(n).Delete
        End With
    End If
    
    If StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(1), "Be", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Or ActiveSheet.ChartObjects.Count = 1 Then

    ElseIf ActiveSheet.ChartObjects.Count = 2 Then
        ActiveSheet.ChartObjects(2).Activate
        For n = 1 To ActiveChart.SeriesCollection.Count
            If StrComp(ActiveChart.SeriesCollection(n).Name, "Chem shft in BE", 1) = 0 Then
                ActiveChart.SeriesCollection(n).Delete
                Exit For
            End If
        Next
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
            .ChartType = xlXYScatter
            .XValues = rngElemBeC.Offset(0, 2)
            .Values = rngElemBeC.Offset(0, 3)
            .MarkerStyle = xlMarkerStylePlus
            .MarkerSize = 3 / Sqr(windowSize)
            .HasDataLabels = True
            .ErrorBar Direction:=xlX, Include:=xlBoth, Type:=xlCustom, Amount:=rngElemBeC.Offset(0, 1), MinusValues:=rngElemBeC.Offset(0, 1)
            .Name = "Chem shft in KE"
        n = 0
        Set pts = .Points
        For Each pt In pts
            n = n + 1
            With pt.DataLabel
                .Text = rngElemBeC.Offset(0, -1).Cells(n).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 8 / Sqr(windowSize)
            End With
        Next
        End With
        
        If ActiveChart.HasLegend = True Then
            With ActiveSheet.ChartObjects(2).Chart
                n = .Legend.LegendEntries.Count
                .Legend.LegendEntries(n).Delete
            End With
        End If
    End If
    
    strErr = "skip"
    If strAna = "FitComp" Then Exit Sub
    
    Call GetOut
End Sub

Sub GetCompare()
    Dim OpenFileName As Variant, fcmp As Variant, sBG As Variant, ncmp As Integer, rng As Range
    
    If StrComp(TimeCheck, "yes", 1) = 0 Then TimeCheck = vbNullString
    Worksheets(strSheetGraphName).Activate
    
    If StrComp(Cells(2, 1).Value, "PE shifts", 1) = 0 Then
        Results = ",Pe,Sh,Ab,,1," 'for XAS mode
    ElseIf StrComp(Cells(2, 1).Value, "PE", 1) = 0 Then
        If StrComp(Cells(10, 1).Value, "Be", 1) = 0 Then
            Results = ",Be,Sh," 'for XPS mode
        Else
            Results = ",Ke,Be," 'for XPS mode
        End If
        Results = Results & "In,,2,"
    ElseIf StrComp(Cells(2, 1).Value, "KE shifts", 1) = 0 Then
        If StrComp(Cells(1, 1).Value, "AES elec.", 1) = 0 Then
            Results = ",Ke,Ae,De,,3," ' for AES mode
        End If
    ElseIf StrComp(Cells(2, 1).Value, "Shifts", 1) = 0 Then
        Results = ",Po,Sh,Ab,,4," 'for DC mode
    ElseIf StrComp(Cells(2, 1).Value, "x offset", 1) = 0 Then
        Results = ",Po,Pn,Pp,,5," 'for RGA
    Else
        Call GetOut
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    
    If Cells(51, para + 9).Value = vbNullString Then
        Results = Results & "2"   ' XPS and AES modes without any factors plots only data.
    ElseIf Cells(42, para + 12).Value > 0 Then
        Results = Results & "5"   ' XPS mode with chemical shifts plots Data, XPS, AES, and Chem factors.
    ElseIf StrComp(Cells(2, 1).Value, "KE shifts", 1) = 0 Then
        Results = Results & "3"   ' AES mode plots Data and AES factors.
    Else
        Results = Results & "4" ' XPS mode without chemical shifts plots Data, XPS, and AES factors.
    End If
            
    If backSlash = "/" Then
        OpenFileName = Select_File_Or_Files_Mac("xlsx")
    Else
        ChDrive mid$(ActiveWorkbook.Path, 1, 1)
        ChDir ActiveWorkbook.Path
        OpenFileName = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Please select a file", MultiSelect:=True)
    End If
    
    If IsArray(OpenFileName) Then
        If UBound(OpenFileName) + cmp > CInt(para / 3) Then
            TimeCheck = MsgBox("Stop a comparison because you select too many files: " & (UBound(OpenFileName) + ncomp - (ncomp - cmp)) & " over the total limit: " & CInt(para / 3), vbExclamation)
            Cells(1, 4 + (cmp * 3)).Value = vbNullString
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf UBound(OpenFileName) > 1 And backSlash = "\" Then
            ' http://www.cpearson.com/excel/SortingArrays.aspx
            ' put the array values on the worksheet
            Cells(50, para + 25).Value = "List comps"
            Set rng = ActiveSheet.Cells(51, para + 25).Resize(UBound(OpenFileName) - LBound(OpenFileName) + 1, 1)
            rng = Application.Transpose(OpenFileName)
            ' sort the range
            rng.Sort key1:=rng, order1:=xlAscending, MatchCase:=False
            
            ' load the worksheet values back into the array
            For q = 1 To rng.Rows.Count
                OpenFileName(q) = rng(q, 1)
            Next q
            
            Range(Cells(50, para + 25), Cells(50 + UBound(OpenFileName), para + 25)).ClearContents
        End If
        
        Application.Calculation = xlCalculationManual
        Call EachComp(OpenFileName, strAna, fcmp, sBG, cmp, ncmp, ncomp)
        Application.Calculation = xlCalculationAutomatic
        
        Workbooks(wb).Sheets(strSheetGraphName).Activate
        If Not (ncmp - cmp) = 0 Then Call offsetmultiple
        If ncomp > cmp Then
            Cells(45, para + 10).Value = ncomp  ' total number of data compared but less than cmp
        Else
            Cells(45, para + 10).Value = ncmp   ' total number of data compared over cmp
        End If
    Else
        TimeCheck = "stop"
        Cells(1, 4 + (cmp * 3)).Value = vbNullString
        Call GetOut
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    
    Cells(1, 4 + (cmp * 3)).Value = vbNullString
    
    If ExistSheet("samples") Then
        Results = "n" & ncmp & "c" & cmp
        Call CombineLegend
    End If
    
    Call GetOut
End Sub

Sub GetOut()
    If Not Cells(8, 101).Value = 0 Then
        If StrComp(TimeCheck, "yes", 1) = 0 Then TimeCheck = "yes1"
    Else
        If ExistSheet(strSheetFitName) And strAna = "FitRatioAnalysis" Then
            Worksheets(strSheetFitName).Activate
        ElseIf ExistSheet(strSheetGraphName) Then
            Worksheets(strSheetGraphName).Activate
        End If
    End If

    Cells(1, 1).Select
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
    If StrComp("Fit", mid$(ActiveSheet.Name, 1, 3)) = 0 And IsNumeric(TimeCheck) = False Then 'Cells(9, 1).Value = "Solve LSM" Then
        If IsNumeric(Cells(9 + sftfit2, 2)) Then
            If fileNum >= Cells(17, 101).Value And Cells(9 + sftfit2, 2).Value < 10 Then   ' limit in # of iteration
                If a1 = a2 Then
                    TimeCheck = MsgBox("Tolerance result; " & vbCrLf & "Amp. ratio: " & a0 & " > " & a2 & ".", vbExclamation, "Iteration over " & fileNum & " !")
                ElseIf a0 = 0 Then
                    TimeCheck = MsgBox("Tolerance result; " & vbCrLf & "BE diff.: " & a1 & " > " & a2 & ".", vbExclamation, "Iteration over " & fileNum & " !")
                Else
                    TimeCheck = MsgBox("Tolerance results; " & vbCrLf & "Amp. ratio: " & a0 & " < " & a2 & vbCrLf & "BE diff.: " & a1 & " > " & a2 & ".", vbExclamation, "Iteration over " & fileNum & " !")
                End If
            ElseIf fileNum > 1 And Cells(9 + sftfit2, 2).Value < 10 Then
                If a1 = 0 Then
                    TimeCheck = MsgBox("Tolerance result; " & vbCrLf & "Amp. ratio: " & a0 & " < " & a2 & ".", vbInformation, "Iteration: " & fileNum)
                ElseIf a0 = 0 Then
                    TimeCheck = MsgBox("Tolerance result; " & vbCrLf & "BE diff.: " & a1 & " < " & a2 & ".", vbInformation, "Iteration: " & fileNum)
                Else
                    TimeCheck = MsgBox("Tolerance results; " & vbCrLf & "Amp. ratio: " & a0 & " < " & a2 & vbCrLf & "BE diff.: " & a1 & " < " & a2 & ".", vbInformation, "Iteration: " & fileNum)
                End If
            Else
                If IsEmpty(Cells(18, 101).Value) Then Cells(18, 101).FormulaR1C1 = "=Average(R21C2:R" & (20 + numData) & "C2)"
                If IsNumeric(Cells(18, 101).Value) Then
                    If Abs(Cells(18, 101).Value) < 0.000001 Then
                        TimeCheck = MsgBox("Fitting does not work properly, because avaraged In data is less than 1E-6!")
                    ElseIf Abs(Cells(18, 101).Value) > 1E+29 Then
                        TimeCheck = MsgBox("Fitting does not work properly, because avaraged In data is more than 1E+29!")
                    End If
                End If
            End If
        End If
    End If
    
    testMacro = vbNullString
    
    Application.DisplayAlerts = False
    
    If Len(ActiveWorkbook.Path) < 2 Then
        Application.Dialogs(xlDialogSaveAs).Show
    Else
        On Error GoTo Error1
        ActiveWorkbook.SaveAs fileName:=wbpath + backSlash + wb, FileFormat:=xlOpenXMLWorkbook
    End If
    
    Application.DisplayAlerts = True

    strErr = "skip"
    Exit Sub
Error1:
    MsgBox Error(Err)
    wb = mid$(ActiveWorkbook.Name, 1, InStr(ActiveWorkbook.Name, ".") - 1) + "_bk.xlsx"
    ActiveWorkbook.SaveAs fileName:=wbpath + backSlash + wb, FileFormat:=xlOpenXMLWorkbook
    Err.Clear
    strErr = "skip"
    Resume Next
End Sub

Sub GetAutoScale()
    Dim numDataT As Integer, npts As Integer, pstart As Integer, pend As Integer, jc As Integer, dt As Integer, dc As Integer, rng As Range, rg As Range
    Dim iniRow1 As Single, iniRow2 As Single, endRow1 As Single, endRow2 As Single, strAuto As String, maxv As Single, calv As Single, rngx As Range
    
    strAuto = LCase(Cells(1, 1).Value)
    
    If StrComp(strAuto, "autop", 1) = 0 And IsEmpty(Cells(40, para + 11).Value) = False Then strAuto = Cells(40, para + 11).Value

    npts = 0
    
    For dt = 0 To ncomp
        Set rngx = Range(Cells(11, (1 + (dt * 3))), Cells(11, (1 + (dt * 3))).End(xlDown))
        numDataT = Application.CountA(rngx)
        
        Set rng = Range(Cells(11, (3 + (dt * 3))), Cells(10 + numDataT, (3 + (dt * 3))))
        
        numDataT = Application.CountA(rng)
        If Len(strAuto) > 4 Then
            If StrComp(mid$(strAuto, 5, 1), "(", 1) = 0 And StrComp(mid$(strAuto, Len(strAuto), 1), ")", 1) = 0 Then
                If IsNumeric(mid$(strAuto, 6, InStr(6, strAuto, ",", 1) - 6)) And IsNumeric(mid$(strAuto, InStr(6, strAuto, ",", 1) + 1, Len(strAuto) - InStr(6, strAuto, ",", 1) - 1)) Then
                    pstart = Application.Floor(mid$(strAuto, 6, InStr(6, strAuto, ",", 1) - 6), 1)
                    pend = Application.Ceiling(mid$(strAuto, InStr(6, strAuto, ",", 1) + 1, Len(strAuto) - InStr(6, strAuto, ",", 1) - 1), 1)

                    If pstart >= 1 And pend > pstart Then
                    
                    Else
                        pstart = 1
                        pend = 10
                    End If
                Else
                    pstart = 1
                    pend = 10
                End If
                
                Set rng = Range(Cells(11 + numDataT - pend, (3 + (dt * 3))), Cells(11 + numDataT - pstart, (3 + (dt * 3))))
                Set dataData = Range(Cells(10 + pstart, (3 + (dt * 3))), Cells(10 + pend, (3 + (dt * 3))))
                
                If Application.WorksheetFunction.Average(dataData) > Application.WorksheetFunction.Average(rng) Then  ' PES mode
                    Cells(9, 3 * dt + 2).Value = Application.WorksheetFunction.Average(rng)
                    Cells(9, 3 * dt + 3).Value = 1 / Abs(Application.WorksheetFunction.Average(dataData) - Cells(9, 3 * dt + 2).Value)
                Else ' XAS mode
                    Cells(9, 3 * dt + 2).Value = Application.WorksheetFunction.Average(dataData)
                    Cells(9, 3 * dt + 3).Value = 1 / Abs(Application.WorksheetFunction.Average(rng) - Cells(9, 3 * dt + 2).Value)
                
                End If
            ElseIf StrComp(mid$(strAuto, 5, 1), "[", 1) = 0 And StrComp(mid$(strAuto, Len(strAuto), 1), "]", 1) = 0 And InStr(6, strAuto, ":", 1) > 0 And InStr(6, strAuto, ",", 1) > 0 Then
                stepEk = Abs(Cells(7, 3 * dt + 2).Value)
                ' BE range specified in "auto[273:274,291.5:294]" to calibrate offset in 293 and 274, and multiple in 291.5 and 294 eV as a unity
                
                If IsNumeric(mid$(strAuto, 6, InStr(6, strAuto, ":", 1) - 6)) Then
                    If mid$(strAuto, 6, InStr(6, strAuto, ":", 1) - 6) < 0 Then
                        iniRow1 = Application.Floor(mid$(strAuto, 6, InStr(6, strAuto, ":", 1) - 6), -1 * stepEk)
                    Else
                        iniRow1 = Application.Floor(mid$(strAuto, 6, InStr(6, strAuto, ":", 1) - 6), stepEk)
                    End If
                Else
                    iniRow1 = 0
                End If
                If IsNumeric(mid$(strAuto, InStr(6, strAuto, ",", 1) + 1, Len(strAuto) - InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) - 1)) Then
                    If mid$(strAuto, InStr(6, strAuto, ",", 1) + 1, Len(strAuto) - InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) - 1) < 0 Then
                        iniRow2 = Application.Floor(mid$(strAuto, InStr(6, strAuto, ",", 1) + 1, Len(strAuto) - InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) - 1), -1 * stepEk)
                    Else
                        iniRow2 = Application.Floor(mid$(strAuto, InStr(6, strAuto, ",", 1) + 1, Len(strAuto) - InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) - 1), stepEk)
                    End If
                Else
                    iniRow2 = 0
                End If
                If IsNumeric(mid$(strAuto, InStr(6, strAuto, ":", 1) + 1, InStr(InStr(6, strAuto, ":", 1) + 1, strAuto, ",", 1) - InStr(6, strAuto, ":", 1) - 1)) Then
                    If mid$(strAuto, InStr(6, strAuto, ":", 1) + 1, InStr(InStr(6, strAuto, ":", 1) + 1, strAuto, ",", 1) - InStr(6, strAuto, ":", 1) - 1) < 0 Then
                        endRow1 = Application.Ceiling(mid$(strAuto, InStr(6, strAuto, ":", 1) + 1, InStr(InStr(6, strAuto, ":", 1) + 1, strAuto, ",", 1) - InStr(6, strAuto, ":", 1) - 1), -1 * stepEk)
                    Else
                        endRow1 = Application.Ceiling(mid$(strAuto, InStr(6, strAuto, ":", 1) + 1, InStr(InStr(6, strAuto, ":", 1) + 1, strAuto, ",", 1) - InStr(6, strAuto, ":", 1) - 1), stepEk)
                    End If
                Else
                    endRow1 = 0
                End If
                If IsNumeric(mid$(strAuto, InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) + 1, Len(strAuto) - InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) - 1)) Then
                    If mid$(strAuto, InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) + 1, Len(strAuto) - InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) - 1) < 0 Then
                        endRow2 = Application.Ceiling(mid$(strAuto, InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) + 1, Len(strAuto) - InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) - 1), -1 * stepEk)
                    Else
                        endRow2 = Application.Ceiling(mid$(strAuto, InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) + 1, Len(strAuto) - InStr(InStr(6, strAuto, ",", 1) + 1, strAuto, ":", 1) - 1), stepEk)
                    End If
                Else
                    endRow2 = 0
                End If
                
                If StrComp(LCase(Cells(10, 3 * dt + 1).Value), "pe", 1) = 0 Then
                    If iniRow1 = endRow1 Then
                        Cells(9, 3 * dt + 2).Value = 0
                    Else
                        For jc = 0 To numDataT - 1
                            If iniRow1 <= Cells(12 + numDataT + 8 + jc, 3 * dt + 2).Value And IsEmpty(Cells(11 + jc, 3 * dt + 3).Value) = False Then
                                pstart = jc + 1
                                Exit For
                            End If
                        Next
                        
                        For jc = 0 To numDataT - 1
                            If endRow1 <= Cells(12 + numDataT + 8 + jc, 3 * dt + 2).Value And IsEmpty(Cells(11 + jc, 3 * dt + 3).Value) = False Then
                                pend = jc + 1
                                Exit For
                            End If
                        Next
                        
                        If pstart >= 1 And pend > pstart Then
                            Set rng = Range(Cells(11 + pstart - 1, (3 + (dt * 3))), Cells(11 + pend - 1, (3 + (dt * 3))))
                            Cells(9, 3 * dt + 2).Value = Application.WorksheetFunction.Average(rng)
                        End If
                    End If
                    
                    If iniRow2 = endRow2 Then
                        Cells(9, 3 * dt + 3).Value = 1
                    Else
                        For jc = 0 To numDataT - 1
                            If iniRow2 >= Cells(11 + (numDataT * 2) + 8 - jc, 3 * dt + 2).Value And IsEmpty(Cells(11 + jc, 3 * dt + 3).Value) = False Then
                                pend = jc + 1
                                Exit For
                            End If
                        Next
                        
                        For jc = 0 To numDataT - 1
                            If endRow2 >= Cells(11 + (numDataT * 2) + 8 - jc, 3 * dt + 2).Value And IsEmpty(Cells(11 + jc, 3 * dt + 3).Value) = False Then
                                pstart = jc + 1
                                Exit For
                            End If
                        Next
                    
                        If pstart >= 1 And pend > pstart Then
                            Set dataData = Range(Cells(10 + numDataT - pend + 1, (3 + (dt * 3))), Cells(10 + numDataT - pstart + 1, (3 + (dt * 3))))
                            Cells(9, 3 * dt + 3).Value = 1 / Abs(Application.WorksheetFunction.Average(dataData) - Cells(9, 3 * dt + 2).Value)
                        End If
                    End If
                Else
                    If iniRow1 = endRow1 Then
                        Cells(9, 3 * dt + 2).Value = 0
                    Else
                        For jc = 0 To numDataT - 1
                            If iniRow1 <= Cells(11 + (numDataT * 2) + 8 - jc, 3 * dt + 2).Value And IsEmpty(Cells(11 + jc, 3 * dt + 3).Value) = False Then
                                pstart = jc + 1
                                Exit For
                            
                            End If
                        Next
                        
                        For jc = 0 To numDataT - 1
                            If endRow1 <= Cells(11 + (numDataT * 2) + 8 - jc, 3 * dt + 2).Value And IsEmpty(Cells(11 + jc, 3 * dt + 3).Value) = False Then
                                pend = jc + 1
                                Exit For
                            End If
                        Next
                        
                        If pstart >= 1 And pend > pstart Then
                            Set rng = Range(Cells(10 + numDataT - pend + 1, (3 + (dt * 3))), Cells(10 + numDataT - pstart + 1, (3 + (dt * 3))))
                            Cells(9, 3 * dt + 2).Value = Application.WorksheetFunction.Average(rng)
                        End If
                    End If
                    
                    If iniRow2 = endRow2 Then
                        Cells(9, 3 * dt + 3).Value = 1
                    Else
                        For jc = 0 To numDataT - 1
                            If iniRow2 >= Cells(12 + numDataT + 8 + jc, 3 * dt + 2).Value And IsEmpty(Cells(11 + jc, 3 * dt + 3).Value) = False Then
                                pend = jc + 1
                                Exit For
                            End If
                        Next
                        
                        For jc = 0 To numDataT - 1
                            If endRow2 >= Cells(12 + numDataT + 8 + jc, 3 * dt + 2).Value And IsEmpty(Cells(11 + jc, 3 * dt + 3).Value) = False Then
                                pstart = jc + 1
                                Exit For
                            End If
                        Next
                        
                        If pstart >= 1 And pend > pstart Then
                            Set dataData = Range(Cells(11 + pstart - 1, (3 + (dt * 3))), Cells(11 + pend - 1, (3 + (dt * 3))))
                            Cells(9, 3 * dt + 3).Value = 1 / Abs(Application.WorksheetFunction.Average(dataData) - Cells(9, 3 * dt + 2).Value)
                        End If
                    End If
                End If
            ElseIf StrComp(mid$(strAuto, 5, 1), "{", 1) = 0 And StrComp(mid$(strAuto, Len(strAuto), 1), "}", 1) = 0 Then ' calibrate BE at max value
                npts = 0
                maxv = Application.Max(rng)
                
                For Each rg In rng
                    If rg = maxv Then
                        pstart = rg.Row
                    End If
                Next
                
                'pstart = Application.Match(maxv, rng, 0) + 11
                Debug.Print maxv, pstart, mid$(strAuto, 6, Len(strAuto) - 6)
                
                If IsEmpty(mid$(strAuto, 6, Len(strAuto) - 6)) = False Then
                    If IsNumeric(mid$(strAuto, 6, Len(strAuto) - 6)) Then
                        calv = mid$(strAuto, 6, Len(strAuto) - 6)
                    Else
                        calv = 284.6
                    End If
                Else
                    calv = 284.6
                End If
                
                Cells(4, 3 * dt + 2).Value = 0  ' reset char value to be calibrated
                Cells(4, 3 * dt + 2).Value = Cells(pstart, (2 + (dt * 3))).Value - calv
             ElseIf StrComp(mid$(strAuto, 5, 1), "'", 1) = 0 And StrComp(mid$(strAuto, Len(strAuto), 1), "'", 1) = 0 Then ' char to a value
                npts = 0
                maxv = Application.Max(rng)

                For Each rg In rng
                    If rg = maxv Then
                        pstart = rg.Row
                    End If
                Next

                'pstart = Application.Match(maxv, rng, 0) + 11
                Debug.Print maxv, pstart, mid$(strAuto, 6, Len(strAuto) - 6)

                If IsEmpty(mid$(strAuto, 6, Len(strAuto) - 6)) = False Then
                    If IsNumeric(mid$(strAuto, 6, Len(strAuto) - 6)) Then
                        calv = mid$(strAuto, 6, Len(strAuto) - 6)
                    Else
                        calv = 0
                    End If
                Else
                    calv = 0
                End If

                If Cells(2, 1).Value = "PE shifts" Then
                    dc = -2
                Else
                    dc = 0
                End If
                
                Cells(4 + dc, 3 * dt + 2).Value = calv ' reset char value as a constant
                
            ElseIf IsNumeric(mid$(strAuto, 5, Len(strAuto) - 4)) = True Then
                npts = mid$(strAuto, 5, Len(strAuto) - 4)
                If npts >= 0 And npts < numDataT / 2 Then
                    
                Else
                    npts = 0
                End If
                
                If npts = 0 Then       ' Auto0 makes all default
                    Cells(9, 3 * dt + 2).Value = 0
                    Cells(9, 3 * dt + 3).Value = 1
                ElseIf Cells(10 + npts, (3 + (dt * 3))).Value > Cells(11 + numDataT - npts, (3 + (dt * 3))).Value Then  ' PES mode
                    Cells(9, 3 * dt + 2).Value = Cells(11 + numDataT - npts, (3 + (dt * 3))).Value
                    Cells(9, 3 * dt + 3).Value = 1 / (Cells(10 + npts, (3 + (dt * 3))).Value - Cells(11 + numDataT - npts, (3 + (dt * 3))).Value)
                Else    ' XAS mode
                    Cells(9, 3 * dt + 2).Value = Cells(10 + npts, (3 + (dt * 3))).Value
                    Cells(9, 3 * dt + 3).Value = 1 / (Cells(11 + numDataT - npts, (3 + (dt * 3))).Value - Cells(10 + npts, (3 + (dt * 3))).Value)
                End If
            ElseIf StrComp(strAuto, "autowf", 1) = 0 Then
                ' point calibration in "autowf" for cutoff data
                npts = 0
                maxv = Application.Max(rng)
                
                For Each rg In rng
                    If rg = maxv Then
                        pstart = rg.Row
                    End If
                Next
                
                Debug.Print maxv, pstart
                
                If Cells(11 + npts, 3).Value < Cells(10 + numDataT - npts, 3).Value Then
                    Cells(9, 3 * dt + 2).Value = Cells(11 + npts, (3 + (dt * 3))).Value
                    Cells(9, 3 * dt + 3).Value = 1 / (Cells(pstart, (3 + (dt * 3))).Value - Cells(11 + npts, (3 + (dt * 3))).Value)
                Else
                    Cells(9, 3 * dt + 2).Value = Cells(10 + numDataT - npts, (3 + (dt * 3))).Value
                    Cells(9, 3 * dt + 3).Value = 1 / (Cells(pstart, (3 + (dt * 3))).Value - Cells(10 + numDataT - npts, (3 + (dt * 3))).Value)
                End If
            ElseIf StrComp(strAuto, "automax", 1) = 0 Then
                ' point calibration in "automax" for cutoff data
                npts = 0
                maxv = Application.Max(rng)
                
                For Each rg In rng
                    If rg = maxv Then
                        pstart = rg.Row
                    End If
                Next
                
                Debug.Print maxv, pstart
                
                If StrComp(LCase(Cells(10, 3 * dt + 1).Value), "pe", 1) = 0 Then 'XAS mode
                    Cells(9, 3 * dt + 2).Value = Cells(11 + npts, (3 + (dt * 3))).Value
                    Cells(9, 3 * dt + 3).Value = 1 / (Cells(pstart, (3 + (dt * 3))).Value - Cells(11 + npts, (3 + (dt * 3))).Value)
                Else    ' PES mode
                    Cells(9, 3 * dt + 2).Value = Cells(10 + numDataT - npts, (3 + (dt * 3))).Value
                    Cells(9, 3 * dt + 3).Value = 1 / (Cells(pstart, (3 + (dt * 3))).Value - Cells(10 + numDataT - npts, (3 + (dt * 3))).Value)
                End If
            Else
                npts = 0
                
                If StrComp(LCase(Cells(10, 3 * dt + 1).Value), "pe", 1) = 0 Then 'XAS mode
                    Cells(9, 3 * dt + 2).Value = Cells(11 + npts, (3 + (dt * 3))).Value
                    Cells(9, 3 * dt + 3).Value = 1 / (Cells(10 + numDataT - npts, (3 + (dt * 3))).Value - Cells(11 + npts, (3 + (dt * 3))).Value)
                Else    ' PES mode
                    Cells(9, 3 * dt + 2).Value = Cells(10 + numDataT - npts, (3 + (dt * 3))).Value
                    Cells(9, 3 * dt + 3).Value = 1 / (Cells(11 + npts, (3 + (dt * 3))).Value - Cells(10 + numDataT - npts, (3 + (dt * 3))).Value)
                End If
            End If
        Else ' point calibration in "auto" at start and end points
            npts = 0
            If Cells(11 + npts, (3 + (dt * 3))).Value > Cells(10 + numDataT - npts, (3 + (dt * 3))).Value Then  ' PES mode
                Cells(9, 3 * dt + 2).Value = Cells(10 + numDataT - npts, (3 + (dt * 3))).Value
                Cells(9, 3 * dt + 3).Value = 1 / (Cells(11 + npts, (3 + (dt * 3))).Value - Cells(10 + numDataT - npts, (3 + (dt * 3))).Value)
            Else    ' XAS mode
                Cells(9, 3 * dt + 2).Value = Cells(11 + npts, (3 + (dt * 3))).Value
                Cells(9, 3 * dt + 3).Value = 1 / (Cells(10 + numDataT - npts, (3 + (dt * 3))).Value - Cells(11 + npts, (3 + (dt * 3))).Value)
            End If
        End If
    Next
    
    Cells(40, para + 11).Value = strAuto
    If StrComp(mid$(ActiveSheet.Name, 1, 4), "Cmp_", 1) = 0 Then
        Cells(1, 1).Value = vbNullString
        End
    Else
        Cells(1, 1).Value = "Grating"
        
        If ncomp > 0 Then
            strErr = "skip"
        Else
            off = 0
            multi = 1
        End If
    End If
End Sub

Sub ExportCmp(ByRef strXas As String)
    Dim rng As Range, numDataT As Integer
    
    If LCase(Cells(1, 1).Value) = "exp" Or strXas = "Is" Then
        If ExistSheet(strSheetAnaName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetAnaName).Delete
            Application.DisplayAlerts = True
        End If
            
        Worksheets.Add().Name = strSheetAnaName
        Set sheetAna = Worksheets(strSheetAnaName)
        Set sheetGraph = Worksheets(strSheetGraphName)
    
        wb = ActiveWorkbook.Name
        sheetGraph.Activate
        
        If strXas = "Is" Then
            Cells(1, 1).Value = "Grating"
            ncomp = 0
        Else
            Cells(1, 1).Value = "Goto " & strSheetAnaName
        End If
        
        For q = 0 To ncomp
            Set rng = Range(Cells(11, (2 + (q * 3))), Cells(11, (2 + (q * 3))).End(xlDown))
            numDataT = Application.CountA(rng)
            sheetGraph.Range(Cells(11 + numDataT + 8, (2 + (q * 3))), Cells(11 + (numDataT * 2) + 8, (3 + (q * 3)))).Copy
            sheetAna.Cells(1, 1 + (q * 2)).PasteSpecial Paste:=xlValues
            sheetAna.Cells(1, 1 + (q * 2)).Value = "BE/eV"
        Next
        
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    
    Application.CutCopyMode = False
    Cells(1, 1).Select
    
    If strXas = "Is" Then
    Else
        strErr = "skip"
    End If
End Sub

Sub Convert2Txt(ByRef strXas As String)
    Dim numDataT As Integer, numDataF As Integer, ElemT As String, rng As Range, strCpa As String, strTest As String
    
    Set rng = [1:1]
    iCol = Application.CountA(rng)
    strCpa = ActiveWorkbook.Path
    strSheetAnaName = ActiveSheet.Name
    Set sheetAna = Worksheets(strSheetAnaName)
    ElemT = vbNullString
    numDataF = FreeFile
    ' http://www.homeandlearn.org/write_to_a_text_file.html
    For q = 0 To (iCol / 2) - 1
        If iCol <= 3 Then
            If strXas = "Ip" Then
                strLabel = strSheetAnaName
            ElseIf strXas = "Is" Then
                strLabel = strSheetDataName
            Else
                strLabel = strSheetDataName
            End If
            iCol = 2
        Else
            strLabel = sheetAna.Cells(1, 2 + (q * 2)).Value
        End If
        
        strTest = strCpa & backSlash & strLabel & ".txt"
        Set rng = sheetAna.Range(Cells(1, 2 + (q * 2)), Cells(1, (2 + (q * 2))).End(xlDown))
        numDataT = Application.CountA(rng)
        
        Open strTest For Output As #numDataF
        For j = 1 To numDataT
            For k = 1 To 2
                If k = 2 Then
                    ElemT = ElemT + Trim(sheetAna.Cells(j, k + (q * 2)).Value)
                Else
                    ElemT = Trim(sheetAna.Cells(j, k + (q * 2)).Value) + vbTab
                End If
            Next k
            Print #numDataF, ElemT
            ElemT = vbNullString
        Next j
        
        Close #numDataF
        numDataF = numDataF + 1
    Next q
    
    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub

    Application.CutCopyMode = False
    
    If strXas = "Is" Or strXas = "Ip" Then
    Else
        strErr = "skip"
    End If
End Sub

Sub FitRatioAnalysis()
    Dim C1 As Variant, C2 As Variant, C3 As Variant, peakNum As Integer, fitNum As Integer, bookNum As Integer
    Dim OpenFileName As Variant, fcmp As Variant, sBG As Variant, ncmp As Integer, ncomp As Integer, rng As Range
    
    strSheetAnaName = "Ana_" + strSheetDataName
    strSheetFitName = "Rto_" + strSheetDataName
    strSheetGraphName = "Graph_" + strSheetDataName
    
    If ExistSheet(strSheetFitName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetFitName).Delete
        Application.DisplayAlerts = True
    End If
        
    Worksheets.Add().Name = strSheetFitName
    Set sheetAna = Worksheets(strSheetAnaName)
    Set sheetFit = Worksheets(strSheetFitName)
    Set sheetGraph = Worksheets(strSheetGraphName)
    
    If backSlash = "/" Then
        OpenFileName = Select_File_Or_Files_Mac("xlsx")
    Else
        ChDrive mid$(ActiveWorkbook.Path, 1, 1)
        ChDir ActiveWorkbook.Path
        OpenFileName = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Please select a file", MultiSelect:=True)
    End If
    
    If IsArray(OpenFileName) Then
        If UBound(OpenFileName) > para / 3 Then
            TimeCheck = MsgBox("Stop a comparison because you select too many files: " & UBound(OpenFileName) & " over the total limit: " & para / 3, vbExclamation)
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf UBound(OpenFileName) > 1 And backSlash = "\" Then
            ' http://www.cpearson.com/excel/SortingArrays.aspx
            ' put the array values on the worksheet
            Set rng = sheetFit.Range("A1").Resize(UBound(OpenFileName) - LBound(OpenFileName) + 1, 1)
            rng = Application.Transpose(OpenFileName)
            ' sort the range
            rng.Sort key1:=rng, order1:=xlAscending, MatchCase:=False
            
            ' load the worksheet values back into the array
            For q = 1 To rng.Rows.Count
                OpenFileName(q) = rng(q, 1)
            Next q
        End If
        
        strAna = "FitRatioAnalysis"
        
        sheetAna.Activate
        spacer = sheetAna.Cells(2, para + 1).Value
        peakNum = sheetAna.Cells(3, para + 1).Value         ' # of Fit peaks
        fitNum = sheetAna.Cells(4, para + 1).Value   ' # of Fit files
        sheetAna.Cells(5, para + 1).Value = UBound(OpenFileName)  ' # of Ana files
        sheetAna.Cells(5, para).Value = "# ana files"
        bookNum = UBound(OpenFileName)
        sheetAna.Cells(1, 1) = vbNullString
        C3 = sheetAna.Range(Cells(1, 1), Cells(para * 3 - 1, para * 3 - 1)) ' No check in matching among the peak names.
        sheetFit.Activate
        sheetFit.Range(Cells(1, 1), Cells(para * 3 - 1, para * 3 - 1)) = C3
        C2 = sheetFit.Range(Cells(4, para / 2), Cells(3 + fitNum, para - 1))    ' store the BGs
        
        For q = 1 To fitNum
            C2(q, 1) = C3(3 + q, peakNum + 6) & C3(3 + q, peakNum + 7) & C3(3 + q, peakNum + 8)
        Next
        
        sheetFit.Range(Cells(1, 5 + peakNum), Cells((spacer + fitNum) * 5 + 3, 10 + peakNum * 2)).ClearContents
        sheetFit.Cells(1, 5).Value = ActiveWorkbook.Name
        sheetFit.Cells(2, 5).Value = sheetAna.Name
        sheetFit.Cells(1, 1).Value = "Multiple-element ratio analysis"
        C3 = sheetFit.Range(Cells(1, 1), Cells(para * 3 - 1, para * 3 - 1))
        
        Results = "0," & strl(1) & "," & strl(2) & "," & strl(3) & ",,,"
        ncomp = 0
        cmp = 0     ' position of comp, should be zero
        fcmp = C3
        sBG = C2
        
        Call EachComp(OpenFileName, strAna, fcmp, sBG, cmp, ncmp, ncomp)       ' Copy fitting parameters in each Fit sheet
        
        C3 = fcmp
        C2 = sBG
        
        sheetFit.Activate
        peakNum = sheetFit.Cells(3, para + 1).Value    ' # of peaks
        
        C3(1, 4) = "File"
        C3(2, 4) = "Sheet"
        C3(3, peakNum + 6) = "Background"      ' G is # of peaks in the main sheet. Peaks over this # do not appear.
        C3(2, peakNum + 8 + bookNum) = "Difference"   ' bookNum represents number of BGs

        C3(3 + (spacer + fitNum - 1), peakNum + 6) = "Total peak area"
        C3(2 + (spacer + fitNum - 1), peakNum + 9) = "T.I.Area ratio"
    
        C3(3 + (spacer + fitNum - 1) * 2, peakNum + 6) = "Summation"             ' you can choose
        C3(2 + (spacer + fitNum - 1) * 2, peakNum + 9) = "S.I.Area ratio"            ' normalized by summation
        C3(3 + (spacer + fitNum - 1) * 2, 2 * peakNum + 9) = "Total ratio"
        
        C3(3 + (spacer + fitNum - 1) * 3, peakNum + 6) = "Summation"               ' you can choose
        C3(2 + (spacer + fitNum - 1) * 3, peakNum + 9) = "N.I.Area ratio"            ' normalized by summation
        C3(3 + (spacer + fitNum - 1) * 3, 2 * peakNum + 9) = "Total ratio"
        C3(3 + (spacer + fitNum - 1) * 4, peakNum + 6) = "Average"
        
        For n = 0 To 4      ' n represents # of parameters to be summarized
            Range(Cells(3 - n + (spacer + fitNum) * n, 5), Cells(3 - n + (spacer + fitNum) * n, 4 + peakNum)).Interior.ColorIndex = 38
            Cells(3 + (spacer + fitNum - 1) * n, 1).Interior.ColorIndex = 3
            Range(Cells(3 + (spacer + fitNum - 1) * n, 2), Cells(3 + (spacer + fitNum - 1) * n, 3)).Interior.ColorIndex = 4
            Cells(3 + (spacer + fitNum - 1) * n, 4).Interior.ColorIndex = 5
            
            If n = 0 Then
                Range(Cells(3 + (spacer + fitNum - 1) * n, peakNum + 6), Cells(3 + (spacer + fitNum - 1) * n, peakNum + 6 + bookNum)).Interior.ColorIndex = 6
            Else
                Range(Cells(3 + (spacer + fitNum - 1) * n, peakNum + 6), Cells(3 + (spacer + fitNum - 1) * n, peakNum + 7)).Interior.ColorIndex = 6
            End If
            Cells(3 + (spacer + fitNum - 1) * n, 4).Font.ColorIndex = 2
            
            For k = 0 To fitNum - 1
                C3(4 + k + (spacer + fitNum - 1) * n, 4) = peakNum
            Next
            
            For k = 0 To peakNum - 1
                C3(3 + (spacer + fitNum - 1) * 2, peakNum + 9 + k) = C3(3 + (spacer + fitNum - 1) * 2, 5 + k)
                C3(3 + (spacer + fitNum - 1) * 3, peakNum + 9 + k) = C3(3 + (spacer + fitNum - 1) * 3, 5 + k)
            Next
        Next
        
        Cells(1, 4).Interior.ColorIndex = 9
        Cells(2, 4).Interior.ColorIndex = 10
        
        For n = 0 To 1
            Cells(1 + n, 4).Font.ColorIndex = 2
        Next
        Range(Cells(2 + (spacer + fitNum - 1) * 0, peakNum + 8 + bookNum), Cells(2 + (spacer + fitNum - 1) * 0, peakNum + 9 + bookNum)).Interior.ColorIndex = 8  ' Difference
        
        For n = 1 To 4
            Range(Cells(2 + (spacer + fitNum - 1) * n, peakNum + 9), Cells(2 + (spacer + fitNum - 1) * n, peakNum + 10)).Interior.ColorIndex = 8   ' Area ratio
        Next
        
        Cells(3 + (spacer + fitNum - 1) * 2, 2 * peakNum + 9).Interior.ColorIndex = 26   ' Total ratio in S. Area ratio
        Cells(3 + (spacer + fitNum - 1) * 3, 2 * peakNum + 9).Interior.ColorIndex = 26   ' Total ratio in N. Area ratio
        Range(Cells(3 + (spacer + fitNum - 1) * 2, peakNum + 9), Cells(3 + (spacer + fitNum - 1) * 2, 2 * peakNum + 8)).Interior.ColorIndex = 38  ' Peak names in S. Area ratio
        Range(Cells(3 + (spacer + fitNum - 1) * 3, peakNum + 9), Cells(3 + (spacer + fitNum - 1) * 3, 2 * peakNum + 8)).Interior.ColorIndex = 38  ' Peak names in N. Area ratio
        sheetFit.Range(Cells(1, 1), Cells(para - 1, para - 1)) = C3
        sheetFit.Range(Cells(4, peakNum + 6), Cells(3 + fitNum, 2 * peakNum + 6)) = C2 ' back BG (A)
        
        For n = 0 To fitNum - 1
            Cells(4 + n + 1 * (spacer + fitNum - 1), peakNum + 6).FormulaR1C1 = "=Sum(RC5:RC" & (peakNum + 4) & ")"                     ' Total P.Area
            Cells(4 + n + 2 * (spacer + fitNum - 1), peakNum + 6).FormulaR1C1 = "=Sum(RC5:RC" & (peakNum + 4) & ")"                     ' Total S.Area
            Cells(4 + n + 3 * (spacer + fitNum - 1), peakNum + 6).FormulaR1C1 = "=Sum(RC5:RC" & (peakNum + 4) & ")"                     ' Total N.Area
            Cells(4 + n + 4 * (spacer + fitNum - 1), peakNum + 6).FormulaR1C1 = "=Average(RC5:RC" & (peakNum + 4) & ")"                 ' Avg FHHM
            
            For p = 0 To peakNum - 2
                Cells(4 + n, peakNum + 8 + bookNum + p).FormulaR1C1 = "=(RC" & (6 + p) & " - RC" & (5 + p) & ")"                             ' Difference
                Cells(4 + n + 1 * (spacer + fitNum - 1), peakNum + 9 + p).FormulaR1C1 = "=(RC" & (5 + p) & " / RC" & (6 + p) & ")"   ' P.Area ratio
            Next
            
            For p = 0 To peakNum - 1
                Cells(4 + n + 2 * (spacer + fitNum - 1), peakNum + 9 + p).FormulaR1C1 = "=(100 * RC" & (5 + p) & "/RC" & (peakNum + 6) & ")"  ' S.Area ratio
            Next
            Cells(4 + n + 2 * (spacer + fitNum - 1), 2 * peakNum + 9).FormulaR1C1 = "=Sum(RC[" & (-peakNum) & "]:RC[-1])"               ' Total S.Area ratio
            
            For p = 0 To peakNum - 1
                Cells(4 + n + 3 * (spacer + fitNum - 1), peakNum + 9 + p).FormulaR1C1 = "=(100 * RC" & (5 + p) & "/RC" & (peakNum + 6) & ")"  ' N.Area ratio
            Next
            Cells(4 + n + 3 * (spacer + fitNum - 1), 2 * peakNum + 9).FormulaR1C1 = "=Sum(RC[" & (-peakNum) & "]:RC[-1])"               ' Total N.Area ratio
        Next
        
        For n = 0 To 4
            If n > 0 Then
                For k = 0 To peakNum - 1
                    Cells(3 + (spacer + fitNum - 1) * n, k + 5).FormulaR1C1 = "=R3C" & (k + 5) & ""
                Next
            End If
            
            Set dataBGraph = Range(Cells(4 + (spacer + fitNum - 1) * n, 5), Cells(4 + (spacer + fitNum - 1) * n, 5).Offset(fitNum - 1, peakNum - 1))
            
            Charts.Add
            ActiveChart.ChartType = xlLineMarkers
            ActiveChart.SetSourceData Source:=dataBGraph, PlotBy:=xlColumns
            ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetFitName

            For k = 1 To peakNum
                ActiveChart.SeriesCollection(k).Name = "='" & ActiveSheet.Name & "'!R3C" & (4 + k) & ""  ' Cells(3, 4 + k).Value
                ActiveChart.SeriesCollection(k).AxisGroup = 1
            Next
            
            If Cells(4 + (spacer + fitNum - 1) * n, 4).Value > 1 And n = 0 Then    ' difference
                For k = 1 To peakNum - 1
                    Set dataKGraph = Range(Cells(4 + (spacer + fitNum - 1) * n, 2 * peakNum + 7 + k - 1), Cells(4 + (spacer + fitNum - 1) * n + fitNum - 1, 2 * peakNum + 7 + k - 1))
                    ActiveChart.SeriesCollection.NewSeries
                    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
                        .ChartType = xlColumnClustered
                        .Values = dataKGraph
                        Cells((3 + (spacer + fitNum - 1) * n), peakNum + 7 + k + bookNum).FormulaR1C1 = "=R3C" & (5 + k) & " & ""-"" & R3C" & (4 + k) & ""
                        Cells((3 + (spacer + fitNum - 1) * n), peakNum + 7 + k + bookNum).Interior.ColorIndex = 38
                        .Name = "='" & ActiveSheet.Name & "'!R3C" & (peakNum + 7 + k + bookNum) & ""                'Cells(3, 5 + k).Value + "-" + Cells(3, 4 + k).Value
                        .AxisGroup = 2
                    End With
                Next
            ElseIf Cells(4 + (spacer + fitNum - 1) * n, 4).Value > 1 And n = 1 Then
                For k = 1 To peakNum - 1
                    Set dataKGraph = Range(Cells(4 + (spacer + fitNum - 1) * n, peakNum + 9 + k - 1), Cells(4 + (spacer + fitNum - 1) * n + fitNum - 1, peakNum + 9 + k - 1))
                    ActiveChart.SeriesCollection.NewSeries
                    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
                        .ChartType = xlColumnClustered
                        .Values = dataKGraph
                        Cells((3 + (spacer + fitNum - 1) * n), peakNum + 8 + k).FormulaR1C1 = "=R3C" & (4 + k) & " & ""/"" & R3C" & (5 + k) & ""
                        Cells((3 + (spacer + fitNum - 1) * n), peakNum + 8 + k).Interior.ColorIndex = 38
                        .Name = "='" & ActiveSheet.Name & "'!R" & (3 + (spacer + fitNum - 1) * n) & "C" & (peakNum + 8 + k) & ""               'Cells(3, 4 + k).Value + "/" + Cells(3, 5 + k).Value
                        .AxisGroup = 2
                    End With
                Next
            ElseIf Cells(4 + (spacer + fitNum - 1) * n, 4).Value > 0 And n >= 2 And n <= 3 Then
                For k = 1 To peakNum
                    Set dataKGraph = Range(Cells(4 + (spacer + fitNum - 1) * n, peakNum + 9 + k - 1), Cells(4 + (spacer + fitNum - 1) * n + fitNum - 1, peakNum + 9 + k - 1))
                    ActiveChart.SeriesCollection.NewSeries
                    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
                        .ChartType = xlAreaStacked100
                        Cells((3 + (spacer + fitNum - 1) * n), peakNum + 8 + k).FormulaR1C1 = "= ""Rto_"" & R3C" & (4 + k) & ""
                        .Name = "='" & ActiveSheet.Name & "'!R" & (3 + (spacer + fitNum - 1) * n) & "C" & (peakNum + 8 + k) & ""     'Cells(3, 4 + k).Value

                        .Values = dataKGraph
                        .AxisGroup = 2
                    End With
                Next
            End If
            
            With ActiveChart.Axes(xlCategory, xlPrimary)
                .HasTitle = True
                .AxisTitle.Text = "Samples"
                .AxisTitle.Font.Size = 12
                .AxisTitle.Font.Bold = False
            End With

            With ActiveChart.Axes(xlValue, xlPrimary)
                .HasTitle = True
                If n = 0 Then
                    .AxisTitle.Text = "Binding energy (eV)"
                ElseIf n = 1 Then
                    .AxisTitle.Text = "T.I. Area"
                ElseIf n = 2 Then
                    .AxisTitle.Text = "S.I. Area"
                ElseIf n = 3 Then
                    .AxisTitle.Text = "N.I. Area"
                ElseIf n = 4 Then
                    .AxisTitle.Text = "FWHM (eV)"
                End If
                .AxisTitle.Font.Size = 12
                .AxisTitle.Font.Bold = False
            End With
            
            If n < 3 And peakNum > 1 Then
                With ActiveChart.Axes(xlValue, xlSecondary)
                    .HasTitle = True
                    If n = 0 Then
                        .AxisTitle.Text = "Difference (eV)"
                    ElseIf n = 1 Then
                        .AxisTitle.Text = "Ratio (peak-to-peak)"
                    ElseIf n = 2 Then
                        .AxisTitle.Text = "Ratio (%)"
                    End If
                    .AxisTitle.Font.Size = 12
                    .AxisTitle.Font.Bold = False
                End With
            End If
        
            With ActiveSheet.ChartObjects(1 + n)
                .Top = 20 + (500 / (windowSize * 2)) * n
                .Left = 200 * 5
                .Width = (550 * windowRatio) / (windowSize * 2)
                .Height = 500 / (windowSize * 2)
                
                With .Chart.Legend
                    .Position = xlLegendPositionRight
                    .IncludeInLayout = True
                    .Left = (850 / (windowSize * 2))
                    .Top = (50 / (windowSize * 2))
                    With .Format.Fill
                        .Visible = msoTrue
                        .ForeColor.RGB = RGB(255, 255, 255)
                        .ForeColor.TintAndShade = 0.1
                    End With
                End With
                With .Chart
                    .PlotArea.Width = (((550 * windowRatio) - 100) / (windowSize * 2))
                    .ChartArea.Border.LineStyle = 0
                End With
            End With
        Next
        
        sheetFit.Activate
    Else
        TimeCheck = "stop"
    End If
    
SkipFitRatioAnalysis:
    Call GetOut
End Sub

Sub FitAnalysis()
    Dim C1 As Variant, C2 As Variant, C3 As Variant, peakNum As Integer, fitNum As Integer, bookNum As Integer, imax As Integer
    Dim OpenFileName As Variant, fcmp As Variant, sBG As Variant, ncmp As Integer, ncomp As Integer, rng As Range, strSheetCmpName As String, strTest As String
    
    If Len(strSheetDataName) > 25 Then strSheetDataName = mid$(strSheetDataName, 1, 25)

    peakNum = Workbooks(wb).Sheets("Fit_" + strSheetDataName).Cells(8 + sftfit2, 2).Value
    C1 = Workbooks(wb).Sheets("Fit_" + strSheetDataName).Range(Cells(1, 5), Cells(19 + sftfit2, 4 + peakNum))
    C2 = Workbooks(wb).Sheets("Fit_" + strSheetDataName).Range(Cells(1, 1), Cells(1, 3))
            
    strSheetAnaName = "Ana_" + strSheetDataName
    strSheetFitName = "Fit_" + strSheetDataName
    strSheetGraphName = "Graph_" + strSheetDataName
    strSheetCmpName = "Cmp_" + strSheetDataName

    If ExistSheet(strSheetAnaName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetAnaName).Delete
        Application.DisplayAlerts = True
    End If
        
    Worksheets.Add().Name = strSheetAnaName
    Set sheetAna = Worksheets(strSheetAnaName)
    Set sheetFit = Worksheets(strSheetFitName)
    Set sheetGraph = Worksheets(strSheetGraphName)

    If backSlash = "/" Then
        OpenFileName = Select_File_Or_Files_Mac("xlsx")
    Else
        ChDrive mid$(ActiveWorkbook.Path, 1, 1)
        ChDir ActiveWorkbook.Path
        OpenFileName = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Please select a file", MultiSelect:=True)
    End If
    
    If IsArray(OpenFileName) Then
        If UBound(OpenFileName) > para / 3 Then
            TimeCheck = MsgBox("Stop a comparison because you select too many files: " & UBound(OpenFileName) & " over the total limit: " & para / 3, vbExclamation)
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf UBound(OpenFileName) > 1 And backSlash = "\" Then
            ' http://www.cpearson.com/excel/SortingArrays.aspx
            ' put the array values on the worksheet
            Set rng = sheetFit.Cells(1, 110).Resize(UBound(OpenFileName) - LBound(OpenFileName) + 1, 1)
            rng = Application.Transpose(OpenFileName)
            ' sort the range
            rng.Sort key1:=rng, order1:=xlAscending, MatchCase:=False
            
            ' load the worksheet values back into the array
            For q = 1 To rng.Rows.Count
                OpenFileName(q) = rng(q, 1)
            Next q
            
            Range(Cells(1, 110), Cells(UBound(OpenFileName), 110)).ClearContents
        End If
        
        strAna = "FitAnalysis"
        
        Cells(1, para).Value = "Parameters"
        Cells(2, para).Value = "Spacer"
        Cells(3, para).Value = "# peaks"
        Cells(4, para).Value = "# fit files"
        Cells(2, para + 1).Value = spacer
        Cells(3, para + 1).Value = peakNum
        fitNum = UBound(OpenFileName)
        Cells(4, para + 1).Value = fitNum + 1
        
        C3 = sheetAna.Range(Cells(1, 1), Cells((4 + spacer * 4) + 5 * fitNum, 9 + 2 * peakNum)) ' No check in matching among the peak names.

        C3(3, peakNum + 6) = "Background"      ' G is # of peaks in the main sheet. Peaks over this # do not appear.
        C3(2, peakNum + 9) = "Difference"
        C3(2, 1) = "BE"
        C3(2 + (spacer + fitNum), 1) = "T.I.Area"
        C3(3 + (spacer + fitNum), peakNum + 6) = "Total peak area"
        numData = sheetFit.Cells(5, 101).Value
        C3(2 + (spacer + fitNum), peakNum + 9) = "T.I.Area ratio"
        
        C3(2 + (spacer + fitNum) * 2, 1) = "S.I.Area"
        C3(3 + (spacer + fitNum) * 2, peakNum + 6) = "Summation"               ' you can choose
        C3(2 + (spacer + fitNum) * 2, peakNum + 9) = "S.I.Area ratio"            ' normalized by summation
        C3(3 + (spacer + fitNum) * 2, 2 * peakNum + 9) = "Total ratio"
        
        C3(2 + (spacer + fitNum) * 3, 1) = "N.I.Area"
        C3(3 + (spacer + fitNum) * 3, peakNum + 6) = "Summation"               ' you can choose
        C3(2 + (spacer + fitNum) * 3, peakNum + 9) = "N.I.Area ratio"            ' normalized by summation
        C3(3 + (spacer + fitNum) * 3, 2 * peakNum + 9) = "Total ratio"
        
        C3(2 + (spacer + fitNum) * 4, 1) = "FWHM"
        C3(3 + (spacer + fitNum) * 4, peakNum + 6) = "Average"
        
        For iCol = 0 To peakNum - 1
            C3(3, iCol + 5) = C1(1, iCol + 1)                                 ' Peak #1
            C3(4, iCol + 5) = C1(2, iCol + 1)                                 ' BE
            C3(3 + (spacer + fitNum), iCol + 5) = C1(1, iCol + 1)         ' Peak #2
            C3(3 + (spacer + fitNum) * 2, iCol + 5) = C1(1, iCol + 1)     ' Peak #3
            C3(3 + (spacer + fitNum) * 3, iCol + 5) = C1(1, iCol + 1)     ' Peak #4
            C3(3 + (spacer + fitNum) * 2, iCol + 9 + peakNum) = C1(1, iCol + 1) ' Peak #3 for ratio
            C3(3 + (spacer + fitNum) * 3, iCol + 9 + peakNum) = C1(1, iCol + 1) ' Peak #4 for ratio
            
            If C1(16 + sftfit2, iCol + 1) > 0 Then
                C3(4 + (spacer + fitNum), iCol + 5) = C1(16 + sftfit2, iCol + 1)      ' T.I.Area
                C3(4 + (spacer + fitNum) * 2, iCol + 5) = C1(17 + sftfit2, iCol + 1)  ' S.I.Area
                C3(4 + (spacer + fitNum) * 3, iCol + 5) = C1(18 + sftfit2, iCol + 1)  ' N.I.Area
            Else
                C3(4 + (spacer + fitNum), iCol + 5) = C1(10 + sftfit2, iCol + 1)      ' P.Area
                C3(4 + (spacer + fitNum) * 2, iCol + 5) = C1(11 + sftfit2, iCol + 1)  ' S.Area
                C3(4 + (spacer + fitNum) * 3, iCol + 5) = C1(12 + sftfit2, iCol + 1)  ' N.Area
            End If

            C3(3 + (spacer + fitNum) * 4, iCol + 5) = C1(1, iCol + 1)     ' Peak #5
            C3(4 + (spacer + fitNum) * 4, iCol + 5) = C1(4, iCol + 1)     ' FWHM
        Next

        For n = 0 To 4      ' n represents # of parameters to be summarized
            C3(3 + (spacer + fitNum) * n, 1) = "File"
            C3(3 + (spacer + fitNum) * n, 2) = "Sheet"
            C3(3 + (spacer + fitNum) * n, 4) = "# peaks"
            C3(4 + (spacer + fitNum) * n, 4) = sheetFit.Cells(8 + sftfit2, 2).Value
            C3(4 + (spacer + fitNum) * n, 1) = wb                  ' File name
            C3(4 + (spacer + fitNum) * n, 2) = strSheetFitName    ' Sheet name
            Range(Cells(3 + (spacer + fitNum) * n, 5), Cells(3 + (spacer + fitNum) * n, 4 + peakNum)).Interior.ColorIndex = 38
            Cells(3 + (spacer + fitNum) * n, 1).Interior.ColorIndex = 3
            Range(Cells(3 + (spacer + fitNum) * n, 2), Cells(3 + (spacer + fitNum) * n, 3)).Interior.ColorIndex = 4
            Cells(3 + (spacer + fitNum) * n, 4).Interior.ColorIndex = 33
            Range(Cells(3 + (spacer + fitNum) * n, peakNum + 6), Cells(3 + (spacer + fitNum) * n, peakNum + 7)).Interior.ColorIndex = 6
            Range(Cells(2 + (spacer + fitNum) * n, peakNum + 9), Cells(2 + (spacer + fitNum) * n, peakNum + 10)).Interior.ColorIndex = 8
        Next

        Cells(3 + (spacer + fitNum) * 2, 2 * peakNum + 9).Interior.ColorIndex = 26
        Cells(3 + (spacer + fitNum) * 3, 2 * peakNum + 9).Interior.ColorIndex = 26
        Range(Cells(3 + (spacer + fitNum) * 2, peakNum + 9), Cells(3 + (spacer + fitNum) * 2, 2 * peakNum + 8)).Interior.ColorIndex = 38
        Range(Cells(3 + (spacer + fitNum) * 3, peakNum + 9), Cells(3 + (spacer + fitNum) * 3, 2 * peakNum + 8)).Interior.ColorIndex = 38
        
        For n = 0 To 2
            C3(4, peakNum + 6 + n) = C2(1, 1 + n)                                   ' BG
        Next

        Results = "0," & strl(1) & "," & strl(2) & "," & strl(3) & ",,,"
        ncomp = 0
        cmp = 0     ' position of compared data to be added should be 0
        fcmp = C3   ' peak parameters from the base file, next to be added form the selected files
        
        Call EachComp(OpenFileName, strAna, fcmp, sBG, cmp, ncmp, ncomp)       ' Copy fitting parameters in each Fit sheet
        
        C3 = fcmp
        sheetAna.Activate
        sheetAna.Range(Cells(1, 1), Cells((4 + spacer * 4) + 5 * fitNum, 9 + 2 * peakNum)) = C3
        
        For n = 0 To fitNum - graphexist
            Cells(4 + n + spacer + fitNum, peakNum + 6).FormulaR1C1 = "=Sum(RC5:RC" & (peakNum + 4) & ")"                      ' Total P.Area
            Cells(4 + n + 2 * (spacer + fitNum), peakNum + 6).FormulaR1C1 = "=Sum(RC5:RC" & (peakNum + 4) & ")"                     ' Total S.Area
            Cells(4 + n + 3 * (spacer + fitNum), peakNum + 6).FormulaR1C1 = "=Sum(RC5:RC" & (peakNum + 4) & ")"                     ' Total N.Area
            Cells(4 + n + 4 * (spacer + fitNum), peakNum + 6).FormulaR1C1 = "=Average(RC5:RC" & (peakNum + 4) & ")"                 ' Avg FHHM
            For p = 0 To peakNum - 2
                Cells(4 + n, peakNum + 9 + p).FormulaR1C1 = "=(RC" & (6 + p) & " - RC" & (5 + p) & ")"                            ' Difference
                Cells(4 + n + spacer + fitNum, peakNum + 9 + p).FormulaR1C1 = "=(RC" & (5 + p) & " / RC" & (6 + p) & ")"    ' P.Area ratio
            Next
            
            For p = 0 To peakNum - 1
                Cells(4 + n + 2 * (spacer + fitNum), peakNum + 9 + p).FormulaR1C1 = "=(100 * RC" & (5 + p) & "/RC" & (peakNum + 6) & ")"  ' S.Area ratio
            Next
            
            Cells(4 + n + 2 * (spacer + fitNum), 2 * peakNum + 9).FormulaR1C1 = "=Sum(RC[" & (-peakNum) & "]:RC[-1])"               ' Total S.Area ratio
            
            For p = 0 To peakNum - 1
                Cells(4 + n + 3 * (spacer + fitNum), peakNum + 9 + p).FormulaR1C1 = "=(100 * RC" & (5 + p) & "/RC" & (peakNum + 6) & ")"  ' N.Area ratio
            Next
            
            Cells(4 + n + 3 * (spacer + fitNum), 2 * peakNum + 9).FormulaR1C1 = "=Sum(RC[" & (-peakNum) & "]:RC[-1])"               ' Total N.Area ratio
        Next
        
        For n = 0 To 4
            If n > 0 Then
                For k = 0 To peakNum - 1
                    Cells(3 + (spacer + fitNum) * n, k + 5).FormulaR1C1 = "=R3C" & (k + 5) & ""
                Next
            End If
            
            Set dataBGraph = Range(Cells(4 + (spacer + fitNum) * n, 5), Cells(4 + (spacer + fitNum) * n, 5).Offset(fitNum, Cells(4 + (spacer + fitNum) * n, 4) - 1))
            
            Charts.Add
            ActiveChart.ChartType = xlLineMarkers
            ActiveChart.SetSourceData Source:=dataBGraph, PlotBy:=xlColumns
            ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetAnaName
            
            For k = 1 To peakNum
                If IsEmpty(Cells(3, 4 + k).Value) = True Then
                Else
                    ActiveChart.SeriesCollection(k).Name = "='" & ActiveSheet.Name & "'!R3C" & (4 + k) & ""  ' Cells(3, 4 + k).Value
                    ActiveChart.SeriesCollection(k).AxisGroup = 1
                End If
            Next
            
            If Cells(4 + (spacer + fitNum) * n, 4).Value > 1 And n < 2 Then
                For k = 1 To peakNum - 1
                    Set dataKGraph = Range(Cells(4 + (spacer + fitNum) * n, peakNum + 9 + k - 1), Cells(4 + (spacer + fitNum) * n + fitNum, peakNum + 9 + k - 1))
                    ActiveChart.SeriesCollection.NewSeries
                    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
                        
                        .ChartType = xlColumnClustered
                        .Values = dataKGraph
                        If n = 0 Then
                            Cells(3, peakNum + 8 + k).FormulaR1C1 = "=R3C" & (5 + k) & " & ""-"" & R3C" & (4 + k) & ""
                            Cells(3, peakNum + 8 + k).Interior.ColorIndex = 38
                            .Name = "='" & ActiveSheet.Name & "'!R3C" & (peakNum + 8 + k) & ""             'Cells(3, 5 + k).Value + "-" + Cells(3, 4 + k).Value
                        ElseIf n = 1 Then
                            Cells((3 + (spacer + fitNum) * n), peakNum + 8 + k).FormulaR1C1 = "=R3C" & (4 + k) & " & ""/"" & R3C" & (5 + k) & ""
                            Cells((3 + (spacer + fitNum) * n), peakNum + 8 + k).Interior.ColorIndex = 38
                            .Name = "='" & ActiveSheet.Name & "'!R" & (3 + (spacer + fitNum) * n) & "C" & (peakNum + 8 + k) & ""            'Cells(3, 4 + k).Value + "/" + Cells(3, 5 + k).Value
                        End If
                        
                        .AxisGroup = 2
                    End With
                Next
            ElseIf Cells(4 + (spacer + fitNum) * n, 4).Value > 0 And n >= 2 And n <= 3 Then
                For k = 1 To peakNum
                    Set dataKGraph = Range(Cells(4 + (spacer + fitNum) * n, peakNum + 9 + k - 1), Cells(4 + (spacer + fitNum) * n + fitNum, peakNum + 9 + k - 1))
                    ActiveChart.SeriesCollection.NewSeries
                    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
                        .ChartType = xlAreaStacked100
                        Cells((3 + (spacer + fitNum) * n), peakNum + 8 + k).FormulaR1C1 = "= ""Rto_"" & R3C" & (4 + k) & ""
                        .Name = "='" & ActiveSheet.Name & "'!R" & (3 + (spacer + fitNum) * n) & "C" & (peakNum + 8 + k) & ""   'Cells(3, 4 + k).Value
                        .Values = dataKGraph
                        .AxisGroup = 2
                    End With
                Next
            End If
            
            With ActiveChart.Axes(xlCategory, xlPrimary)
                .HasTitle = True
                .AxisTitle.Text = "Samples"
                .AxisTitle.Font.Size = 12
                .AxisTitle.Font.Bold = False
            End With

            With ActiveChart.Axes(xlValue, xlPrimary)
                .HasTitle = True
                If n = 0 Then
                    .AxisTitle.Text = "Binding energy (eV)"
                ElseIf n = 1 Then
                    .AxisTitle.Text = "T.I. Area"
                ElseIf n = 2 Then
                    .AxisTitle.Text = "S.I. Area"
                ElseIf n = 3 Then
                    .AxisTitle.Text = "N.I. Area"
                ElseIf n = 4 Then
                    .AxisTitle.Text = "FWHM (eV)"
                End If
                .AxisTitle.Font.Size = 12
                .AxisTitle.Font.Bold = False
            End With
            
            If n < 3 And peakNum > 1 Then
                With ActiveChart.Axes(xlValue, xlSecondary)
                    .HasTitle = True
                    If n = 0 Then
                        .AxisTitle.Text = "Difference (eV)"
                    ElseIf n = 1 Then
                        .AxisTitle.Text = "Ratio (peak-to-peak)"
                    ElseIf n = 2 Then
                        .AxisTitle.Text = "Ratio (%)"
                    End If
                    .AxisTitle.Font.Size = 12
                    .AxisTitle.Font.Bold = False
                End With
            End If
        
            With ActiveSheet.ChartObjects(1 + n)
                .Top = 20 + (500 / (windowSize * 2)) * n
                .Left = 200 * 5
                .Width = (550 * windowRatio) / (windowSize * 2)
                .Height = 500 / (windowSize * 2)
                
                With .Chart.Legend
                    .Position = xlLegendPositionRight
                    .IncludeInLayout = True
                    .Left = (850 / (windowSize * 2))
                    .Top = (50 / (windowSize * 2))
                    With .Format.Fill
                        .Visible = msoTrue
                        .ForeColor.RGB = RGB(255, 255, 255)
                        .ForeColor.TintAndShade = 0.1
                    End With
                End With
                With .Chart
                    .PlotArea.Width = (((550 * windowRatio) - 100) / (windowSize * 2))
                    .ChartArea.Border.LineStyle = 0
                End With
            End With
        Next
        
        Cells(1, 1).Select
        strSheetAnaName = strSheetCmpName
        
        If ExistSheet(strSheetAnaName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetAnaName).Delete
            Application.DisplayAlerts = True
        End If
            
        Worksheets.Add().Name = strSheetAnaName
        Set sheetAna = Worksheets(strSheetAnaName)
        sheetFit.Activate
        numData = sheetFit.Cells(5, 101).Value
        imax = numData + 10
        
        C1 = sheetFit.Range(Cells(20 + sftfit, 1), Cells(20 + sftfit + numData, 1))    ' tmp
        C2 = sheetFit.Range(Cells(20 + sftfit, 4), Cells(20 + sftfit + numData, 4))     ' en

        sheetAna.Activate
        sheetAna.Range(Cells(10, 1), Cells(10 + numData, 1)) = C1
        sheetAna.Range(Cells(10, 3), Cells(10 + numData, 3)) = C2
        
        sheetGraph.Activate

        If IsEmpty(Cells(51, para + 10)) = False Then
            If Cells(42, para + 12) >= (Cells(43, para + 12) + Cells(42, para + 12)) Then
                sheetGraph.Range(Cells(40, para + 9), Cells((50 + Cells(42, para + 12).Value), para + 30)).Copy Destination:=sheetAna.Cells(40, para + 9)
            Else
                sheetGraph.Range(Cells(40, para + 9), Cells((50 + Cells(43, para + 12).Value + Cells(42, para + 12).Value), para + 30)).Copy Destination:=sheetAna.Cells(40, para + 9)
            End If
            
            sheetAna.Activate
            sheetAna.Cells(41, para + 10).Value = Application.Min(sheetAna.Range(Cells(11, 3), Cells(10 + numData, 3)))
            sheetAna.Cells(42, para + 10).Value = Application.Max(sheetAna.Range(Cells(11, 3), Cells(10 + numData, 3)))
            sheetAna.Cells(45, para + 10).Value = fitNum
        End If
        
        sheetAna.Activate
        Cells(1, 2).Value = wb
        Cells(9, 1).Value = "Offset/multp"
        Cells(9, 2).Value = 0
        Cells(9, 3).Value = 1
        
        If StrComp(mid$(Cells(10, 1), 1, 2), "BE", 1) = 0 Then
            strl(1) = "Be"
            strl(2) = "Sh"
            strl(3) = "In"
            
            If IsEmpty(Cells(4, 2)) Then
                Cells(4, 1) = "Shift"
                Cells(4, 2) = 0
                Cells(4, 3) = "eV"
                Cells(10, 2) = "Shift"
                Range(Cells(4, 1), Cells(4, 1)).Interior.ColorIndex = 3
                Range(Cells(4, 2), Cells(4, 3)).Interior.ColorIndex = 38
            End If
            
            Cells(11, 2).FormulaR1C1 = "=R4C + RC[-1]"
            Cells(10 + (imax), 2).FormulaR1C1 = "=R4C + R[-" & (imax - 1) & "]C[-1]"
        ElseIf StrComp(mid$(Cells(10, 1), 1, 2), "PE", 1) = 0 Then
            strl(1) = "Pe"
            strl(2) = "Sh"
            strl(3) = "Ab"
            
            If IsEmpty(Cells(2, 2)) Then
                Cells(2, 1) = "Shift"
                Cells(2, 2) = 0
                Cells(2, 3) = "eV"
                Cells(10, 2) = "Shift"
                Range(Cells(2, 1), Cells(2, 1)).Interior.ColorIndex = 3
                Range(Cells(2, 2), Cells(2, 3)).Interior.ColorIndex = 38
            End If
            
            Cells(11, 2).FormulaR1C1 = "=R2C + RC[-1]"
            Cells(10 + (imax), 2).FormulaR1C1 = "=R2C + R[-" & (imax - 1) & "]C[-1]"
        ElseIf StrComp(mid$(Cells(10, 1), 1, 2), "ME", 1) = 0 Then
            strl(1) = "Po"
            strl(2) = "Sh"
            strl(3) = "Ab"
            
            If IsEmpty(Cells(2, 2)) Then
                Cells(2, 1) = "Shift"
                Cells(2, 2) = 0
                Cells(2, 3) = "a.u."
                Cells(10, 2) = "Shift"
                Range(Cells(2, 1), Cells(2, 1)).Interior.ColorIndex = 3
                Range(Cells(2, 2), Cells(2, 3)).Interior.ColorIndex = 38
            End If
            
            Cells(11, 2).FormulaR1C1 = "=R2C + RC[-1]"
            Cells(10 + (imax), 2).FormulaR1C1 = "=R2C + R[-" & (imax - 1) & "]C[-1]"
        End If
        
        Range(Cells(11, 2), Cells((imax), 2)).FillDown
        Cells(10 + (imax), 1).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
        Range(Cells(10 + (imax), 1), Cells((2 * imax) - 1, 1)).FillDown
        Range(Cells(10 + (imax), 2), Cells((2 * imax) - 1, 2)).FillDown
        Cells(10 + (imax), 3).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C[-1])*R9C"
        Range(Cells(10 + (imax), 3), Cells((2 * imax) - 1, 3)).FillDown
        Cells(9, 1).Interior.Color = RGB(139, 195, 74)  ' added 20160324
        Range(Cells(9, 2), Cells(9, 3)).Interior.Color = RGB(197, 225, 165)  ' added 20160324
        Set dataBGraph = Range(Cells(10 + (imax), 2), Cells((2 * imax) - 1, 3))
        Dim SourceRangeColor1 As Long
        
        Charts.Add
        ActiveChart.ChartType = xlXYScatterLinesNoMarkers 'xlXYScatterSmoothNoMarkers
        ActiveChart.SetSourceData Source:=dataBGraph, PlotBy:=xlColumns
        ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetAnaName
        ActiveChart.SeriesCollection(1).Name = "='" & ActiveSheet.Name & "'!R1C2"   'ActiveWorkbook.Name  '"Fit sub BG" ' added 20160324
        ActiveChart.ChartTitle.Delete
        SourceRangeColor1 = ActiveChart.SeriesCollection(1).Border.Color
        
        With ActiveChart.Axes(xlCategory, xlPrimary)
            If StrComp(strl(1), "Pe", 1) = 0 Then
                .MinimumScale = startEb
                .MaximumScale = endEb
                strl(0) = "Photon energy (eV)"
            ElseIf StrComp(strl(1), "Po", 1) = 0 Then
                .MinimumScale = startEb
                .MaximumScale = endEb
                strl(0) = "Position (a.u.)"
            Else
                .MinimumScale = endEb
                .MaximumScale = startEb
                .ReversePlotOrder = True
                .Crosses = xlMaximum
                strl(0) = "Binding energy (eV)"
            End If
            .HasTitle = True
            .AxisTitle.Text = strl(0)
        End With
        
        With ActiveChart.Axes(xlCategory, xlPrimary)
            .MinorTickMark = xlOutside
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .HasMajorGridlines = True
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With ActiveChart.Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Intensity (arb. units)"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .MajorGridlines.Border.LineStyle = xlDot
            .Crosses = xlMinimum
        End With
    
        With ActiveSheet.ChartObjects(1)
            .Top = 20
            .Left = 200
            .Width = (550 * windowRatio) / windowSize
            .Height = 500 / windowSize
            With .Chart.Legend
                .Position = xlLegendPositionRight
                .IncludeInLayout = True
                .Left = (850 / windowSize)
                .Top = (50 / windowSize)
                With .Format.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(255, 255, 255)
                    .ForeColor.TintAndShade = 0.1
                End With
            End With
            With .Chart
                .PlotArea.Width = (((550 * windowRatio) - 40) / windowSize)
                .ChartArea.Border.LineStyle = 0
            End With
        End With
        
        Range(Cells(10, 2), Cells(10, 2)).Interior.Color = SourceRangeColor1
        Range(Cells(9 + (imax), 2), Cells(9 + (imax), 2)).Interior.Color = SourceRangeColor1
        
        strTest = mid$(Cells(1, 2).Value, 1, Len(Cells(1, 2).Value) - 5)
        Cells(8 + (imax), 2).Value = Cells(1, 2).Value
        Cells(9 + (imax), 1).Value = strl(2) + strTest
        Cells(9 + (imax), 2).Value = strl(3) + strTest
        Cells(9 + (imax), 3).Value = strl(4) + strTest
        
        strAna = "FitComp"
        Set sheetGraph = Worksheets(strSheetAnaName)
        Call PlotElem
        Call PlotChem
        Results = "0," & strl(1) & "," & strl(2) & "," & strl(3) & ",,,"
        
        Call EachComp(OpenFileName, strAna, fcmp, sBG, cmp, ncmp, ncomp)       ' Copy BG-substracted data in each Fit sheets.
        
        sheetAna.Activate
    Else
        TimeCheck = "stop"
    End If
    
    Call GetOut
End Sub

Sub SheetCheckGenerator()    ' Check scan grating data
    Dim C1 As Variant, C2 As Variant, C3 As Variant, dataCheck As Range, dataIntCheck As Range, strSheetCheckName As String, sheetCheck As Worksheet
    
    Worksheets.Add().Name = strSheetCheckName
    Set sheetCheck = Worksheets(strSheetCheckName)

    Cells(1, 1).Value = "X"
    Cells(1, 2).Value = "Y"
    Cells(1, 3).Value = "Norm"
    
    Set dataCheck = Range(Cells(2, 1), Cells(1 + numData, 2))
    Set dataIntCheck = Range(Cells(2, 3), Cells(1 + numData, 3))
        
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    ActiveChart.SetSourceData Source:=dataCheck, PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetCheckName
    ActiveChart.SeriesCollection(1).Border.ColorIndex = 41
    ActiveChart.SeriesCollection(1).Name = Cells(1, 2).Value
    ActiveChart.ChartTitle.Delete
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(2)
        .XValues = dataIntCheck.Offset(0, -2)
        .Values = dataIntCheck
        .AxisGroup = xlSecondary
        .Border.ColorIndex = 4
        .Name = Cells(1, 3).Value
    End With
    
    With ActiveSheet.ChartObjects(1)
        .Top = 20
        .Left = 200
        .Width = (550 * windowRatio) / windowSize
        .Height = 500 / windowSize
        .Chart.Legend.Delete
    End With

    With ActiveChart.Axes(xlCategory, xlPrimary)
        .MinorTickMark = xlOutside
        .MinimumScale = startEb
        .MaximumScale = endEb
        .HasTitle = True
        .AxisTitle.Text = strl(0)
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .HasMajorGridlines = True
        .MajorGridlines.Border.LineStyle = xlDot
    End With
    With ActiveChart.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Intensity (arb. unit)"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .MajorGridlines.Border.LineStyle = xlDot
        .MinimumScale = 0
    End With
    With ActiveChart.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = "Nomalized factor (arb.unit)"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
    End With
    
    Cells(1, 2).Interior.Color = ActiveChart.SeriesCollection(1).Border.Color
    Cells(1, 3).Interior.Color = ActiveChart.SeriesCollection(2).Border.Color
    
    Cells(1, 1).Select
    
    sheetGraph.Activate
End Sub

Sub FitInitial()
    Dim C1 As Variant, mySeries As Series, myChartOBJ As ChartObject
    
    If StrComp(strAna, "ana", 1) = 0 Or StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then
        Worksheets(strSheetGraphName).Activate
        Set sheetGraph = Worksheets(strSheetGraphName)
        numData = Cells(41, para + 12).Value '((Cells(6, 2).Value - Cells(5, 2).Value) / Cells(7, 2).Value) + 1
        Set dataBGraph = Range(Cells(20 + numData, 2), Cells(20 + numData, 2).Offset(numData - 1, 1))
        Set dataKeGraph = Range(Cells(20 + numData, 1), Cells(20 + numData, 1).Offset(numData - 1, 0))

        Call scalecheck
        If StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then
            Cells(10, 3).Value = "Ab"
        Else
            Cells(10, 3).Value = "In"
        End If
        If ExistSheet(strSheetFitName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetFitName).Delete
            Application.DisplayAlerts = True
        End If
    ElseIf StrComp(strl(3), "De", 1) = 0 Then
        If ExistSheet(strSheetFitName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetFitName).Delete
            Application.DisplayAlerts = True
        End If
        Call GetOut
        Exit Sub
    End If
    
    Worksheets.Add().Name = strSheetFitName
    Set sheetFit = Worksheets(strSheetFitName)
    
    Call descriptFit
    
    C1 = dataBGraph
    
    If sheetGraph.Cells(7, 2).Value > 0.001 Then
        For n = 1 To numData
            C1(n, 1) = Round(C1(n, 1), 3)   ' This makes round en off to third decimal places.
        Next
    End If
    
    Range(Cells(21 + sftfit, 1), Cells((numData + 20 + sftfit), 2)).Value = C1
    Set dataBGraph = Range(Cells(21 + sftfit, 1), Cells((numData + 20 + sftfit), 2))
    Set dataKGraph = Range(Cells(21 + sftfit, 1), Cells((numData + 20 + sftfit), 1))
    Set dataKeGraph = Range(Cells(11, 103), Cells(15, 104))
    
    If StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then
        Cells(11 + sftfit2, 2).Value = Cells(21 + sftfit, 1).Value
        Cells(12 + sftfit2, 2).Value = Cells(numData + 20 + sftfit, 1).Value
    Else
        Cells(11 + sftfit2, 2).Value = Cells(numData + 20 + sftfit, 1).Value
        Cells(12 + sftfit2, 2).Value = Cells(21 + sftfit, 1).Value
    End If
    
    'Charts.Add
    ActiveWorkbook.Charts.Add Before:=Worksheets(Worksheets.Count)  ' it makes no additional series in plot
    
    If Abs(startEb - endEb) < fitLimit Then
        ActiveChart.ChartType = xlXYScatter
    Else
        ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    End If
    ActiveChart.SetSourceData Source:=dataBGraph, PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetFitName
    ActiveChart.SeriesCollection(1).Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C2"

    ' new Errorplot
    
    'Charts.Add
    ActiveWorkbook.Charts.Add Before:=Worksheets(Worksheets.Count)
    If Abs(startEb - endEb) < fitLimit Then
        ActiveChart.ChartType = xlXYScatter
    Else
        ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    End If
    
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetFitName
    ActiveChart.SeriesCollection.NewSeries
    
    With ActiveChart.SeriesCollection(1)
        .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C4" ' In-BG
        .XValues = dataKGraph
        .Values = dataKGraph.Offset(, 3)
        .AxisGroup = xlPrimary
        .MarkerStyle = 8
        .MarkerSize = 8
        .MarkerForegroundColorIndex = 25    '16
        .MarkerBackgroundColorIndex = xlNone
    End With
    
    k = ActiveChart.SeriesCollection.Count
    For n = k To 2 Step -1
        ActiveChart.SeriesCollection(n).Delete
    Next
    
    For Each myChartOBJ In ActiveSheet.ChartObjects
        With myChartOBJ
            .Top = 20
            .Left = 500
            .Width = (550 * windowRatio) / windowSize
            .Height = 500 / windowSize
            .Chart.Legend.Delete
            .Chart.HasTitle = False
        End With

        With myChartOBJ.Chart.Axes(xlCategory, xlPrimary)
            .MinorTickMark = xlOutside
            .HasTitle = True
            If strl(1) = "Pe" Then
                .AxisTitle.Text = "Photon energy (eV)"
                .MinimumScale = startEb
                .MaximumScale = endEb
                .ReversePlotOrder = False
                .Crosses = xlMinimum
            ElseIf strl(1) = "Po" Then
                .AxisTitle.Text = "Position (a.u.)"
                .MinimumScale = startEb
                .MaximumScale = endEb
                .ReversePlotOrder = False
                .Crosses = xlMinimum
            Else
                .AxisTitle.Text = "Binding energy (eV)"
                .MinimumScale = endEb
                .MaximumScale = startEb
                .ReversePlotOrder = True
                .Crosses = xlMaximum
            End If
            
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .HasMajorGridlines = True
            If numMajorUnit <> 0 Then
                .MajorUnit = numMajorUnit
            Else
                .MinimumScaleIsAuto = True
                .MaximumScaleIsAuto = True
            End If
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        
        With myChartOBJ.Chart.Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = "Intensity normalized by Ip (arb. units)"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .MajorGridlines.Border.LineStyle = xlDot
            .MinimumScale = dblMin - (dblMax - dblMin) * 0.02
            .MaximumScale = dblMax + (dblMax - dblMin) * 0.1
        End With
        With myChartOBJ.Chart
            .ChartArea.Border.LineStyle = 0
        End With
        
        For Each mySeries In myChartOBJ.Chart.SeriesCollection
            mySeries.Format.Line.Weight = 1
            If Abs(startEb - endEb) < fitLimit Then
                mySeries.ChartType = xlXYScatter
                mySeries.MarkerStyle = 8
                mySeries.MarkerSize = 8
                mySeries.MarkerForegroundColorIndex = 1
                mySeries.MarkerBackgroundColorIndex = xlNone
            Else
                mySeries.ChartType = xlXYScatterLinesNoMarkers
                mySeries.Border.ColorIndex = 1
                mySeries.Border.Weight = xlThin
                mySeries.Border.LineStyle = xlContinuous
            End If
        Next
    Next
    
    If ActiveSheet.ChartObjects.Count > 1 Then
        With ActiveSheet.ChartObjects(2)
            '.Top = 600 / windowSize
            .Top = 1 * (500 / windowSize) + 20
            .Height = 250 / windowSize
            With .Chart.Axes(xlValue, xlPrimary)
                .AxisTitle.Text = "BG-subtracted intensity (arb. units)"
                .MinimumScaleIsAuto = True
                .MaximumScaleIsAuto = True
            End With
        End With
    End If
    
    ActiveSheet.ChartObjects(2).Activate

    If Abs(startEb - endEb) < fitLimit Then
         With ActiveChart.SeriesCollection(1)
            .MarkerStyle = 8
            .MarkerSize = 5
            .MarkerForegroundColorIndex = xlNone
            .MarkerBackgroundColorIndex = 38    'xlNone
        End With
    Else
        With ActiveChart.SeriesCollection(1)
            .Border.ColorIndex = 1
            .Border.Weight = xlThin
            .Border.LineStyle = xlContinuous
        End With
    End If

    ' new Boxplot
    ActiveWorkbook.Charts.Add Before:=Worksheets(Worksheets.Count)
    ActiveChart.SetSourceData Source:=dataKeGraph, PlotBy:=1
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetFitName
    ActiveChart.ChartType = xlStockOHLC
    
    ActiveChart.SeriesCollection(1).Name = "Q3"
    ActiveChart.SeriesCollection(2).Name = "max"
    ActiveChart.SeriesCollection(3).Name = "min"
    ActiveChart.SeriesCollection(4).Name = "Q1"
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(5)
        .ChartType = xlXYScatter
        .Name = "='" & ActiveSheet.Name & "'!R16C102"
        .XValues = Cells(11, 103)
        .Values = Cells(16, 103)
        .AxisGroup = xlSecondary
        .MarkerStyle = 7
        .MarkerSize = 16
        .MarkerForegroundColorIndex = xlNone    '16
        .MarkerBackgroundColorIndex = 3
    End With
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(6)
        .ChartType = xlXYScatter
        .Name = "='" & ActiveSheet.Name & "'!R13C102"
        .XValues = Cells(11, 103)
        .Values = Cells(13, 103)
        .AxisGroup = xlSecondary
        .MarkerStyle = 7
        .MarkerSize = 16
        .MarkerForegroundColorIndex = xlNone    '16
        .MarkerBackgroundColorIndex = 1
    End With
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(7)
        .ChartType = xlXYScatter
        .Name = "='" & ActiveSheet.Name & "'!R14C102"
        .XValues = Cells(11, 103)
        .Values = Cells(14, 103)
        .AxisGroup = xlSecondary
        .MarkerStyle = 7
        .MarkerSize = 16
        .MarkerForegroundColorIndex = xlNone    '16
        .MarkerBackgroundColorIndex = 1
    End With
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(8)
        .ChartType = xlXYScatter
        .Name = "='" & ActiveSheet.Name & "'!R17C102"
        .XValues = Cells(11, 103)
        .Values = Cells(17, 103)
        .AxisGroup = xlSecondary
        .MarkerStyle = 9
        .MarkerSize = 16
        .MarkerForegroundColorIndex = 1    '16
        .MarkerBackgroundColorIndex = xlNone
    End With
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(9)
        .ChartType = xlXYScatter
        .Name = "='" & ActiveSheet.Name & "'!R18C102"
        .XValues = Cells(11, 103)
        .Values = Cells(18, 103)
        .AxisGroup = xlSecondary
        .MarkerStyle = 8
        .MarkerSize = 16
        .MarkerForegroundColorIndex = 1    '16
        .MarkerBackgroundColorIndex = xlNone
    End With
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(10)
        .ChartType = xlXYScatter
        .Name = "='" & ActiveSheet.Name & "'!R19C102"
        .XValues = Cells(11, 103)
        .Values = Cells(19, 103)
        .AxisGroup = xlSecondary
        .MarkerStyle = 8
        .MarkerSize = 16
        .MarkerForegroundColorIndex = 1    '16
        .MarkerBackgroundColorIndex = xlNone
    End With
    
    Range(Cells(11, 104), Cells(16, 104)).Delete

    With ActiveSheet.ChartObjects(3)
        .Top = 1 * (500 / windowSize) + 20
        .Height = 250 / windowSize
        .Left = 500 + (550 * windowRatio) / windowSize
        .Width = (50 * windowRatio) / windowSize
        .Chart.Legend.Delete
        .Chart.ChartArea.Border.LineStyle = 0
        .Chart.HasTitle = False
        .Chart.ChartGroups(1).HiLoLines.Border.ColorIndex = 1
        .Chart.ChartGroups(1).HiLoLines.Format.Line.Weight = 2
        .Chart.ChartGroups(1).DownBars.Border.ColorIndex = 1
        .Chart.ChartGroups(1).DownBars.Interior.ColorIndex = 6
        .Chart.ChartGroups(1).DownBars.Format.Line.Weight = 2
        '.Chart.Axes(xlCategory).Delete
        With .Chart.Axes(xlCategory)
            .HasTitle = True
            .AxisTitle.Text = "Boxplot"
            .TickLabelPosition = xlNone
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With .Chart.Axes(xlValue, xlSecondary)
            .MinimumScale = ActiveChart.Axes(xlValue, xlPrimary).MinimumScale
            .MaximumScale = ActiveChart.Axes(xlValue, xlPrimary).MaximumScale
            .Delete
        End With
    End With
    
    Call GetOutFit
End Sub

Sub FitInitialGuess()
    Dim obchk As String, obchk2 As String, dbltchk As Single, dblt As Integer, C1 As Variant, C2 As Variant, C3 As Variant
    
    dblt = 0
    dbltchk = 0
    sheetGraph.Activate
    
    numXPSFactors = Cells(43, para + 12).Value
    C1 = Range(Cells(51, para + 10), Cells((51 + numXPSFactors), para + 12)) ' peak name and BE
    C2 = Range(Cells(51, para + 16), Cells((51 + numXPSFactors), para + 19)) ' Amp and sensitivity
    sheetFit.Activate
    C3 = Range(Cells(1, 5), Cells(15 + sftfit2 + 3, (numXPSFactors + 5)))
    
    For n = numXPSFactors To 1 Step -1
        If StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then
            If C1(n, 3) > Cells(startR, 1).Value And C1(n, 3) < Cells(endR, 1).Value Then
                j = j + 1
                C3(1, j) = C1(n, 2)
                
                If Len(C1(n, 2)) - Len(C1(n, 1)) > 2 Then
                    obchk = mid$(C1(n, 2), Len(C1(n, 1)) + 2, 2)
                    If StrComp(obchk, "p3", 1) = 0 Then
                        dbltchk = C1(n, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "p3", 1) = 0 And StrComp(obchk, "p1", 1) = 0 Then
                        C3(19, dblt) = "(2;"
                        C3(19, j) = "1)"
                        C3(20, dblt) = "["
                        If dbltchk <= C1(n, 3) Then
                            dbltchk = C1(n, 3) - dbltchk
                            C3(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C1(n, 3)
                            C3(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf StrComp(obchk, "d5", 1) = 0 Then
                        dbltchk = C1(n, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "d5", 1) = 0 And StrComp(obchk, "d3", 1) = 0 Then
                        C3(19, dblt) = "(3;"
                        C3(19, j) = "2)"
                        C3(20, dblt) = "["
                        If dbltchk <= C1(n, 3) Then
                            dbltchk = C1(n, 3) - dbltchk
                            C3(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C1(n, 3)
                            C3(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf StrComp(obchk, "f7", 1) = 0 Then
                        dbltchk = C1(n, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "f7", 1) = 0 And StrComp(obchk, "f5", 1) = 0 Then
                        C3(19, dblt) = "(4;"
                        C3(19, j) = "3)"
                        C3(20, dblt) = "["
                        If dbltchk <= C1(n, 3) Then
                            dbltchk = C1(n, 3) - dbltchk
                            C3(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C1(n, 3)
                            C3(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf dbltchk <> 0 And obchk2 <> vbNullString Then
                        C3(19, dblt) = vbNullString
                        C3(20, dblt) = vbNullString
                        dbltchk = 0
                        obchk2 = vbNullString
                        dblt = 0
                    End If
                End If
                
                C3(2, j) = C1(n, 3)
                C3(6, j) = C2(n, 2)
                C3(9 + sftfit2, j) = C2(n, 1)
                C3(7 + sftfit2, j) = C2(n, 4) ' beta
            End If
        Else
            If C1(n, 3) < Cells(startR, 1).Value And C1(n, 3) > Cells(endR, 1).Value Then
                j = j + 1
                C3(1, j) = C1(n, 2)   ' peak name
                
                If Len(C1(n, 2)) - Len(C1(n, 1)) > 2 Then
                    obchk = mid$(C1(n, 2), Len(C1(n, 1)) + 2, 2)
                    If StrComp(obchk, "p3", 1) = 0 Then
                        dbltchk = C1(n, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "p3", 1) = 0 And StrComp(obchk, "p1", 1) = 0 Then
                        C3(19, dblt) = "(2;"
                        C3(19, j) = "1)"
                        C3(20, dblt) = "["
                        If dbltchk <= C1(n, 3) Then
                            dbltchk = C1(n, 3) - dbltchk
                            C3(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C1(n, 3)
                            C3(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf StrComp(obchk, "d5", 1) = 0 Then
                        dbltchk = C1(n, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "d5", 1) = 0 And StrComp(obchk, "d3", 1) = 0 Then
                        C3(19, dblt) = "(3;"
                        C3(19, j) = "2)"
                        C3(20, dblt) = "["
                        If dbltchk <= C1(n, 3) Then
                            dbltchk = C1(n, 3) - dbltchk
                            C3(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C1(n, 3)
                            C3(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf StrComp(obchk, "f7", 1) = 0 Then
                        dbltchk = C1(n, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "f7", 1) = 0 And StrComp(obchk, "f5", 1) = 0 Then
                        C3(19, dblt) = "(4;"
                        C3(19, j) = "3)"
                        C3(20, dblt) = "["
                        If dbltchk <= C1(n, 3) Then
                            dbltchk = C1(n, 3) - dbltchk
                            C3(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C1(n, 3)
                            C3(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf dbltchk <> 0 And obchk2 <> vbNullString Then
                        C3(19, dblt) = vbNullString
                        C3(20, dblt) = vbNullString
                        dbltchk = 0
                        obchk2 = vbNullString
                        dblt = 0
                    End If
                End If
                
                C3(2, j) = C1(n, 3)   ' BE
                C3(6, j) = C2(n, 2)   ' Amp.
                C3(9 + sftfit2, j) = C2(n, 1) ' sensitivity
                C3(7 + sftfit2, j) = C2(n, 4) ' beta
            End If
        End If
    Next
    
    Range(Cells(1, 5), Cells(15 + sftfit2 + 3, (numXPSFactors + 5))) = C3

    If j > 0 Then
        Range(Cells(4, 5), Cells(4, (4 + j))) = 2
        Range(Cells(5, 5), Cells(5, (4 + j))) = vbNullString
        Range(Cells(7, 5), Cells(7, (4 + j))) = "Gauss"
        Range(Cells(1, 5), Cells(15 + sftfit2 + 4, (4 + j))).Interior.Color = RGB(178, 235, 242) '34
    ElseIf testMacro = "debug" And j = 0 Then      ' this is for the case if no database found and continue processing
        j = 1
        Range(Cells(4, 5), Cells(4, (4 + j))) = 2
        Range(Cells(5, 5), Cells(5, (4 + j))) = vbNullString
        Range(Cells(7, 5), Cells(7, (4 + j))) = "Gauss"
        Range(Cells(1, 5), Cells(15 + sftfit2 + 4, (4 + j))).Interior.Color = RGB(178, 235, 242) '34
        Cells(1, 5) = "noid"
        Cells(2, 5) = (Cells(11 + sftfit2, 2) + Cells(12 + sftfit2, 2)) / 2
        Cells(6, 5) = (Cells(3, 101) - Cells(2, 101))
        Cells(9 + sftfit2, 5) = 1
    ElseIf Not testMacro = "debug" Then
        TimeCheck = MsgBox("No peak in the range! Would you like to fit a peak anyway?", 4, "Fitting error")
        If TimeCheck = 6 Then
            j = 1
            Range(Cells(4, 5), Cells(4, (4 + j))) = 2
            Range(Cells(5, 5), Cells(5, (4 + j))) = vbNullString
            Range(Cells(7, 5), Cells(7, (4 + j))) = "Gauss"
            Range(Cells(1, 5), Cells(15 + sftfit2 + 4, (4 + j))).Interior.Color = RGB(178, 235, 242) '34
            Cells(1, 5) = "noid"
            Cells(2, 5) = (Cells(11 + sftfit2, 2) + Cells(12 + sftfit2, 2)) / 2
            Cells(6, 5) = (Cells(3, 101) - Cells(2, 101))
            Cells(9 + sftfit2, 5) = 1
        Else
            TimeCheck = 0
            j = 0
            Cells(8, 101).Value = 0     ' -1
            Range(Cells(1, 4), Cells(15 + sftfit2 + 4, 55)).ClearContents
            Range(Cells(20 + sftfit, 4), Cells((2 * numData + 22 + sftfit), 55)).ClearContents
            Range(Cells(1, 4), Cells(19 + sftfit2 + 3, 55)).Interior.ColorIndex = xlNone
            Cells(20 + sftfit, 3).Value = "BG"
            Call GetOutFit
            strErrX = "skip"
            Exit Sub
        End If
    End If
    
    Cells(8 + sftfit2, 2).Value = j
    Cells(9, 101).Value = j
End Sub

Sub FitRange(ByRef strCpa As String)
    Dim C1 As Variant, C2 As Variant, rng As Range, numDataN As Integer, myChartOBJ As ChartObject
    
    strSheetGraphName = "Graph_" + strSheetDataName

    If StrComp(mid$(strMode, 8, 5), "range", 1) = 0 Then
        strSheetFitName = ActiveSheet.Name
    Else
        strSheetFitName = "Fit_" + strSheetDataName
    End If
    
    dblMin = Cells(2, 101).Value
    dblMax = Cells(3, 101).Value
    numXPSFactors = Cells(4, 101).Value
    numData = Cells(5, 101).Value
    startEb = Cells(6, 101).Value
    endEb = Cells(7, 101).Value
    cae = Cells(14 + sftfit2, 2).Value
    
    If IsEmpty(Cells(9, 103).Value) Then
        If WorksheetFunction.Round(Cells(12, 101).Value, 1) = 1486.6 Then
            Cells(9, 103).Value = "MultiPak"
        Else
            Cells(9, 103).Value = "Sum"
        End If
    ElseIf LCase(Cells(9, 103).Value) = "multipak" Then
        Cells(9, 103).Value = "MultiPak"
    ElseIf LCase(Cells(9, 103).Value) = "product" Then
        Cells(9, 103).Value = "Product"
    Else
        Cells(9, 103).Value = "Sum"
    End If
    
    pe = Cells(12, 101).Value
    wf = Cells(13, 101).Value
    char = Cells(14, 101).Value
    ns = Cells(10, 101).Value
    
    If StrComp(LCase(Worksheets(strSheetGraphName).Cells(10, 3).Value), "ab", 1) = 0 And StrComp(LCase(Worksheets(strSheetGraphName).Cells(10, 1).Value), "pe", 1) = 0 Then
        strl(1) = "Pe"
    ElseIf StrComp(LCase(Worksheets(strSheetGraphName).Cells(10, 1).Value), "po", 1) = 0 Then
        strl(1) = "Po"
    End If
    
    If Abs(startEb - endEb) > fitLimit Then
        If StrComp(testMacro, "debug", 1) = 0 Then  ' debug mode skip fitting in the specific range.
            TimeCheck = 0
            Call GetOutFit
            strErrX = "skip"
            Exit Sub
        End If

        ElemD = Application.InputBox(Title:="Specify the fitting range", Prompt:="Input BE energy: 320-350eV", Default:="320-350eV", Type:=2)

        If ElemD = "False" Or Len(ElemD) = 0 Then
            TimeCheck = 0
            Call GetOutFit
            strErrX = "skip"
            Exit Sub
        Else
            C1 = Split(ElemD, "-")
            If IsNumeric(mid$(C1(1), 1, Len(C1(1)) - 2)) = True Then
                If mid$(C1(1), 1, Len(C1(1)) - 2) < startEb And mid$(C1(1), 1, Len(C1(1)) - 2) > endEb Then
                    startEb = mid$(C1(1), 1, Len(C1(1)) - 2)
                Else
                    'GoTo GetOutFit
                End If
            Else
                TimeCheck = MsgBox("BE range format is not appropriate!")
                Call GetOutFit
                strErrX = "skip"
                Exit Sub
            End If
            
            If IsNumeric(C1(0)) = True Then
                If C1(0) < startEb And C1(0) > endEb Then
                    endEb = C1(0)
                ElseIf C1(0) > startEb Then
                    startEb = C1(0)
                    endEb = mid$(C1(1), 1, Len(C1(1)) - 2)
                Else
                    TimeCheck = MsgBox("BE range is not in the scanned range!")
                    Call GetOutFit
                    strErrX = "skip"
                    Exit Sub
                End If
            Else
                TimeCheck = MsgBox("BE range format is not appropriate!")
                Call GetOutFit
                strErrX = "skip"
                Exit Sub
            End If
            
            Dim flag As Boolean
            flag = False
            
            For Each sheetFit In Worksheets
                If sheetFit.Name = "Fit_BE" + ElemD Then flag = True
            Next sheetFit
            If flag = True Then
                TimeCheck = MsgBox("The sheet already exists!")
                Worksheets("Fit_BE" + ElemD).Activate
                Call GetOutFit
                strErrX = "skip"
                Exit Sub
            End If
            
            Worksheets(strSheetFitName).Copy Before:=Worksheets(strSheetFitName)
            ActiveSheet.Name = "Fit_BE" + ElemD
            strSheetFitName = "Fit_BE" + ElemD
            
            Cells(1, 100).Value = "Source"
            Cells(1, 101).Value = strSheetDataName
            strSheetGraphName = "Graph_" + Cells(1, 101).Value
            Cells(11 + sftfit2, 2).Value = endEb    ' endEb < startEb
            Cells(12 + sftfit2, 2).Value = startEb
            
            If Abs(startEb - endEb) <= 100 And Abs(startEb - endEb) > 50 Then
                numMajorUnit = 4 * windowSize
            ElseIf Abs(startEb - endEb) <= 50 And Abs(startEb - endEb) > 20 Then
                numMajorUnit = 2 * windowSize
            ElseIf Abs(startEb - endEb) > 100 Then
                numMajorUnit = 50 * windowSize
            ElseIf Abs(startEb - endEb) <= 20 And Abs(startEb - endEb) > 1 Then
                numMajorUnit = 1 * windowSize
            ElseIf Abs(startEb - endEb) <= 1 Then
                numMajorUnit = 0
            End If
            
            If numMajorUnit = 0 Then
            ElseIf StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(3), "De", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then
                startEb = Application.Floor(startEb, numMajorUnit)
            ElseIf startEb > 0 Then
                startEb = Application.Ceiling(startEb, numMajorUnit)
            Else
                startEb = Application.Floor(startEb, (-1 * numMajorUnit))
            End If
            
            If numMajorUnit = 0 Then
            ElseIf StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(3), "De", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then
                endEb = Application.Ceiling(endEb, numMajorUnit)
            ElseIf endEb > 0 Then
                endEb = Application.Floor(endEb, numMajorUnit)
            Else
                endEb = Application.Ceiling(endEb, (-1 * numMajorUnit))
            End If
            
            Cells(6, 101).Value = startEb
            Cells(7, 101).Value = endEb
            Cells(10, 102).Value = "majorUnit"
            Cells(10, 103).Value = numMajorUnit
            If Abs(startEb - endEb) / Abs(Cells(22 + sftfit, 1).Value - Cells(21 + sftfit, 1).Value) < 30 Then
                Cells(10, 101).Value = 3     ' Average # points for Solver around startR and endR points
            ElseIf Abs(startEb - endEb) / Abs(Cells(22 + sftfit, 1).Value - Cells(21 + sftfit, 1).Value) < 60 Then
                Cells(10, 101).Value = 5
            Else
                Cells(10, 101).Value = 10
            End If
            
            If strl(1) = "Pe" Then             ' additional BE step
                Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (4)
                Cells(2, 103).Value = 5       ' max FWHM1 limit
            Else
                Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (20)
            End If
        End If
    End If
    
    If Cells(8 + sftfit2, 2).Value > 29 Then
        TimeCheck = MsgBox("# of peaks are over 30! Would you like to continue anyway?", 4, "Fitting suggestion")
        If TimeCheck = 6 Then
        Else
            TimeCheck = 0
            j = 0
            Call GetOutFit
            strErrX = "skip"
            Exit Sub
        End If
    End If
    
    Set sheetGraph = Worksheets(strSheetGraphName)
    Set sheetFit = Worksheets(strSheetFitName)
    
    sheetFit.Activate
    C1 = sheetFit.Range(Cells(21 + sftfit, 1), Cells((numData + 20 + sftfit), 1))

    k = 0
    j = 0
    
    If StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then
        For n = 1 To numData - 1
            If Cells(11 + sftfit2, 2) >= C1(n, 1) And Cells(11 + sftfit2, 2) < C1((n + 1), 1) Then
                startR = n + 20 + sftfit
            End If
        Next
        For n = 2 To numData
            If Cells(12 + sftfit2, 2) <= C1(n, 1) And Cells(12 + sftfit2, 2) > C1((n - 1), 1) Then
                endR = n + 20 + sftfit
            End If
        Next
        
        If startR < 1 Or CStr(startR) = vbNullString Then
            startR = 21 + sftfit
            Cells(11 + sftfit2, 2).Value = Cells(21 + sftfit, 1).Value
        End If
        
        If endR > numData + 20 + sftfit Or endR < startR Or CStr(endR) = vbNullString Then
            endR = numData + 20 + sftfit
            Cells(12 + sftfit2, 2).Value = Cells(numData + 20 + sftfit, 1).Value
        End If
        
    Else
        For n = 1 To numData - 1
            If Cells(12 + sftfit2, 2) <= C1(n, 1) And Cells(12 + sftfit2, 2) > C1((n + 1), 1) Then
                startR = n + 20 + sftfit
            End If
        Next
        For n = 2 To numData
            If Cells(11 + sftfit2, 2) >= C1(n, 1) And Cells(11 + sftfit2, 2) < C1((n - 1), 1) Then
                endR = n + 20 + sftfit
            End If
        Next
        
        If startR < 1 Or CStr(startR) = vbNullString Then
            startR = 21 + sftfit
            Cells(12 + sftfit2, 2).Value = Cells(21 + sftfit, 1).Value
        End If
        
        If endR > numData + 20 + sftfit Or endR < startR Or CStr(endR) = vbNullString Then
            endR = numData + 20 + sftfit
            Cells(11 + sftfit2, 2).Value = Cells(numData + 20 + sftfit, 1).Value
        End If
    
    End If
    
    numDataN = endR - startR + 1
    
    C1 = Range(Cells(startR, 2), Cells(endR, 2))    'C
    C2 = Range(Cells(startR, 3), Cells(endR, 3))    'A
    C2(numDataN, 1) = C1(numDataN, 1)
    C2((numDataN - 1), 1) = C1(numDataN, 1)
    
    If IsEmpty(Cells(1, 101).Value) = False Then    ' range > fitLimit eV
        Cells(2, 101).Value = Application.Min(C1)
        Cells(3, 101).Value = Application.Max(C1)
        dblMin = Cells(2, 101).Value - ((Cells(3, 101).Value - Cells(2, 101).Value) / 100)
        dblMax = Cells(3, 101).Value + ((Cells(3, 101).Value - Cells(2, 101).Value) / 10)
        n = 0
        
        For Each myChartOBJ In ActiveSheet.ChartObjects
            n = n + 1
            If n = 1 Then
                With myChartOBJ.Chart.Axes(xlCategory, xlPrimary)
                    .MinimumScale = Cells(7, 101).Value
                    .MaximumScale = Cells(6, 101).Value
                    .MajorUnit = Cells(10, 103).Value
                End With
                With myChartOBJ.Chart.Axes(xlValue)
                    .MinimumScale = dblMin
                    .MaximumScale = dblMax
                End With
            ElseIf n = 2 Then
                With myChartOBJ.Chart.Axes(xlCategory, xlPrimary)
                    .MinimumScale = Cells(7, 101).Value
                    .MaximumScale = Cells(6, 101).Value
                    .MajorUnit = Cells(10, 103).Value
                End With
                Exit For
            End If
        Next
    End If
    
    strBG1 = LCase(mid$(Cells(1, 1).Value, 1, 1))
    strBG2 = LCase(mid$(Cells(1, 2).Value, 1, 1))
    
    If strBG1 = LCase(mid$(Cells(16, 100).Value, 1, 1)) And strBG2 = LCase(mid$(Cells(16, 101).Value, 1, 1)) Then
        If Cells(8, 101).Value = 0 Then
            strCpa = "initial"
        Else
            strCpa = "repeat"
        End If
    Else
        If Cells(8, 101).Value = 0 Then
            strCpa = "initial"
        Else
            Cells(8, 101).Value = 0
            strCpa = "repeat"
        End If
    End If
    
    For Each rng In Range(Cells(2, 3), Cells(7 + sftfit2, 4)).Cells
        If IsNumeric(rng.Value) = False Then
            rng.Value = vbNullString
        End If
    Next
End Sub

Sub FormulaCheck()
    Dim numbra1 As Integer, numbra2 As Integer, numbran1 As Integer, numbran2 As Integer, cnt As Integer
    
    cnt = 0
    
recheckformua1:

    numbra1 = 0
    numbra2 = 0
    
    For n = 5 To (4 + j)
        If Not Cells(14 + sftfit2, n) = vbNullString Then
            If InStr(1, Cells(14 + sftfit2, n), "(", 1) > 0 Then
                numbra1 = numbra1 + 1
                numbran1 = n
            ElseIf InStr(1, Cells(14 + sftfit2, n), ")", 1) > 0 Then
                numbra2 = numbra2 + 1
                numbran2 = n
            End If
        End If
    Next
    
    If numbra1 <> numbra2 Then
        If numbra1 > numbra2 Then
            Debug.Print "non match (>)"
            Cells(14 + sftfit2, numbran1) = vbNullString
        ElseIf numbra1 < numbra2 Then
            Debug.Print "non match (<)"
            Cells(14 + sftfit2, numbran2) = vbNullString
        End If
        
        cnt = cnt + 1
        If cnt < 10 Then GoTo recheckformua1
    Else
        Debug.Print "match (=)"
    End If
    
recheckformua2:

    numbra1 = 0
    numbra2 = 0
    
    For n = 5 To (4 + j)
        If Not Cells(15 + sftfit2, n) = vbNullString Then
            If InStr(1, Cells(15 + sftfit2, n), "[", 1) > 0 Then
                numbra1 = numbra1 + 1
                numbran1 = n
            ElseIf InStr(1, Cells(15 + sftfit2, n), "]", 1) > 0 Then
                numbra2 = numbra2 + 1
                numbran2 = n
            End If
        End If
    Next
    
    If numbra1 <> numbra2 Then
        If numbra1 > numbra2 Then
            Debug.Print "non match [>]"
            Cells(15 + sftfit2, numbran1) = vbNullString
        ElseIf numbra1 < numbra2 Then
            Debug.Print "non match [<]"
            Cells(15 + sftfit2, numbran2) = vbNullString
        End If
        
        cnt = cnt + 1
        If cnt < 10 Then GoTo recheckformua2
    Else
        Debug.Print "match [=]"
    End If
        
End Sub

Sub FitCurve()
    Application.Calculation = xlCalculationManual
    
    Dim ls As Single, ratio1 As Single, imax As Integer, rng As Range, strCpa As String
    
    If StrComp(mid$(strMode, 1, 6), "Do fit", 1) = 0 Then
    Else
        Call FitInitial
        Exit Sub
    End If

    If IsEmpty(Cells(19, 101).Value) Then
        MsgBox "VBA code version analyzed in the sheet is too old, regenerate the fit sheet from graph sheet again.", vbInformation
        Call GetOut
        Exit Sub
    End If

    Call FitRange(strCpa)
    If strErrX = "skip" Then Exit Sub
    Call SolverSetup
    
    If StrComp(strBG1, "t", 1) <> 0 And StrComp(strBG2, "t", 1) <> 0 And StrComp(strBG1, "e", 1) <> 0 Then
        Range("DG31").CurrentRegion.ClearContents
        Range("DE31").CurrentRegion.ClearContents
    End If
        
    If Cells(11 + sftfit2, 2).Value < 0 And Cells(12 + sftfit2, 2).Value > 0 And (Cells(12 + sftfit2, 2).Value - Cells(11 + sftfit2, 2).Value) < 6 And mid$(Cells(sftfit2 + 25, 1).Value, 1, 1) = "B" Then
        Call FitEF
        Call GetOutFit
        Exit Sub
    ElseIf StrComp(strBG1, "p", 1) = 0 Then
        If StrComp(strBG2, "s", 1) = 0 Then
            Call PolynominalShirleyBG
        ElseIf StrComp(strBG2, "t", 1) = 0 Then
            Call PolynominalTougaardBG
        Else
            Call PolynominalBG
        End If
    ElseIf StrComp(strBG1, "a", 1) = 0 Then
        Call TangentArcBG
    ElseIf StrComp(strBG1, "t", 1) = 0 Then
        Call TougaardBG
    ElseIf StrComp(strBG1, "v", 1) = 0 Then
        Call VictoreenBG
    ElseIf StrComp(strBG1, "e", 1) = 0 Then
        Call FitEF
        Call GetOutFit
        Exit Sub
    Else
        Call ShirleyBG
    End If
    
    Cells(startR, 4).FormulaR1C1 = "=RC[-2] - RC[-1]"
    Range(Cells(startR, 4), Cells(endR, 4)).FillDown
    
    If startR > 21 + sftfit Then
        Range(Cells(21 + sftfit, 3), Cells(startR - 1, 4)).ClearContents
    End If
    
    If endR < numData + 20 + sftfit Then
        Range(Cells(endR + 2, 3), Cells(numData + 20 + sftfit, 4)).ClearContents
    End If
    
    Set rng = Range(Cells(startR, 1), Cells(endR, 1))
    
    ActiveSheet.ChartObjects(1).Activate
    
    k = ActiveChart.SeriesCollection.Count  ' delete previous data
    For n = k To 2 Step -1
        ActiveChart.SeriesCollection(n).Delete
    Next

    ActiveChart.SeriesCollection.NewSeries

    With ActiveChart.SeriesCollection(2)
        .ChartType = xlXYScatterLinesNoMarkers
        .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C3"
        .XValues = rng
        .Values = rng.Offset(0, 2)
        .Format.Line.Weight = 2
        .Border.ColorIndex = 10
    End With
    
    If strCpa = "repeat" Then
        
    Else
        Call FitInitialGuess
        If strErrX = "skip" Then Exit Sub
    End If

    Call FitEquations
    
    j = Cells(8 + sftfit2, 2).Value 'npa
    Call FormulaCheck
    ActiveSheet.Calculate
    
AsymIteration:
    If IsNumeric(Cells(9 + sftfit2, 2).Value) = False Then
        Call GetOutFit
        Exit Sub
    End If

    ls = Cells(9 + sftfit2, 2).Value
    p = 0
    fileNum = 0     ' # of iteration
    a0 = 0          ' Check tolerance for amp. ration
    a1 = 0          ' Check tolerance for BE diff.
    
    If Cells(1, 1).Value = "Shirley" Then
        a2 = Cells(2, 2).Value
    Else
        a2 = 0.01       ' threshold for peak ratio and BE difference (%/100)
    End If
    
    If IsEmpty(Cells(17, 101).Value) = True Then    ' for old version
        Cells(17, 100).Value = "Iteration limit"
        Cells(17, 101).Value = 10   ' limit of iteration
    End If
    
    For Each rng In Range(Cells(2, 5), Cells(6, (4 + j))).Cells
        If IsNumeric(rng.Value) = False Then
            strErr = "Error in the non-numeric initial fitting parameters."
            TimeCheck = MsgBox(strErr)
            Call GetOutFit
            Exit Sub
        End If
    Next
    
    For Each rng In Range(Cells(2, 3), Cells(7 + sftfit2, 4)).Cells
        If IsNumeric(rng.Value) = False Then
            rng.Value = vbNullString
        End If
    Next
    
Resolve:
    fileNum = fileNum + 1
    
    Call SolverSetup

    If StrComp(Cells(1, 1).Value, "Polynominal", 1) = 0 Then
        If StrComp(Cells(1, 2).Value, "Shirley", 1) = 0 Then
            If StrComp(Cells(1, 3).Value, "ABG", 1) = 0 Then
                SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(7 + sftfit2 - 2, (4 + j))) ' active Shirley
                ' Error here : No Solver reference in VBE - Tools - References - Solver checked.
                For k = 2 To 10
                    If Cells(k, 2).Font.Bold = "True" Then
                        SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
                    ElseIf k = 6 Then
                        SolverAdd CellRef:=Cells(6, 2), Relation:=1, FormulaText:=1 ' max ratio
                        SolverAdd CellRef:=Cells(6, 2), Relation:=3, FormulaText:=0 ' min
                    End If
                Next
                
                SolverAdd CellRef:=Cells(4, 2), Relation:=1, FormulaText:=1 ' max A
                SolverAdd CellRef:=Cells(4, 2), Relation:=3, FormulaText:=0 ' min
            Else
                SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 5), Cells(7 + sftfit2 - 2, (4 + j))) ' active Shirley
            End If
        ElseIf StrComp(Cells(1, 2).Value, "Tougaard", 1) = 0 Or StrComp(Cells(1, 2).Value, "Conv-Tougaard", 1) = 0 Then
            If StrComp(Cells(1, 3).Value, "ABG", 1) = 0 Then
                SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(7 + sftfit2 - 2, (4 + j))) ' active Tougaard
                SolverAdd CellRef:=Range(Cells(2, 2), Cells(3, 2)), Relation:=1, FormulaText:=5000
                SolverAdd CellRef:=Range(Cells(2, 2), Cells(3, 2)), Relation:=3, FormulaText:=0.001
                SolverAdd CellRef:=Cells(4, 2), Relation:=1, FormulaText:=100
                SolverAdd CellRef:=Cells(5, 2), Relation:=1, FormulaText:=10
                SolverAdd CellRef:=Cells(6, 2), Relation:=1, FormulaText:=Cells(2, 101)
            
                For k = 2 To 10
                    If Cells(k, 2).Font.Bold = "True" Then
                        SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
                    ElseIf k = 7 Then
                        SolverAdd CellRef:=Cells(7, 2), Relation:=1, FormulaText:=1 ' max
                        SolverAdd CellRef:=Cells(7, 2), Relation:=3, FormulaText:=0 ' min
                    End If
                Next
            Else
                SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 5), Cells(7 + sftfit2 - 2, (4 + j))) ' static Tougaard
            End If
        Else
            If StrComp(Cells(1, 2).Value, "ABG", 1) = 0 Or StrComp(Cells(1, 3).Value, "ABG", 1) = 0 Then
                SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(7 + sftfit2 - 2, (4 + j)))  ' active Poly
        
                For k = 2 To 5
                    If Cells(k, 2).Font.Bold = "True" Then
                        SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
                    End If
                Next
            Else
                SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 5), Cells(7 + sftfit2 - 2, (4 + j)))  ' static poly
            End If
        End If
    ElseIf StrComp(Cells(1, 1).Value, "Shirley", 1) = 0 Then
        SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 5), Cells(7 + sftfit2 - 2, (4 + j))) ' static Shirley
    ElseIf StrComp(Cells(1, 1).Value, "Tougaard") = 0 Then
        SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 5), Cells(7 + sftfit2 - 2, (4 + j))) ' static Tougaard
    ElseIf StrComp(Cells(1, 1).Value, "Victoreen", 1) = 0 Then
        SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 5), Cells(7 + sftfit2 - 2, (4 + j))) ' static
    ElseIf StrComp(Cells(1, 1).Value, "Arctan", 1) = 0 Then
        SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(7 + sftfit2 - 2, (4 + j))) ' active
        SolverAdd CellRef:=Cells(4, 2), Relation:=3, FormulaText:=Cells(11 + sftfit2, 2).Value        ' This is a point to control the position of inflection
        SolverAdd CellRef:=Cells(4, 2), Relation:=1, FormulaText:=Cells(12 + sftfit2, 2).Value
        SolverAdd CellRef:=Cells(5, 2), Relation:=3, FormulaText:=1 'step width minimum
        SolverAdd CellRef:=Cells(5, 2), Relation:=1, FormulaText:=(Cells(12 + sftfit2, 2).Value - Cells(11 + sftfit2, 2).Value)
        SolverAdd CellRef:=Cells(3, 2), Relation:=3, FormulaText:=0
        SolverAdd CellRef:=Cells(3, 2), Relation:=1, FormulaText:=(Cells(3, 101).Value - Cells(2, 101).Value)
        SolverAdd CellRef:=Cells(2, 2), Relation:=3, FormulaText:=0
        SolverAdd CellRef:=Cells(6, 2), Relation:=3, FormulaText:=-1
        SolverAdd CellRef:=Cells(6, 2), Relation:=1, FormulaText:=1
        SolverAdd CellRef:=Cells(7, 2), Relation:=3, FormulaText:=0
        SolverAdd CellRef:=Cells(7, 2), Relation:=1, FormulaText:=1
        
        For k = 2 To 7
            If Cells(k, 2).Font.Bold = "True" Then
                SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
            End If
        Next
    Else
        SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 5), Cells(7 + sftfit2 - 2, (4 + j)))  ' static
'        SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(7 + sftfit2 - 2, (4 + j)))  ' active
'        For k = 2 To 11
'            If Cells(k, 2).Font.Bold = "True" Then
'                SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
'            End If
'        Next
    End If
    
    If StrComp(strl(1), "Pe", 1) = 0 Or StrComp(strl(1), "Po", 1) = 0 Then
        SolverAdd CellRef:=Range(Cells(2, 5), Cells(2, (4 + j))), Relation:=3, FormulaText:=Cells(startR, 1).Value
        SolverAdd CellRef:=Range(Cells(2, 5), Cells(2, (4 + j))), Relation:=1, FormulaText:=Cells(endR, 1).Value
    Else
        SolverAdd CellRef:=Range(Cells(2, 5), Cells(2, (4 + j))), Relation:=1, FormulaText:=Cells(startR, 1).Value
        SolverAdd CellRef:=Range(Cells(2, 5), Cells(2, (4 + j))), Relation:=3, FormulaText:=Cells(endR, 1).Value
    End If
    
    SolverAdd CellRef:=Range(Cells(4, 5), Cells(4, (4 + j))), Relation:=1, FormulaText:=Cells(2, 103).Value    ' width1 max
    SolverAdd CellRef:=Range(Cells(4, 5), Cells(4, (4 + j))), Relation:=3, FormulaText:=Cells(3, 103).Value    ' width1 min
    SolverAdd CellRef:=Range(Cells(6, 5), Cells(6, (4 + j))), Relation:=3, FormulaText:=0       ' minimum amplitude (1E-6/dblMin)
    SolverAdd CellRef:=Range(Cells(3, 5), Cells(3, (4 + j))), Relation:=2, FormulaText:=0
        
    For n = 1 To j
        If Cells(7, (4 + n)).Value = 0 Or Cells(7, (4 + n)).Value = "Gauss" Then ' G
            SolverAdd CellRef:=Cells(7, (4 + n)), Relation:=2, FormulaText:=0
            SolverAdd CellRef:=Range(Cells(8, (4 + n)), Cells(10, (4 + n))), Relation:=2, FormulaText:=0
            SolverAdd CellRef:=Cells(5, (4 + n)), Relation:=2, FormulaText:=0        ' width2
        ElseIf Cells(7, (4 + n)).Value = 1 Or Cells(7, (4 + n)).Value = "Lorentz" Then
            SolverAdd CellRef:=Cells(7, (4 + n)), Relation:=2, FormulaText:=1
            If Cells(7, (4 + n)).Font.Italic = "True" Then  'Doniach-Sunjic (DS)
                SolverAdd CellRef:=Cells(8, (4 + n)), Relation:=1, FormulaText:=1
                SolverAdd CellRef:=Cells(8, (4 + n)), Relation:=3, FormulaText:=0
                SolverAdd CellRef:=Cells(5, (4 + n)), Relation:=2, FormulaText:=0        ' width2
                SolverAdd CellRef:=Cells(10, (4 + n)), Relation:=2, FormulaText:=0
            Else    ' L
                SolverAdd CellRef:=Range(Cells(8, (4 + n)), Cells(10, (4 + n))), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(5, (4 + n)), Relation:=2, FormulaText:=0        ' width2
            End If
        Else
            If mid$(Cells(11, (4 + n)).Value, 1, 1) = "E" Then
                SolverAdd CellRef:=Range(Cells(9, (4 + n)), Cells(10, (4 + n))), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(5, (4 + n)), Relation:=2, FormulaText:=0        ' width2
            ElseIf mid$(Cells(11, (4 + n)).Value, 1, 1) = "T" Then       ' MultiPak Asymmetric GL with exp tail
                SolverAdd CellRef:=Cells(10, (4 + n)), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(5, (4 + n)), Relation:=2, FormulaText:=0        ' width2
                SolverAdd CellRef:=Cells(8, (4 + n)), Relation:=1, FormulaText:=3         ' max a : Tail scale
                SolverAdd CellRef:=Cells(8, (4 + n)), Relation:=3, FormulaText:=0         ' min a : Tail scale
                SolverAdd CellRef:=Cells(9, (4 + n)), Relation:=1, FormulaText:=Abs(Cells(6, 101).Value - Cells(7, 101).Value)          ' max b : Tail length
                SolverAdd CellRef:=Cells(9, (4 + n)), Relation:=3, FormulaText:=1         ' min b : Tail length
            ElseIf mid$(Cells(11, (4 + n)).Value, 1, 1) = "D" Then
                    SolverAdd CellRef:=Cells(10, (4 + n)), Relation:=2, FormulaText:=0
                    SolverAdd CellRef:=Cells(8, (4 + n)), Relation:=1, FormulaText:=1         ' max a
                    SolverAdd CellRef:=Cells(8, (4 + n)), Relation:=3, FormulaText:=0         ' min a
                    SolverAdd CellRef:=Cells(9, (4 + n)), Relation:=1, FormulaText:=1         ' max b
                    SolverAdd CellRef:=Cells(9, (4 + n)), Relation:=3, FormulaText:=0         ' min b
            ElseIf mid$(Cells(11, (4 + n)).Value, 1, 1) = "U" Then
                SolverAdd CellRef:=Cells(10, (4 + n)), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(8, (4 + n)), Relation:=1, FormulaText:=1         ' max a
                SolverAdd CellRef:=Cells(8, (4 + n)), Relation:=3, FormulaText:=0         ' min a
                SolverAdd CellRef:=Cells(9, (4 + n)), Relation:=1, FormulaText:=1         ' max b
                SolverAdd CellRef:=Cells(9, (4 + n)), Relation:=3, FormulaText:=0         ' min b
            Else
                SolverAdd CellRef:=Range(Cells(8, (4 + n)), Cells(10, (4 + n))), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(5, (4 + n)), Relation:=1, FormulaText:=Cells(4, 103).Value        ' width2 max
                SolverAdd CellRef:=Cells(5, (4 + n)), Relation:=3, FormulaText:=Cells(5, 103).Value         ' width2 min
            End If

            SolverAdd CellRef:=Cells(7, (4 + n)), Relation:=1, FormulaText:=Cells(6, 103).Value         ' max shape
            SolverAdd CellRef:=Cells(7, (4 + n)), Relation:=3, FormulaText:=Cells(7, 103).Value         ' min shape
        End If
    Next

    For n = 5 To (4 + j)
        For k = 1 To 9
            If Cells((k + 1), n).Font.Bold = "True" Then
                SolverAdd CellRef:=Cells((k + 1), n), Relation:=2, FormulaText:=Cells((k + 1), n)
            End If
        Next
        
        If Cells(2, n).Font.Italic = "True" Then
            SolverAdd CellRef:=Cells(2, n), Relation:=1, FormulaText:=Cells(2, n) + Cells(8, 103).Value         ' max BE
            SolverAdd CellRef:=Cells(2, n), Relation:=3, FormulaText:=Cells(2, n) - Cells(8, 103).Value         ' min BE
        End If
    Next
    
    strErr = "Amp. ratio format error: (i; j; k) and i,j,k > 0"
    iRow = 1
    
    For n = 5 To (4 + j)
        If Not Cells(14 + sftfit2, n) = vbNullString Then
            If iRow = 1 And mid$(Cells(14 + sftfit2, n), 1, 1) = "(" And mid$(Cells(14 + sftfit2, n), Len(Cells(14 + sftfit2, n)), 1) = ";" Then
                If IsNumeric(mid$(Cells(14 + sftfit2, n), 2, Len(Cells(14 + sftfit2, n)) - 2)) = True Then
                    ReDim ratio(1)
                    ratio(1) = mid$(Cells(14 + sftfit2, n), 2, Len(Cells(14 + sftfit2, n)) - 2)
                    ratio1 = ratio(1)
                    iRow = iRow + 1
                Else
                    TimeCheck = MsgBox(strErr, vbCritical)
                    Call GetOutFit
                    Exit Sub
                End If
            ElseIf iRow > 1 And mid$(Cells(14 + sftfit2, n), 1, 1) = "(" Then
                TimeCheck = MsgBox(strErr, vbCritical)
                Call GetOutFit
                Exit Sub
            ElseIf iRow > 1 And mid$(Cells(14 + sftfit2, n), Len(Cells(14 + sftfit2, n)), 1) = ";" Then
                If IsNumeric(mid$(Cells(14 + sftfit2, n), 1, Len(Cells(14 + sftfit2, n)) - 1)) = True Then
                    ReDim Preserve ratio(iRow)
                    ratio(iRow) = mid$(Cells(14 + sftfit2, n), 1, InStr(1, Cells(14 + sftfit2, n), ";") - 1)
                    iRow = iRow + 1
                Else
                    TimeCheck = MsgBox(strErr, vbCritical)
                    Call GetOutFit
                    Exit Sub
                End If
            ElseIf iRow > 1 And mid$(Cells(14 + sftfit2, n), Len(Cells(14 + sftfit2, n)), 1) = ")" Then
                If IsNumeric(mid$(Cells(14 + sftfit2, n), 1, Len(Cells(14 + sftfit2, n)) - 1)) = True Then
                    ReDim Preserve ratio(iRow)
                    ratio(iRow) = mid$(Cells(14 + sftfit2, n), 1, Len(Cells(14 + sftfit2, n)) - 1)
                Else
                    TimeCheck = MsgBox(strErr, vbCritical)
                    Call GetOutFit
                    Exit Sub
                End If
                For iCol = iRow - 1 To 0 Step -1        ' max amplitude ratio to be reference, not in the first bracket!
                    If IsNumeric(ratio(iRow - iCol)) = True Then
                        If ratio(iRow - iCol) >= ratio1 Then
                            ratio1 = ratio(iRow - iCol)
                            k = iRow - iCol
                        Else
                            k = 1           ' Added in ver. 7.19
                        End If
                    End If
                Next
                For iCol = iRow - 1 To 0 Step -1
                    If IsNumeric(ratio(iRow - iCol)) = True And ratio(iRow - iCol) > 0 Then
                        If iRow - iCol = k Then
                           SolverAdd CellRef:=Cells(6, n - iRow + k), Relation:=1, FormulaText:=Cells(3, 101).Value - Cells(2, 101).Value
                           Cells(15 + sftfit2, n - iCol + 110).Value = ratio(iRow - iCol) / ratio(k)
                        Else
                           SolverAdd CellRef:=Cells(6, n - iCol), Relation:=2, FormulaText:=Cells(6, n - iRow + k) * ratio(iRow - iCol) / ratio(k)
                           Cells(15 + sftfit2, n - iCol + 110).Value = ratio(iRow - iCol) / ratio(k)
                        End If
                    ElseIf ratio(iRow - iCol) = "NaN" Then
                        SolverAdd CellRef:=Cells(6, n - iCol), Relation:=1, FormulaText:=Cells(3, 101).Value - Cells(2, 101).Value
                    Else
                        Range(Cells(15 + sftfit2, 4 + 110), Cells(16, 4 + j + 110)).ClearContents
                        TimeCheck = MsgBox(strErr, vbCritical)
                        Call GetOutFit
                        Exit Sub
                    End If
                Next
                iRow = 1
            Else
                TimeCheck = MsgBox(strErr, vbCritical)
                Call GetOutFit
                Exit Sub
            End If
        ElseIf iRow > 1 Then
            ReDim Preserve ratio(iRow)
            ratio(iRow) = "NaN"
            iRow = iRow + 1
            SolverAdd CellRef:=Cells(6, n), Relation:=1, FormulaText:=Cells(3, 101).Value - Cells(2, 101).Value
        Else
            SolverAdd CellRef:=Cells(6, n), Relation:=1, FormulaText:=Cells(3, 101).Value - Cells(2, 101).Value
        End If
    Next
    
    strErr = "BE diff format error: [ i; nj; k] and i,j,k > 0" & vbCrLf & " *n* represents negative sign."
    iRow = 0
    
    For n = 5 To (4 + j)
        If Not Cells(15 + sftfit2, n) = vbNullString Then
            If iRow = 0 And StrComp(Cells(15 + sftfit2, n), "[", 1) = 0 Then
                ReDim bediff(1)
                iRow = iRow + 1
            ElseIf iRow > 0 And StrComp(Cells(15 + sftfit2, n), "[", 1) = 0 Then
                TimeCheck = MsgBox(strErr, vbCritical)
                Call GetOutFit
                Exit Sub
            ElseIf iRow > 0 And mid$(Cells(15 + sftfit2, n), Len(Cells(15 + sftfit2, n)), 1) = ";" Then
                If IsNumeric(mid$(Cells(15 + sftfit2, n), 1, Len(Cells(15 + sftfit2, n)) - 1)) = True Then
                    ReDim Preserve bediff(iRow)
                    bediff(iRow) = mid$(Cells(15 + sftfit2, n), 1, Len(Cells(15 + sftfit2, n)) - 1)
                    iRow = iRow + 1
                ElseIf mid$(Cells(15 + sftfit2, n), 1, 1) = "n" Then
                    If IsNumeric(mid$(Cells(15 + sftfit2, n), 2, Len(Cells(15 + sftfit2, n)) - 2)) = True Then
                        ReDim Preserve bediff(iRow)
                        bediff(iRow) = -1 * mid$(Cells(15 + sftfit2, n), 2, Len(Cells(15 + sftfit2, n)) - 2)
                        iRow = iRow + 1
                    Else
                        TimeCheck = MsgBox(strErr, vbCritical)
                        Call GetOutFit
                        Exit Sub
                    End If
                Else
                    TimeCheck = MsgBox(strErr, vbCritical)
                    Call GetOutFit
                    Exit Sub
                End If
            ElseIf iRow > 0 And mid$(Cells(15 + sftfit2, n), Len(Cells(15 + sftfit2, n)), 1) = "]" Then
                If IsNumeric(mid$(Cells(15 + sftfit2, n), 1, Len(Cells(15 + sftfit2, n)) - 1)) = True Then
                    ReDim Preserve bediff(iRow)
                    bediff(iRow) = mid$(Cells(15 + sftfit2, n), 1, Len(Cells(15 + sftfit2, n)) - 1)
                    iRow = iRow
                ElseIf mid$(Cells(15 + sftfit2, n), 1, 1) = "n" Then
                    If IsNumeric(mid$(Cells(15 + sftfit2, n), 2, Len(Cells(15 + sftfit2, n)) - 2)) = True Then
                        ReDim Preserve bediff(iRow)
                        bediff(iRow) = -1 * mid$(Cells(15 + sftfit2, n), 2, Len(Cells(15 + sftfit2, n)) - 2)
                        iRow = iRow
                    Else
                        TimeCheck = MsgBox(strErr, vbCritical)
                        Call GetOutFit
                        Exit Sub
                    End If
                Else
                    TimeCheck = MsgBox(strErr, vbCritical)
                    Call GetOutFit
                    Exit Sub
                End If
                For iCol = iRow - 1 To 0 Step -1
                    If IsNumeric(bediff(iRow - iCol)) = True Then
                        SolverAdd CellRef:=Cells(2, n - iCol), Relation:=2, FormulaText:=Cells(2, n - iRow) + bediff(iRow - iCol)
                    ElseIf bediff(iRow - iCol) = "NaN" Then
                    Else
                        TimeCheck = MsgBox(strErr, vbCritical)
                        Call GetOutFit
                        Exit Sub
                    End If
                Next
                iRow = 0
            Else
                TimeCheck = MsgBox(strErr, vbCritical)
                Call GetOutFit
                Exit Sub
            End If
        ElseIf iRow > 0 Then
            ReDim Preserve bediff(iRow)
            bediff(iRow) = "NaN" 'vbNullString
            iRow = iRow + 1
        End If
    Next

    Results = SolverSolve(UserFinish:=True, ShowRef:="ShowTrial")  ' Results of fitting by Solver
    
    SolverFinish KeepFinal:=1
    
    iRow = 1
    
    For n = 5 To (4 + j)
        If Not Cells(14 + sftfit2, n) = vbNullString Then
            If iRow = 1 And mid$(Cells(14 + sftfit2, n), 1, 1) = "(" Then
                iRow = iRow + 1
            ElseIf iRow > 1 And mid$(Cells(14 + sftfit2, n), Len(Cells(14 + sftfit2, n)), 1) = ";" Then
                iRow = iRow + 1
            ElseIf iRow > 1 And mid$(Cells(14 + sftfit2, n), Len(Cells(14 + sftfit2, n)), 1) = ")" Then
                For iCol = iRow - 1 To 0 Step -1
                    If IsNumeric(Cells(15 + sftfit2, n - iCol + 110)) = True And Cells(6, n - iCol) > 0 And Cells(6, n - iRow + 1) > 0 Then
                        Cells(16 + sftfit2, n - iCol + 110).Value = Cells(6, n - iCol) / Cells(6, n - iRow + 1)
                        If Cells(15 + sftfit2, n - iCol + 110) > 0 Then
                            If Abs((Cells(16 + sftfit2, n - iCol + 110) - Cells(15 + sftfit2, n - iCol + 110)) / Cells(15 + sftfit2, n - iCol + 110)) > a2 And fileNum < Cells(17, 101).Value Then
                                GoTo Resolve
                            ElseIf fileNum >= Cells(17, 101).Value And Abs((Cells(16 + sftfit2, n - iCol + 110) - Cells(15 + sftfit2, n - iCol + 110)) / Cells(15 + sftfit2, n - iCol + 110)) > a2 Then
                                a0 = a0 + Abs((Cells(16 + sftfit2, n - iCol + 110) - Cells(15 + sftfit2, n - iCol + 110)) / Cells(15 + sftfit2, n - iCol + 110))
                                a1 = a2
                                GoTo ExitIter
                            Else
                                a0 = a0 + Abs((Cells(16 + sftfit2, n - iCol + 110) - Cells(15 + sftfit2, n - iCol + 110)) / Cells(15 + sftfit2, n - iCol + 110))
                            End If
                        End If
                    End If
                Next
                iRow = 1
            End If
        ElseIf iRow > 1 Then
            iRow = iRow + 1
        End If
    Next
    
    iRow = 0
    
    For n = 5 To (4 + j)
        If Not Cells(15 + sftfit2, n) = vbNullString Then
            If iRow = 0 And StrComp(Cells(15 + sftfit2, n), "[", 1) = 0 Then
                iRow = iRow + 1
            ElseIf iRow > 0 And mid$(Cells(15 + sftfit2, n), Len(Cells(15 + sftfit2, n)), 1) = ";" Then
                iRow = iRow + 1
            ElseIf iRow > 0 And mid$(Cells(15 + sftfit2, n), Len(Cells(15 + sftfit2, n)), 1) = "]" Then
                For iCol = iRow - 1 To 0 Step -1
                    If IsNumeric(bediff(iRow - iCol)) = True Then
                        If Abs((bediff(iRow - iCol) - (Cells(2, n - iCol) - Cells(2, n - iRow))) / bediff(iRow - iCol)) > a2 And fileNum < Cells(17, 101).Value Then
                            GoTo Resolve
                        ElseIf fileNum >= Cells(17, 101).Value Then
                            a1 = a1 + Abs((bediff(iRow - iCol) - (Cells(2, n - iCol) - Cells(2, n - iRow))) / bediff(iRow - iCol))
                            GoTo ExitIter
                        Else
                            a1 = a1 + Abs((bediff(iRow - iCol) - (Cells(2, n - iCol) - Cells(2, n - iRow))) / bediff(iRow - iCol))
                        End If
                    End If
                Next
                iRow = 0
            End If
        ElseIf iRow > 0 Then
            iRow = iRow + 1
        End If
    Next
    
ExitIter:

    Range(Cells(15 + sftfit2, 4 + 110), Cells(16 + sftfit2, 4 + j + 110)).ClearContents
    
    If Cells(7, (4 + n)).Value < 0.999 And Cells(7, (4 + n)).Value > 0.001 Then
        For n = 1 To j
            If Cells(7, (4 + n)).Font.Italic = "True" Then
                For k = 1 To numData
                    If Cells((startR - 1 + k), 1).Value < Cells(2, (4 + n)).Value Then
                        Cells((startR - 1 + k), (4 + n)).FormulaR1C1 = "=R6C * ((R7C) *((((R5C)/2)^2)/((RC[" & (-3 - n) & "]-R2C)^2 + ((R5C)/2)^2)) + (1- R7C) *(EXP(-(1/2)*((RC[" & (-3 - n) & "]-R2C)/(R5C/2.35))^2)))"
                    Else
                        Cells((startR - 1 + k), (4 + n)).FormulaR1C1 = "=R6C * ((R7C) *((((R4C)/2)^2)/((RC[" & (-3 - n) & "]-R2C)^2 + ((R4C)/2)^2)) + (1- R7C) *(EXP(-(1/2)*((RC[" & (-3 - n) & "]-R2C)/(R4C/2.35))^2)))"
                    End If
                Next
                p = p + 1
            End If
        Next
        If p > 0 Then
            If Abs(Cells(9 + sftfit2, 2).Value - ls) > 1 Then GoTo AsymIteration     ' Tolerance = 1
        End If
    End If
    
    For n = 1 To j
        If Cells(7, (4 + n)).Value > 0 And Cells(7, (4 + n)).Value < 1 Then
            If Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleSingle Then  ' exponential asymmetric blend based Voigt
                Cells(13 + sftfit2, (4 + n)).Value = vbNullString ' R5C to be exp decay parameter.
            ElseIf Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleDouble Then   ' Ulrik Gelius mode asymmetric blend based Voigt (GL sum)
                'Cells(13, (4 + n)).Value = 0    ' used as a b parameter manually adjusted
                Cells(13 + sftfit2, (4 + n)).Value = vbNullString
            ElseIf Cells(11, (4 + n)).Value = "GL" Then
                Cells(13 + sftfit2, (4 + n)).Value = vbNullString
            Else
                Cells(13 + sftfit2, (4 + n)).Value = Cells(5, (4 + n)).Value / Cells(4, (4 + n)).Value
            End If
        ElseIf Cells(7, (4 + n)).Value = 0 Then
            Cells(7, (4 + n)).Value = "Gauss"
            Cells(5, (4 + n)).Value = vbNullString
            Cells(13 + sftfit2, (4 + n)).Value = vbNullString
        ElseIf Cells(7, (4 + n)).Value = 1 Then
            Cells(7, (4 + n)).Value = "Lorentz"
            Cells(13 + sftfit2, (4 + n)).Value = vbNullString
            If Cells(7, (4 + n)).Font.Italic = "True" Then  'Doniach-Sunjic (DS)
                ' FWHM2 to be alpha; asymmetric parameter
            Else
                Cells(5, (4 + n)).Value = vbNullString
            End If
        End If
        
        Cells(3, (4 + n)).FormulaR1C1 = "=(R12C101 - R13C101 - R14C101 - R2C)" ' KE
    Next
    
    Cells(8, 101).Value = Cells(8, 101).Value + 1     ' means already fit once
    
    If startR > 21 + sftfit Then
        Range(Cells(23 + sftfit + numData, 5), Cells(2 + numData + startR - 1, 55)).ClearContents
    End If
    
    If endR < numData + 20 + sftfit Then
        Range(Cells(2 + numData + endR + 1, 5), Cells(2 * numData + 22 + sftfit, 55)).ClearContents
    End If
    
    Call descriptInitialFit
    
    If StrComp(strl(1), "Pe", 1) = 0 Then
        Cells(2, 4).Value = "PE"
        Range(Cells(3, 5), Cells(3, 55)).ClearContents
    ElseIf StrComp(strl(1), "Po", 1) = 0 Then
        Cells(2, 4).Value = "Po"
        Range(Cells(3, 5), Cells(3, 55)).ClearContents
    End If
    
    Call GetOutFit
End Sub

Sub FitEF()
    Dim rng As Range, dataFit As Range
    
    If startR >= 21 + sftfit Then
        If IsEmpty(Cells(startR - 1, 3)) = False Then
            Range(Cells(21 + sftfit, 3), Cells(startR - 1, 5)).ClearContents
            Cells(8, 101).Value = -1
        ElseIf IsEmpty(Cells(startR, 3)) = True Then
            Cells(8, 101).Value = -1
        End If
    End If
    
    If endR <= numData + 20 + sftfit Then
        If IsEmpty(Cells(endR + 1, 3)) = False Then
            Range(Cells(endR + 1, 3), Cells(numData + 20 + sftfit, 5)).ClearContents
            Cells(8, 101).Value = -1
        ElseIf IsEmpty(Cells(endR, 3)) = True Then
            Cells(8, 101).Value = -1
        End If
    End If
    
    If Cells(8, 101).Value > 0 Then
        GoTo SkipInitialEF2
    ElseIf Cells(8, 101).Value < 0 Then
        fcmp = Range(Cells(2, 5), Cells(8, 5))
    End If
    
    Range(Cells(1, 3), Cells(15 + sftfit2, 55)).ClearContents
    Range(Cells(20 + sftfit, 3), Cells((2 * numData + 22 + sftfit), 55)).ClearContents
    Range(Cells(1, 3), Cells(15 + sftfit2, 55)).Interior.ColorIndex = xlNone
    
    Call descriptEFfit1(fcmp)
    Call descriptGConv
    
SkipInitialEF:
    
    Cells(startR, 3).FormulaR1C1 = "= R8C2 * (((R4C2 + R5C2 * (RC[-2] - R2C5))  + (R6C2 * (RC[-2] - R2C5)^2) + (R7C2 * (RC[-2] - R2C5)^3)) +  ((R2C2 + (R3C2 * (RC[-2] - R2C5))) / (1 + EXP(-(RC[-2] - R2C5) * 11604.86 / R4C5))))"
    
    Range(Cells(startR, 3), Cells(endR, 3)).FillDown
    Cells(startR, 4).FormulaR1C1 = "=((RC[-2] - RC[-1])^2)"
    Range(Cells(startR, 4), Cells(endR, 4)).FillDown
    Cells(5 + sftfit2, 2).FormulaR1C1 = "=SUM(R" & startR & "C4:R" & endR & "C4)"
    Cells(startR, 5).FormulaR1C1 = "=(RC[-3] - RC[-2])"
    Range(Cells(startR, 5), Cells(endR, 5)).FillDown
    
    For Each rng In Range(Cells(2, 3), Cells(6, 4)).Cells
        If IsNumeric(rng.Value) = False Then
            rng.Value = 0
        End If
    Next
    
    Call SolverSetup
    SolverOk SetCell:=Cells(5 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(8, 5))
    SolverAdd CellRef:=Cells(2, 5), Relation:=3, FormulaText:=Cells(8 + sftfit2, 2)   ' min
    SolverAdd CellRef:=Cells(2, 5), Relation:=1, FormulaText:=Cells(9 + sftfit2, 2)   ' max
    SolverAdd CellRef:=Cells(3, 2), Relation:=1, FormulaText:=Abs(Cells(2, 2))
    SolverAdd CellRef:=Cells(5, 2), Relation:=1, FormulaText:=Abs(Cells(4, 2))
    SolverAdd CellRef:=Cells(3, 2), Relation:=3, FormulaText:=-1 * Abs(Cells(2, 2))
    SolverAdd CellRef:=Cells(5, 2), Relation:=3, FormulaText:=-1 * Abs(Cells(4, 2))
    SolverAdd CellRef:=Cells(6, 2), Relation:=1, FormulaText:=Abs(Cells(5, 2)) / 10
    SolverAdd CellRef:=Cells(7, 2), Relation:=1, FormulaText:=Abs(Cells(6, 2)) / 10
    SolverAdd CellRef:=Cells(6, 2), Relation:=3, FormulaText:=-1 * Abs(Cells(5, 2)) / 10
    SolverAdd CellRef:=Cells(7, 2), Relation:=3, FormulaText:=-1 * Abs(Cells(6, 2)) / 10
    SolverAdd CellRef:=Cells(4, 5), Relation:=1, FormulaText:=10000
    SolverAdd CellRef:=Cells(4, 5), Relation:=3, FormulaText:=1
    SolverAdd CellRef:=Cells(2, 2), Relation:=3, FormulaText:=Cells(4, 2)
    SolverAdd CellRef:=Cells(8, 2), Relation:=3, FormulaText:=0.0001
    SolverAdd CellRef:=Cells(6, 5), Relation:=2, FormulaText:=Cells(6, 5)

    For k = 2 To 8
        If Cells(k, 2).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
        End If
    Next

    For k = 2 To 6
        If Cells(k, 5).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 5), Relation:=2, FormulaText:=Cells(k, 5)
        End If
    Next

    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
    
SkipInitialEF2:

    p = startR + Cells(10, 101).Value
    q = endR - Cells(10, 101).Value
    
    Cells(p, 6).FormulaR1C1 = "= RC100*(R8C5)"
    Range(Cells(p, 6), Cells(q, 6)).FillDown
    Cells(p, 7).FormulaR1C1 = "=((RC2 - RC[-1])^2)"
    Range(Cells(p, 7), Cells(q, 7)).FillDown
    Cells(6 + sftfit2, 2).FormulaR1C1 = "=SUM(R" & p & "C7:R" & q & "C7)"
    Cells(p, 8).FormulaR1C1 = "=(RC2 - RC[-2])"
    Range(Cells(p, 8), Cells(q, 8)).FillDown
    Range(Cells(startR, 6), Cells(p - 1, 8)).ClearContents
    Range(Cells(q + 1, 6), Cells(endR, 8)).ClearContents

    If Cells(6, 5).Value <= 0.01 Then Cells(6, 5).Value = 1
    
    Call SolverSetup
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 5), Cells(8, 5))
    SolverAdd CellRef:=Cells(6, 5), Relation:=3, FormulaText:=Cells(7, 103)   ' min      Gaussian width to be convoluted
    SolverAdd CellRef:=Cells(6, 5), Relation:=1, FormulaText:=Cells(6, 103)   ' max
    SolverAdd CellRef:=Cells(2, 5), Relation:=3, FormulaText:=Cells(8 + sftfit2, 2)   ' min
    SolverAdd CellRef:=Cells(2, 5), Relation:=1, FormulaText:=Cells(9 + sftfit2, 2)   ' max
    SolverAdd CellRef:=Cells(4, 5), Relation:=1, FormulaText:=10000
    SolverAdd CellRef:=Cells(4, 5), Relation:=3, FormulaText:=1
    SolverAdd CellRef:=Cells(8, 5), Relation:=3, FormulaText:=0

    For k = 2 To 8
        If Cells(k, 5).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 5), Relation:=2, FormulaText:=Cells(k, 5)
        End If
    Next
    
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
    
    Cells(8, 101).Value = Cells(8, 101).Value + 1     ' means already fit once
    
    Call descriptEFfit2
    
    If Cells(8, 101).Value > 1 Then Exit Sub
    
    Set rng = dataData
    Set dataFit = dataIntData
    
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.SeriesCollection.NewSeries  '7.45
    With ActiveChart.SeriesCollection(2)
        .ChartType = xlXYScatterLinesNoMarkers
        .XValues = rng
        .Values = rng.Offset(, 2)
        .Border.ColorIndex = 33
        .Format.Line.Weight = 3
        '.Name = "fit EF (FD)"
        .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C3"
    End With
    
    ActiveChart.SeriesCollection.NewSeries  '7.45
    With ActiveChart.SeriesCollection(3)
        .ChartType = xlXYScatterLinesNoMarkers
        .XValues = dataFit
        .Values = dataFit.Offset(, 5)
        .Border.ColorIndex = 41
        .Format.Line.Weight = 3
        .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C6"
    End With
    
    ActiveSheet.ChartObjects(2).Activate
    ActiveSheet.ChartObjects(2).Chart.Axes(xlValue, xlPrimary).AxisTitle.Text = "Residual"
    With ActiveChart.SeriesCollection(1)
        .ChartType = xlXYScatterLinesNoMarkers
        .XValues = rng
        .Values = rng.Offset(, 4)
        .Border.ColorIndex = 44
        .Format.Line.Weight = 3
        .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C5"
    End With
    
    ActiveChart.SeriesCollection.NewSeries  '7.45
    With ActiveChart.SeriesCollection(2)
        .ChartType = xlXYScatterLinesNoMarkers
        .XValues = dataFit
        .Values = dataFit.Offset(, 7)
        .Border.ColorIndex = 43
        .Format.Line.Weight = 3
        .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C8"
    End With
End Sub

Sub GetOutFit()
    If Not Cells(1, 1).Value = "EF" And Cells(8, 101).Value > 0 Then
        Call descriptInitialFit
    ElseIf Cells(1, 1).Value = "EF" Then
        strBG1 = "e"
    End If
    
    If StrComp(strBG1, "p", 1) = 0 Then
        If StrComp(strBG2, "s", 1) = 0 Then
            Cells(5, 2).Value = fileNum
            Cells(5, 1).Value = "Iteration"
            Cells(5, 2).Font.Bold = "False"
        ElseIf StrComp(strBG2, "t", 1) = 0 Then
        Else
            Range(Cells(6, 1), Cells(7 + sftfit2 - 2, 2)).ClearContents
            Range(Cells(6, 1), Cells(7 + sftfit2 - 2, 2)).Interior.ColorIndex = xlNone
            Cells(5, 1).Value = "a3"
        End If
    ElseIf StrComp(strBG1, "a", 1) = 0 Or StrComp(strBG1, "r", 1) = 0 Then
        Range(Cells(8, 1), Cells(7 + sftfit2 - 2, 2)).ClearContents
        Range(Cells(8, 1), Cells(7 + sftfit2 - 2, 2)).Interior.ColorIndex = xlNone
        Cells(6, 1).Value = "Slope"
        Cells(7, 1).Value = "ratio L:A"
    ElseIf StrComp(strBG1, "t", 1) = 0 Then
        'Cells(5, 2).Value = fileNum
        Cells(5, 1).Value = "Norm"
        Cells(5, 2).Font.Bold = "False"
        Range(Cells(7, 1), Cells(7 + sftfit2 - 2, 2)).ClearContents
        Range(Cells(7, 1), Cells(7 + sftfit2 - 2, 2)).Interior.ColorIndex = xlNone
    ElseIf StrComp(strBG1, "v", 1) = 0 Then
        If Cells(8, 2).Value = vbNullString Then
            Cells(8, 1).Value = "No edge"
        ElseIf Cells(8, 2).Value < Cells(12 + sftfit2, 2).Value And Cells(8, 2).Value > Cells(11 + sftfit2, 2).Value Then
            Cells(8, 1).Value = "Pre-edge"
        Else
            Cells(8, 1).Value = "Both ends"
        End If
        Range(Cells(10, 1), Cells(7 + sftfit2 - 2, 2)).ClearContents
        Range(Cells(10, 1), Cells(7 + sftfit2 - 2, 2)).Interior.ColorIndex = xlNone
    ElseIf StrComp(strBG1, "e", 1) = 0 Then
        Cells(8, 1).Value = "Norm (FD)"
        Cells(6, 1).Value = "Poly2nd"
        Cells(7, 1).Value = "Poly3rd"
        Cells(5, 1).Value = "Slope BG"
        Range(Cells(9, 4), Cells(19 + sftfit2, 5)).ClearContents
        Range(Cells(9, 4), Cells(19 + sftfit2, 5)).Interior.ColorIndex = xlNone
        Range(Cells(9, 1), Cells(9, 2)).ClearContents
    Else
        Cells(5, 2).Value = fileNum
        Cells(5, 1).Value = "Iteration"
        Cells(5, 2).Font.Bold = "False"
        Range(Cells(6, 1), Cells(7 + sftfit2 - 2, 2)).ClearContents
        Range(Cells(6, 1), Cells(7 + sftfit2 - 2, 2)).Interior.ColorIndex = xlNone
    End If
    
    For n = 1 To j
        If mid$(Cells(11, (4 + n)).Value, 1, 1) = "T" Then
            Cells(10, (4 + n)) = vbNullString
            Cells(5, (4 + n)) = vbNullString
        ElseIf Cells(7, (4 + n)).Value = 0 Or Cells(7, (4 + n)).Value = "Gauss" Or Cells(11, (4 + n)).Value = "GL" Then  ' G
            Range(Cells(8, (4 + n)), Cells(10, (4 + n))) = vbNullString
            Cells(5, (4 + n)) = vbNullString
        ElseIf Cells(7, (4 + n)).Value = 1 Or Cells(7, (4 + n)).Value = "Lorentz" Then
            Range(Cells(8, (4 + n)), Cells(10, (4 + n))) = vbNullString
            Cells(5, (4 + n)) = vbNullString
        Else
            If Cells(1, 1).Value = "EF" Then
                Cells(10, (4 + n)) = vbNullString
            Else
                Range(Cells(8, (4 + n)), Cells(10, (4 + n))) = vbNullString
            End If
        End If
    Next
    
    Cells(7 + sftfit2, 1).Value = "Peak fit"
    Cells(7 + sftfit2, 2).Value = vbNullString
    Range(Cells(2, 3), Cells(7 + sftfit2, 3)).ClearContents
    Range(Cells(2, 3), Cells(7 + sftfit2, 3)).Interior.ColorIndex = xlNone
    Cells(1, 1).Select
    
    Application.Calculation = xlCalculationAutomatic
    Call GetOut
End Sub

Sub EngBL()
    Dim C1 As Variant, C2 As Variant, C3 As Variant, C4 As Variant, imax As Integer, SourceRangeColor1 As Long, strTest As String
    
    If ExistSheet(strSheetGraphName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetGraphName).Delete
        Application.DisplayAlerts = True
    End If
    
    wf = 26.5
    k = 0
    
    If StrComp(testMacro, "debug", 1) = 0 Then
        
    Else
        wf = Application.InputBox(Title:="Calc. 1st har. PE", Prompt:="Input the U60 gap: mm", Default:=wf, Type:=1)
        If wf <= 0 Or Len(wf) = 0 Then
            k = 1
            char = Application.InputBox(Title:="Calc. U60 gap", Prompt:="Input the 1st har. photon energy: eV", Default:=char, Type:=1)
            If char < 0 Or char > 300 Then char = 40
        ElseIf wf < 25 Or wf > 200 Then
            wf = 26.5
        End If
    End If
    
    If StrComp(strMode, "GE/eV", 1) = 0 Then
        C1 = dataKeData                                      ' PE
        C2 = dataKeData.Offset(, 1)                          ' Ip
        If StrComp(Cells(1, 3).Value, "Ie", 1) = 0 Then
            C3 = dataKeData.Offset(, 2)                      ' Ie
        Else
            C3 = dataKeData.Offset(, para + 30)              ' empty Ip
        End If
        C4 = C2
        startEb = Cells(2, 1).Value
        endEb = Cells(numData + 1, 1).Value
        stepEk = Cells(3, 1).Value - Cells(2, 1).Value
        g = 0
        maxXPSFactor = 1
    Else
        startEb = Cells(12, 1).Value
        endEb = Cells(12, 1).End(xlDown).Value
        stepEk = Abs(Cells(13, 1).Value - Cells(12, 1).Value)
        numData = ((endEb - startEb) / stepEk) + 1
        g = mid$(Cells(5, 2).Value, 1, 4)
        
        C1 = Range(Cells(12, 1), Cells(12, 1).Offset(numData - 1, 0))    ' PE
        C2 = Range(Cells(12, 2), Cells(12, 2).Offset(numData - 1, 0))    ' Ip
        C3 = Range(Cells(12, 3), Cells(12, 3).Offset(numData - 1, 0))    ' Ie
        C4 = C2
        maxXPSFactor = 1000000000000#
    End If
    
    Worksheets.Add().Name = strSheetGraphName
    Set sheetGraph = Worksheets(strSheetGraphName)
    sheetGraph.Activate
    
    If k = 0 Then
        Cells(3, 2).Value = wf
        Cells(41, para + 14).FormulaR1C1 = "= " & a0 & " + " & a1 & " * Exp(" & a2 & " * R3C2)"    ' B (T)
        Cells(42, para + 14).FormulaR1C1 = "= 0.934 * " & lambda & " * (R41C" & (para + 14) & ")"                          ' K
        Cells(4, 2).FormulaR1C1 = "=950 * ((" & gamma & ") ^ 2) / (((((R42C" & (para + 14) & ") ^ 2) / 2) + 1) * " & lambda & ")" ' 1st har.
        char = Cells(4, 2).Value
        [A4:A4].Interior.ColorIndex = 45
        [B4:C4].Interior.ColorIndex = 44
        [A3:A3].Interior.ColorIndex = 3
        [B3:C3].Interior.ColorIndex = 38
    Else
        Cells(4, 2).Value = char
        Cells(42, para + 14).FormulaR1C1 = "= Sqrt((((950 *((" & gamma & ")^2))/(R4C2 * " & lambda & "))-1) * 2)"    ' K
        Cells(41, para + 14).FormulaR1C1 = "= R42C" & (para + 14) & "/(" & lambda & " * 0.934)"                            ' B (T)
        Cells(3, 2).FormulaR1C1 = "=(Ln((R41C" & (para + 14) & " - " & a0 & ")/(" & a1 & ")))/(" & a2 & ")"   ' 1st har.
        wf = Cells(3, 2).Value
        [A3:A3].Interior.ColorIndex = 45
        [B3:C3].Interior.ColorIndex = 44
        [A4:A4].Interior.ColorIndex = 3
        [B4:C4].Interior.ColorIndex = 38
    End If

    Cells(2, 1).Value = "PE shifts"
    Cells(2, 2).Value = pe
    Cells(3, 1).Value = "U60 gap"
    Cells(3, 3).Value = "mm"
    Cells(4, 1).Value = "1st har."
    Cells(4, 3).Value = "eV"
    Cells(41, para + 13).Value = "B (T)"
    Cells(42, para + 13).Value = "K"
    Cells(5, 1).Value = "Start PE"
    Cells(6, 1).Value = "End PE"
    Cells(7, 1).Value = "Step PE"
    Cells(8, 1).Value = "# scan"
    Cells(5, 2).Value = startEb
    Cells(6, 2).Value = endEb
    Cells(7, 2).Value = stepEk
    Cells(8, 2).Value = 1
    Cells(9, 1).Value = "Offset/multp"
    Cells(9, 2).Value = off
    Cells(9, 3).Value = multi
    Cells(10, 1).Value = "PE"
    Cells(10, 2).Value = "+shift"
    Cells(10, 3).Value = "Ab"
    strl(0) = "Photon energy (eV)"
    strl(1) = "Pe"
    strl(2) = "Sh"
    strl(3) = "Ab"
    [C2:C2].Value = "eV"
    [C5:C7].Value = "eV"
    [C8:C8].Value = "times"
    [A2:A2].Interior.ColorIndex = 3
    [B2:C2].Interior.ColorIndex = 38
    [A5:A8].Interior.ColorIndex = 41
    [B5:C8].Interior.ColorIndex = 37
    [A9:A9].Interior.ColorIndex = 43
    [B9:C9].Interior.ColorIndex = 35
    
    For n = 1 To numData
        C2(n, 1) = C2(n, 1) * maxXPSFactor  ' pA unit
        
        If IsNumeric(C3(n, 1)) = True Then
            If C3(n, 1) > 0 Then
            Else
                C3(n, 1) = 100
            End If
        Else
            C3(n, 1) = 100
        End If
        
        C4(n, 1) = (C2(n, 1) / C3(n, 1)) * 100 ' normalized to 100mA
    Next
    
    Range(Cells(11, 1), Cells(10 + numData, 1)) = C1
    Range(Cells(11, 3), Cells(10 + numData, 3)) = C4
    Cells(11, 2).FormulaR1C1 = "=R2C2 + RC[-1]"
    Range(Cells(11, 2), Cells(10 + numData, 2)).FillDown
    
    imax = numData + 10
    Cells(10 + (imax), 1).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
    Range(Cells(10 + (imax), 1), Cells((2 * imax) - 1, 1)).FillDown
    Cells(10 + (imax), 2).FormulaR1C1 = "=R2C + R[-" & (imax - 1) & "]C[-1]"
    Range(Cells(10 + (imax), 2), Cells((2 * imax) - 1, 2)).FillDown
    Cells(10 + (imax), 3).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C[-1]) *R9C"
    Range(Cells(10 + (imax), 3), Cells((2 * imax) - 1, 3)).FillDown
    Set dataData = Range(Cells(10 + (imax), 2), Cells(10 + (imax), 2).Offset(numData - 1, 1))
    startEb = Application.Floor(Cells(10 + (imax), 2).Value, char)
    endEb = Application.Ceiling(Cells(9 + (imax) + numData, 2).Value, char)
    
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers 'xlXYScatterSmoothNoMarkers
    ActiveChart.SetSourceData Source:=dataData, PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetGraphName
    ActiveChart.SeriesCollection(1).Name = "Ip"
    ActiveChart.ChartTitle.Delete
    
    With ActiveChart.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = strl(0)
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .HasMajorGridlines = True
        .MajorUnit = (char)
        .MinimumScale = startEb
        .MaximumScale = endEb
    End With
    
    With ActiveChart.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "Ip (pA/100mA)"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        If StrComp(testMacro, "debug", 1) = 0 Then
            .ScaleType = xlScaleLinear
        Else
            .HasMinorGridlines = True
            .MinorUnit = 1
            .MinimumScale = 1
            .ScaleType = xlScaleLogarithmic
        End If
    End With
    
    With ActiveSheet.ChartObjects(1)
        .Top = 20
        .Left = 200
        .Width = (550 * windowRatio) / windowSize
        .Height = 500 / windowSize
        .Chart.Legend.Delete
    End With
    
    SourceRangeColor1 = ActiveChart.SeriesCollection(1).Border.Color
    Range(Cells(10, 1), Cells(10, 2)).Interior.Color = SourceRangeColor1
    Range(Cells(9 + (imax), 1), Cells(9 + (imax), 2)).Interior.Color = SourceRangeColor1
    strTest = mid$(strSheetGraphName, InStr(strSheetGraphName, "_") + 1, Len(strSheetGraphName) - 6)
    Cells(8 + (imax), 2).Value = strTest + ".xlsx"
    Cells(9 + (imax), 1).Value = strl(1) + strTest
    Cells(9 + (imax), 2).Value = strl(2) + strTest
    Cells(9 + (imax), 3).Value = strl(3) + strTest

    Call SheetCheckGenerator
End Sub

Sub HigherOrderCheck()
    Dim strhighpe As String, C1 As Variant, strcheck As String
    
    strhighpe = Cells(2, 3).Value
    
    If Len(strhighpe) < 4 Then Exit Sub
    
    If mid$(strhighpe, 1, 1) = ";" And mid$(strhighpe, Len(strhighpe) - 2, 3) = " eV" Then
        n = 1
        j = 0
        strcheck = mid$(strhighpe, 2, Len(strhighpe) - 4)
        
        For iRow = 1 To Len(strcheck)
            strLabel = mid$(strcheck, iRow, 1)

            If IsNumeric(strLabel) = False Then
                If strLabel = ";" Or strLabel = "." Then
                Else
                    Exit Sub
                End If
            End If
        Next

        If InStr(1, strcheck, ";", 1) > 0 Then
            C1 = Split(strcheck, ";")
            If UBound(C1) > 8 Then Exit Sub  ' limit of higher order or ghost is 8
            For n = LBound(C1) To UBound(C1)
                If CSng(C1(n)) > 0 Then
                    ReDim Preserve highpe(j + 1)
                    highpe(j + 1) = CSng(C1(n))
                    j = j + 1
                End If
            Next
        Else
            If CSng(strcheck) > 0 Then
                ReDim Preserve highpe(j + 1)
                highpe(j + 1) = CSng(strcheck)
                j = j + 1
            End If
        End If
    End If
End Sub

Sub FormatData()   ' this is a template for data loading.
    Dim iniRow As Integer, endRow As Integer, totalDataPoints As Integer, eneCol As Single, speCol As Single, cnt As Integer, msgap As Integer
    
    If StrComp(strMode, "CLAM2", 1) = 0 Then        ' XPS mode
        strMode = "KE/eV"
        peX = CInt(mid$(Cells(8, 1).Value, 19, (Len(Cells(8, 1).Value) - 18 - 2)))
    ElseIf StrComp(strMode, "Photo", 1) = 0 Then    ' XAS mode
        strMode = "PE/eV"
    Else
        
    End If
    
    If graphexist = 0 And strMode = "KE/eV" Then
        ' if parameters are already specified in text, read from text. AlKa: 1486.6, MgKa: 1253.6 eV
        
        If peX = 0 Then
            peX = Application.InputBox(Title:="Manual input mode", Prompt:="Input a photon energy [eV] or cancel to switch AES mode", Default:=600, Type:=1)
        End If
        
        pe = peX
            
        If pe <= 0 Then
            strMode = "AE/eV"
        End If
        
        wf = 4
        char = 0
        
        ' initialize parameters adjustable
        off = 0
        multi = 1
        ncomp = 0
        highpe(0) = pe
        ' optional parameters
        cae = 100       ' pass energy
        g = 1200        ' grating line density
    End If
    
    If strMode = "KE/eV" Then
        ' Data position specified here by row and column in text data
        eneCol = 1  ' kinetic energy column
        speCol = 7  ' photoelectron spectral column
    
        iniRow = 22                     ' initial row position in text data at 22 row
        endRow = Cells(iniRow, speCol).End(xlDown).Row
        numData = endRow - iniRow + 1
        msgap = 3   ' gap between multple scanned data
    ElseIf strMode = "PE/eV" Then
        If Cells(7, 7).Value = "If/Ip" Then
            eneCol = 1  ' photon energy column
            speCol = 7  ' TFY spectral column
        Else
            eneCol = 1  ' photon energy column
            speCol = 5  ' TEY spectral column
        End If
    
        iniRow = 12                     ' initial row position in text data at 12 row
        endRow = Cells(iniRow, speCol).End(xlDown).Row
        numData = endRow - iniRow + 1
        msgap = 3   ' gap between multple scanned data
    Else
        ' Data position specified here by row and column in text data
        eneCol = 1  ' kinetic energy column
        speCol = 7  ' photoelectron spectral column
    
        iniRow = 22                     ' initial row position in text data
        endRow = Cells(iniRow, speCol).End(xlDown).Row
        numData = endRow - iniRow + 1
        msgap = 3   ' gap between multple scanned data
    End If
    
    ' Check multiple scanned data
    cnt = 0
    Do
        endRow = iniRow + (numData + msgap) * cnt + numData - 1
        If IsEmpty(Cells(endRow, speCol)) = True Then Exit Do
        cnt = cnt + 1
    Loop
    
    iniRow = iniRow + (numData + msgap) * (cnt - 1)
    endRow = iniRow + numData - 1
    scanNum = cnt
    'Debug.Print cnt, iniRow, endRow, numData
    
    Set dataKeData = Range(Cells(iniRow, eneCol), Cells(endRow, eneCol))  ' x-axis: kinetic energy
    Set dataIntData = dataKeData.Offset(, speCol - 1)                  ' y-axis: spectral intensity
    
    Set dataData = Union(dataKeData, dataIntData)
    
    ' measurement parameters
    startEk = Cells(iniRow, eneCol).Value       ' start kinetic energy
    endEk = Cells(endRow, eneCol).Value         ' end kinetic energy
    stepEk = Cells(iniRow + 1, eneCol).Value - Cells(iniRow, eneCol).Value      ' step of energy
    
    numData = CInt(((endEk - startEk) / stepEk) + 1)  ' number of points
End Sub

Sub KeBL()
    Dim C1 As Variant, s As Variant
    
    If graphexist = 0 Then
        If Cells(1, 2).Value = "AlKa" Then
            pe = 1486.6
            multi = 0.001
        ElseIf strMode = "KE/eV" Or strMode = "BE/eV" Then
            If StrComp(testMacro, "debug", 1) = 0 Then
                If peX = 0 Then
                    peX = Application.InputBox(Title:="Manual input mode", Prompt:="Input a photon energy [eV] or cancel to switch AES mode", Default:=650, Type:=1)
                End If
                pe = peX
            Else
                pe = Application.InputBox(Title:="Manual input mode", Prompt:="Input a photon energy [eV] or cancel to switch AES mode", Default:=650, Type:=1)
            End If
            
            highpe(0) = pe
            
            If pe <= 0 Then
                Cells(1, 1).Value = "AE/eV"
                strMode = "AE/eV"
            End If
            multi = 1
        End If
        
        If strMode = "BE/eV" Then
            wf = 4
        ElseIf strMode = "QE/eV" Then
            wf = 1
        Else
            wf = 4
        End If
        
        char = 0
        cae = 100
        off = 0
        ncomp = 0
    End If
    
    numData = ActiveSheet.UsedRange.Rows.Count - 1
    C1 = Range(Cells(2, 1), Cells(numData + 1, 3))
    
    If numData > 1 And IsNumeric(C1(1, 1)) Then
        For Each s In C1
            If IsEmpty(C1(numData, 1)) = False And IsEmpty(C1(numData, 2)) = False Then
                If IsNumeric(C1(numData, 1)) And IsNumeric(C1(numData, 2)) Then Exit For
            End If
            numData = numData - 1
        Next
    Else
        Call GetOut
        Exit Sub
    End If

    startEk = Cells(2, 1).Value
    endEk = Cells(numData + 1, 1).Value
    stepEk = Cells(3, 1).Value - Cells(2, 1).Value

    Set dataData = Range(Cells(2, 1), Cells(numData + 1, 2))
    Set dataKeData = Range(Cells(2, 1), Cells(numData + 1, 1))
    Set dataIntData = dataKeData.Offset(, 1)
    
    If strMode = "BE/eV" Then
        If startEk < endEk Then
            C1 = Range(Cells(2, 1), Cells(numData + 1, 3))
            
            If ExistSheet("Sort_" & strSheetDataName) Then
                Application.DisplayAlerts = False
                Worksheets("Sort_" & strSheetDataName).Delete
                Application.DisplayAlerts = True
            End If
    
            Worksheets.Add().Name = "Sort_" & strSheetDataName
            Range(Cells(2, 1), Cells(numData + 1, 3)) = C1
            Range(Cells(2, 1), Cells(numData + 1, 3)).Sort key1:=Cells(2, 1), order1:=xlDescending
            Set dataData = Range(Cells(2, 1), Cells(numData + 1, 2))
            Set dataKeData = Range(Cells(2, 1), Cells(numData + 1, 1))
            Set dataIntData = dataKeData.Offset(, 1)
            startEk = Cells(2, 1).Value
            endEk = Cells(numData + 1, 1).Value
            stepEk = Cells(3, 1).Value - Cells(2, 1).Value
            Cells(1, 1).Value = strMode & "/sort"
            Cells(1, 2).Value = "Y/sort"
            Cells(1, 3).Value = "Ie/sort"
        End If
    ElseIf InStr(strMode, "E/eV") > 0 Then
        If startEk > endEk Then
            C1 = Range(Cells(2, 1), Cells(numData + 1, 3))
            
            If ExistSheet("Sort_" & strSheetDataName) Then
                Application.DisplayAlerts = False
                Worksheets("Sort_" & strSheetDataName).Delete
                Application.DisplayAlerts = True
            End If
    
            Worksheets.Add().Name = "Sort_" & strSheetDataName
            Range(Cells(2, 1), Cells(numData + 1, 3)) = C1
            Range(Cells(2, 1), Cells(numData + 1, 3)).Sort key1:=Cells(2, 1), order1:=xlAscending
            Set dataData = Range(Cells(2, 1), Cells(numData + 1, 2))
            Set dataKeData = Range(Cells(2, 1), Cells(numData + 1, 1))
            Set dataIntData = dataKeData.Offset(, 1)
            startEk = Cells(2, 1).Value
            endEk = Cells(numData + 1, 1).Value
            stepEk = Cells(3, 1).Value - Cells(2, 1).Value
            Cells(1, 1).Value = strMode & "/sort"
            Cells(1, 2).Value = "Y/sort"
            Cells(1, 3).Value = "Ie/sort"
        End If
    End If
End Sub

Sub offsetmultiple()
    ActiveSheet.ChartObjects(1).Activate
    With ActiveSheet.ChartObjects(1)
        .Top = 150
    End With

    With ActiveChart.Axes(xlValue)
        .MinimumScaleIsAuto = True
        .MaximumScaleIsAuto = True
    End With
            
    If strl(1) = "Ke" Then
        ActiveSheet.ChartObjects(2).Activate
        With ActiveSheet.ChartObjects(2)
            .Top = 150 + (500 / windowSize)
        End With

        With ActiveChart.Axes(xlValue)
            .MinimumScaleIsAuto = True
            .MaximumScaleIsAuto = True
        End With
    End If
End Sub

Sub EachComp(ByRef OpenFileName As Variant, strAna As String, fcmp As Variant, sBG As Variant, cmp As Integer, ncmp As Integer, ncomp)
    Dim SourceRangeColor1 As Long, SourceRangeColor2 As Long, strCpa As String, sheetTarget As Worksheet, strNorm As String, strTest As String
    Dim Target As Variant, C1 As Variant, C2 As Variant, C3 As Variant, C4 As Variant, imax As Integer, NumSheets As Integer, peakNum As Integer, fitNum As Integer
    
    If strAna = "FitRatioAnalysis" Then
        peakNum = sheetFit.Cells(3, para + 1).Value         ' # of Fit peaks
        fitNum = sheetFit.Cells(4, para + 1).Value   ' # of Fit files
    End If
        
    C3 = fcmp   ' Name of peaks
    C4 = sBG    ' Name of BGs
    C1 = Split(Results, ",")
    ReDim strl(6)
    
    For n = 0 To 6
        strl(n) = C1(n)
    Next
    
    n = cmp     ' position of comp to be continued adding
    graphexist = 0
    
    For Each Target In OpenFileName
        If StrComp(Target, ActiveWorkbook.FullName, 1) = 0 Then
            graphexist = 1     ' in case the original file opens
            GoTo SkipOpen
        End If
        
        strTest = mid$(Target, InStrRev(Target, backSlash) + 1, Len(Target) - InStrRev(Target, backSlash))
        
        If Not WorkbookOpen(strTest) Then
            Workbooks.Open Target
            Workbooks(strTest).Activate
            j = 0
            If Err.Number > 0 Then
                MsgBox "Error in " & Target, vbOKOnly, "Error code: " & Err.Number
                GoTo SkipOpen
            ElseIf StrComp(ActiveWorkbook.Name, strTest, 1) <> 0 Then
                MsgBox "Error in " & Target
                GoTo SkipOpen
            End If
        Else
            Workbooks(strTest).Activate
            strLabel = ActiveSheet.Name
            j = 1
        End If
        
        strSheetDataName = mid$(Target, InStrRev(Target, backSlash) + 1, Len(Target) - InStrRev(Target, backSlash) - 5)
        If Len(strSheetDataName) > 25 Then strSheetDataName = mid$(strSheetDataName, 1, 25)
        
        If StrComp(mid$(strAna, 1, 3), "Fit", 1) = 0 Then    ' FitAnalysis, FitComp, FitRatioAnalysis
            If strAna = "FitRatioAnalysis" Then
                strCpa = "Ana_" + strSheetDataName
            ElseIf mid$(strSheetFitName, 1, 9) = "Fit_Norm_" Then
                strCpa = "Fit_Norm_" + strSheetDataName
            Else
                strCpa = "Fit_" + strSheetDataName
            End If
        ElseIf mid$(strSheetGraphName, 1, 11) = "Graph_Norm_" Then
            strCpa = "Graph_Norm_" + strSheetDataName    ' for Graph_Norm
        Else
            strCpa = "Graph_" + strSheetDataName    ' for .xlsx
        End If
        
        Target = mid$(Target, InStrRev(Target, backSlash) + 1, Len(Target) - InStrRev(Target, backSlash) - 5) + ".xlsx"
        
        If ExistSheet(strCpa) Then
            Workbooks(Target).Sheets(strCpa).Activate
            If Cells(2, 1).Value = "" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        Else
            NumSheets = Sheets.Count
            strCpa = ""
            For ns = 1 To NumSheets
                Sheets(ns).Activate
                If StrComp(mid$(strAna, 1, 3), "Fit", 1) = 0 Then
                    If strAna = "FitRatioAnalysis" And mid$(ActiveSheet.Name, 1, 4) = "Ana_" Then
                        If ExistSheet(mid$(ActiveSheet.Name, 5, Len(ActiveSheet.Name))) Then
                            strCpa = ActiveSheet.Name
                            Exit For
                        End If
                    ElseIf mid$(ActiveSheet.Name, 1, 9) = "Fit_Norm_" Then
                        If ExistSheet(mid$(ActiveSheet.Name, 10, Len(ActiveSheet.Name))) Then
                            strCpa = ActiveSheet.Name
                            Exit For
                        End If
                    Else
                        If mid$(ActiveSheet.Name, 1, 4) = "Fit_" Then
                            If ExistSheet(mid$(ActiveSheet.Name, 5, Len(ActiveSheet.Name))) Then
                                strCpa = ActiveSheet.Name
                                Exit For
                            End If
                        End If
                    End If
                ElseIf StrComp(strAna, "Graph_Norm", 1) = 0 Then
                    If mid$(ActiveSheet.Name, 1, 11) = "Graph_Norm_" Then
                        If ExistSheet(mid$(ActiveSheet.Name, 12, Len(ActiveSheet.Name))) Then
                            strCpa = ActiveSheet.Name
                            Exit For
                        End If
                    End If
                Else
                    If mid$(ActiveSheet.Name, 1, 6) = "Graph_" Then
                        If ExistSheet(mid$(ActiveSheet.Name, 7, Len(ActiveSheet.Name))) Then
                            strCpa = ActiveSheet.Name
                            Exit For
                        End If
                    End If
                End If
            Next
            
            If strCpa = "" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        End If
        
        If strl(5) = 1 Then
            If Not Cells(2, 1).Value = "PE shifts" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        ElseIf strl(5) = 2 Then
            If Not Cells(2, 1).Value = "PE" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        ElseIf strl(5) = 3 Then
            If Not Cells(2, 1).Value = "KE shifts" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        ElseIf strl(5) = 4 Then
            If Not Cells(2, 1).Value = "Shifts" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        ElseIf strl(5) = 5 Then
            If Not Cells(2, 1).Value = "x offset" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        End If

        If StrComp(mid$(strCpa, 1, 6), "Graph_", 1) = 0 Then
            If StrComp(mid$(strCpa, 1, 11), "Graph_Norm_", 1) = 0 Then
                Set sheetTarget = Workbooks(Target).Worksheets("Graph_Norm_" + strSheetDataName)
                Debug.Print "Graph_Norm_" + strSheetDataName
            Else
                Set sheetTarget = Workbooks(Target).Worksheets("Graph_" + strSheetDataName)
                Debug.Print "Graph_" + strSheetDataName
            End If
            
            If StrComp(sheetTarget.Cells(40, para + 9).Value, "Ver.", 1) = 0 Then
                iCol = para
            Else
                For iCol = 1 To 500
                Debug.Print sheetTarget.Cells(40, iCol + 9).Value, iCol
                    If StrComp(sheetTarget.Cells(40, iCol + 9).Value, "Ver.", 1) = 0 Then
                        Exit For
                    ElseIf iCol = 500 Then
                        MsgBox "Graph sheet has no parameters to be compared."
                        End
                    End If
                Next
            End If
            
            If mid$(sheetTarget.Cells(40, iCol + 10).Value, 1, 4) <= 8.05 And StrComp(mid$(strAna, 1, 3), "Fit", 1) = 0 Then
                MsgBox "Macro code used in some data comparison is obsolete!"
                End
            ElseIf mid$(sheetTarget.Cells(40, iCol + 10).Value, 1, 4) >= 6.56 Then
                numData = sheetTarget.Cells(41, iCol + 12).Value
            Else
                MsgBox "Macro code used in some data comparison is obsolete!"
                End
            End If
            
            strl(4) = sheetTarget.Cells(10, 1).Value       'check whether BE/eV or KE/eV. If BE/eV, only BE graph available
        ElseIf StrComp(mid$(strCpa, 1, 4), "Fit_", 1) = 0 Then
            If StrComp(mid$(strCpa, 1, 9), "Fit_Norm_", 1) = 0 Then
                Set sheetTarget = Workbooks(Target).Worksheets("Fit_Norm_" + strSheetDataName)
                Debug.Print "Fit_Norm_" + strSheetDataName
            Else
                Set sheetTarget = Workbooks(Target).Worksheets("Fit_" + strSheetDataName)
                Debug.Print "Fit_" + strSheetDataName
            End If
            
            If StrComp(sheetTarget.Cells(19, 100).Value, "Ver.", 1) = 0 Then
                iCol = para
            Else
                MsgBox "Fit sheet has no parameters to be compared."
                End
            End If
            
            numData = sheetTarget.Cells(5, 101).Value
            
        ElseIf StrComp(mid$(strCpa, 1, 9), "Ana_", 1) = 0 Then
            Debug.Print "Ana_" + strSheetDataName
            
            If StrComp(Cells(1, para + 1).Value, "Parameters", 1) = 0 Then
            Else
                For iCol = 1 To 500
                    If Cells(1, iCol).Value = "Parameters" Then
                        Exit For
                    ElseIf iCol = 500 Then
                        MsgBox "Ana sheet has no parameters to be compared."
                        End
                    End If
                Next
                para = iCol
            End If
        End If
        
        If strAna = "FitAnalysis" Then
            peakNum = Workbooks(Target).Sheets(strCpa).Cells(8 + sftfit2, 2).Value
            C1 = Workbooks(Target).Sheets(strCpa).Range(Cells(1, 5), Cells(18 + sftfit2, 4 + peakNum))
            C2 = Workbooks(Target).Sheets(strCpa).Range(Cells(1, 1), Cells(1, 3))
            
            For iCol = 0 To peakNum - 1
                For iRow = 0 To peakNum - 1
                    If C3(3, iCol + 5) = C1(1, iRow + 1) Then                                 ' Check Name of peak
                        C3(5 + n, iCol + 5) = C1(2, iRow + 1)                                 ' BE
                        If C1(16 + sftfit2, iRow + 1) > 0 Then
                            C3(5 + n + spacer + UBound(OpenFileName), iCol + 5) = C1(16 + sftfit2, iRow + 1)      ' T.I.Area
                            C3(5 + n + 2 * (spacer + UBound(OpenFileName)), iCol + 5) = C1(17 + sftfit2, iRow + 1)  ' S.I.Area
                            C3(5 + n + 3 * (spacer + UBound(OpenFileName)), iCol + 5) = C1(18 + sftfit2, iRow + 1)    ' N.I.Area
                        Else
                            C3(5 + n + spacer + UBound(OpenFileName), iCol + 5) = 0
                            C3(5 + n + 2 * (spacer + UBound(OpenFileName)), iCol + 5) = 0  ' S.Area
                            C3(5 + n + 3 * (spacer + UBound(OpenFileName)), iCol + 5) = 0    ' N.Area
                        End If
                        
                        C3(5 + n + 4 * (spacer + UBound(OpenFileName)), iCol + 5) = C1(4, iRow + 1)     ' FWHM
                        Exit For
                    End If
                Next
            Next
    
            For p = 0 To 4
                C3(5 + (spacer + UBound(OpenFileName)) * p + n, 1) = Target
                C3(5 + (spacer + UBound(OpenFileName)) * p + n, 2) = strCpa
                C3(5 + (spacer + UBound(OpenFileName)) * p + n, 4) = Workbooks(Target).Sheets(strCpa).Cells(8 + sftfit2, 2).Value
            Next
            
            For p = 0 To 2
                C3(5 + n, peakNum + 6 + p) = C2(1, 1 + p)
            Next

            If j = 0 Then
                Workbooks(Target).Close True
            Else
                Workbooks(Target).Sheets(strLabel).Activate
                j = 0
            End If
            
            n = n + 1
            GoTo SkipOpen
        ElseIf strAna = "FitRatioAnalysis" Then
            Dim spacera As Integer
            Dim peakNuma As Integer
            Dim fitNuma As Integer
            Dim iCola As Integer
            Dim iRowa As Integer
            
            spacera = Workbooks(Target).Sheets(strCpa).Cells(2, para + 1).Value     ' spacer
            peakNuma = Workbooks(Target).Sheets(strCpa).Cells(3, para + 1).Value    ' # of peaks
            fitNuma = Workbooks(Target).Sheets(strCpa).Cells(4, para + 1).Value    ' # of files
            C1 = Workbooks(Target).Sheets(strCpa).Range(Cells(1, 1), Cells((4 + spacera * 4) + 5 * fitNuma, 9 + 2 * peakNuma)) ' No check in matching among the peak names.
            C2 = Workbooks(Target).Sheets(strCpa).Range(Cells(4, 6 + peakNuma), Cells(3 + fitNuma, 8 + peakNuma))
            C3(1, peakNum + 5) = Target
            C3(2, peakNum + 5) = strCpa

            For iCola = 0 To peakNuma - 1
                For iRowa = 0 To fitNum   ' include the peak name
                    C3(3 + iRowa, iCola + peakNum + 5) = C1(3 + iRowa, iCola + 5)                                 ' BE
                    C3(2 + iRowa + 1 * (spacer + fitNum), iCola + peakNum + 5) = C1(2 + iRowa + 1 * (spacera + fitNuma), iCola + 5)      ' P.Area
                    C3(1 + iRowa + 2 * (spacer + fitNum), iCola + peakNum + 5) = C1(1 + iRowa + 2 * (spacera + fitNuma), iCola + 5)  ' S.Area
                    C3(0 + iRowa + 3 * (spacer + fitNum), iCola + peakNum + 5) = C1(0 + iRowa + 3 * (spacera + fitNuma), iCola + 5)    ' N.Area
                    C3(-1 + iRowa + 4 * (spacer + fitNum), iCola + peakNum + 5) = C1(-1 + iRowa + 4 * (spacera + fitNuma), iCola + 5)     ' FWHM
                Next
            Next
            
            For p = 0 To fitNum - 1
                C4(p + 1, n + 2) = C2(1 + p, 1) & C2(1 + p, 2) & C2(1 + p, 3)
            Next
            
            If j = 0 Then
                Workbooks(Target).Close True
            Else
                Workbooks(Target).Sheets(strLabel).Activate
                j = 0
            End If
            
            peakNum = peakNum + peakNuma
            n = n + 1
            GoTo SkipOpen
        ElseIf strAna = "FitComp" Then
            numData = Workbooks(Target).Sheets(strCpa).Cells(5, 101).Value
            C1 = Workbooks(Target).Sheets(strCpa).Range(Cells(20 + sftfit, 1), Cells(20 + sftfit + numData, 1)).Value  'tmp
            C2 = Workbooks(Target).Sheets(strCpa).Range(Cells(20 + sftfit, 4), Cells(20 + sftfit + numData, 4)).Value   'en

            sheetAna.Activate
            sheetAna.Range(Cells(10, (4 + (n * 3))), Cells(10 + numData, (4 + (n * 3)))).Value = C1
            sheetAna.Range(Cells(10, (6 + (n * 3))), Cells(10 + numData, (6 + (n * 3)))).Value = C2

            If StrComp(mid$(Cells(10, (4 + (n * 3))).Value, 1, 2), "BE", 1) = 0 Then
                strl(1) = "Be"
                strl(2) = "Sh"
                strl(3) = "In"
                Cells(4, (4 + (n * 3))) = "Shift"
                Cells(4, (5 + (n * 3))) = 0
                Cells(4, (6 + (n * 3))) = "eV"
                Cells(10, (5 + (n * 3))) = "Shift"
                Range(Cells(4, (4 + (n * 3))), Cells(4, (4 + (n * 3)))).Interior.ColorIndex = 3
                Range(Cells(4, (5 + (n * 3))), Cells(4, (6 + (n * 3)))).Interior.ColorIndex = 38
            ElseIf StrComp(mid$(Cells(10, (4 + (n * 3))).Value, 1, 2), "PE", 1) = 0 Then
                strl(1) = "Pe"
                strl(2) = "Sh"
                strl(3) = "Ab"
                Cells(2, (4 + (n * 3))).Value = "Shift"
                Cells(2, (5 + (n * 3))).Value = 0
                Cells(2, (6 + (n * 3))).Value = "eV"
                Cells(10, (5 + (n * 3))).Value = "Shift"
                Range(Cells(2, (4 + (n * 3))), Cells(2, (4 + (n * 3)))).Interior.ColorIndex = 3
                Range(Cells(2, (5 + (n * 3))), Cells(2, (6 + (n * 3)))).Interior.ColorIndex = 38
            ElseIf StrComp(mid$(Cells(10, (4 + (n * 3))).Value, 1, 2), "ME", 1) = 0 Then
                strl(1) = "Po"
                strl(2) = "Sh"
                strl(3) = "Ab"
                Cells(2, (4 + (n * 3))).Value = "Shift"
                Cells(2, (5 + (n * 3))).Value = 0
                Cells(2, (6 + (n * 3))).Value = "a.u."
                Cells(10, (5 + (n * 3))).Value = "Shift"
                Range(Cells(2, (4 + (n * 3))), Cells(2, (4 + (n * 3)))).Interior.ColorIndex = 3
                Range(Cells(2, (5 + (n * 3))), Cells(2, (6 + (n * 3)))).Interior.ColorIndex = 38
            End If
            
            strSheetGraphName = strSheetAnaName
            strl(6) = 2
        Else
            Workbooks(Target).Sheets(strCpa).Range(Cells(2, 1), Cells(10 + numData, 3)).Copy Destination:=Workbooks(wb).Sheets(strSheetGraphName).Cells(2, (4 + (n * 3)))
            Workbooks(wb).Sheets(strSheetGraphName).Activate
        End If
        
        strCasa = Cells(1, (5 + (n * 3))).Value
        Cells(1, (5 + (n * 3))).Value = Target
        Cells(9, (4 + (n * 3))).Value = "Offset/multp"
        Cells(9, (5 + (n * 3))).Value = 0
            
        If WorksheetFunction.Round(Cells(2, (5 + (n * 3))), 1) = 1486.6 And StrComp(mid$(strAna, 1, 3), "Fit", 1) <> 0 Then
            Cells(9, (6 + (n * 3))).Value = 0.001
        ElseIf WorksheetFunction.Round(Cells(2, (5 + (n * 3))), 1) = 1468.8 Then    ' fix mis-spelling for Alka PE
            Cells(2, (5 + (n * 3))).Value = 1486.6
            Cells(9, (6 + (n * 3))).Value = 0.001
        Else
            Cells(9, (6 + (n * 3))).Value = 1
        End If

        Cells(9, (4 + (n * 3))).Interior.Color = RGB(139, 195, 74)  ' added 20160324
        Range(Cells(9, (5 + (n * 3))), Cells(9, (6 + (n * 3)))).Interior.Color = RGB(197, 225, 165)  ' added 20160324
        
        If Cells(3, (4 + (n * 3))).Interior.ColorIndex = 45 Then
            Cells(3, (5 + (n * 3))).FormulaR1C1 = "=(Ln((((Sqrt((((950 *((" & gamma & ")^2))/(R4C * " & lambda & "))-1) * 2))/(" & lambda & " * 0.934)) - " & a0 & ")/(" & a1 & ")))/(" & a2 & ")" ' gap
        ElseIf Cells(4, (4 + (n * 3))).Interior.ColorIndex = 45 Then
            Cells(4, (5 + (n * 3))).FormulaR1C1 = "=950 * ((" & gamma & ") ^ 2) / (((((0.934 * " & lambda & " * (" & a0 & " + " & a1 & " * Exp(" & a2 & " * R3C))) ^ 2) / 2) + 1) * " & lambda & ")" ' 1st har.
        End If
        
        imax = numData + 10
        
        If strl(1) = "Ke" And strl(3) = "In" Then
            If strl(4) = "Be" Then
                Cells(11, (5 + (n * 3))).FormulaR1C1 = "=R4C + RC[-1]"
                Cells(10 + (imax), (5 + (n * 3))).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
            ElseIf strl(4) = "Ek" Then ' this is a trigger to handle "BE/eV" data
                Cells(11, (4 + (n * 3))).FormulaR1C1 = "=R2C[1] - RC[1] - R3C[1]"
                Cells(10 + (imax), (5 + (n * 3))).FormulaR1C1 = "=-R4C + R[-" & (imax - 1) & "]C"
            Else
                Cells(11, (5 + (n * 3))).FormulaR1C1 = "=R2C - R3C - R4C - RC[-1]"
                Cells(10 + (imax), (5 + (n * 3))).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
            End If
        ElseIf strl(1) = "Pe" Or strl(1) = "Po" Then
            If strl(3) = "Pp" Then
                Cells(11, (5 + (n * 3))).FormulaR1C1 = "=R3C * (R2C + RC[-1])"
                Cells(10 + (imax), (5 + (n * 3))).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
            Else
                Cells(11, (5 + (n * 3))).FormulaR1C1 = "=R2C + RC[-1]"
                Cells(10 + (imax), (5 + (n * 3))).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
            End If
        ElseIf strl(1) = "Be" Then
            If strl(4) = "Ke" Or strl(4) = "Ek" Then  ' old data used with "Ek"
                Cells(11, (5 + (n * 3))).FormulaR1C1 = "=R2C - R3C - R4C - RC[-1]"
                Cells(10 + (imax), (5 + (n * 3))).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
            Else
                Cells(11, (5 + (n * 3))).FormulaR1C1 = "=R4C + RC[-1]"
                Cells(10 + (imax), (5 + (n * 3))).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
            End If
        ElseIf strl(3) = "De" Then
            Cells(10 + (imax), (4 + (n * 3))).FormulaR1C1 = "=R2C + R[-" & (imax - 1) & "]C"
            Range(Cells(10 + (imax), (4 + (n * 3))), Cells((2 * imax) - 1, (4 + (n * 3)))).FillDown
            Cells(10 + (imax), (5 + (n * 3))).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C) * R9C[1]"
            Range(Cells(10 + (imax), (5 + (n * 3))), Cells((2 * imax) - 1, (5 + (n * 3)))).FillDown
            Cells(10 + (imax), (6 + (n * 3))).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C) * R9C"
            Range(Cells(10 + (imax), (6 + (n * 3))), Cells((2 * imax) - 1, (6 + (n * 3)))).FillDown
            GoTo AESmode
        End If
        
        If strl(1) = "Ke" And strl(3) = "In" And strl(4) = "Ek" Then
            Range(Cells(11, (4 + (n * 3))), Cells((imax), (4 + (n * 3)))).FillDown
        Else
            Range(Cells(11, (5 + (n * 3))), Cells((imax), (5 + (n * 3)))).FillDown
        End If
        
        Cells(10 + (imax), (4 + (n * 3))).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
        Range(Cells(10 + (imax), (4 + (n * 3))), Cells((2 * imax) - 1, (4 + (n * 3)))).FillDown
        
        Range(Cells(10 + (imax), (5 + (n * 3))), Cells((2 * imax) - 1, (5 + (n * 3)))).FillDown
        Cells(10 + (imax), (6 + (n * 3))).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C[-1]) * R9C"
        Range(Cells(10 + (imax), (6 + (n * 3))), Cells((2 * imax) - 1, (6 + (n * 3)))).FillDown
        
AESmode:
        
        Range(Cells((2 * imax), (4 + (n * 3))), Cells((2 * imax), (6 + (n * 3))).End(xlDown)).Clear
        Range(Cells((imax + 1), (4 + (n * 3))), Cells((imax + 9), (6 + (n * 3)))).Clear
        
        Set dataKeGraph = Range(Cells(10 + (imax), (4 + (n * 3))), Cells((2 * imax - 1), (4 + (n * 3))))
        
        If j = 0 Then
            Workbooks(Target).Close True
        Else
            Workbooks(Target).Sheets(strLabel).Activate
            j = 0
        End If
        
        Workbooks(wb).Sheets(strSheetGraphName).Activate
        ActiveSheet.ChartObjects(1).Activate
        
        If n > ncomp - 1 Then       ' n: position of comp to be added, ncomp: already added
            ActiveChart.SeriesCollection.NewSeries
            p = ActiveChart.SeriesCollection.Count
        Else
            p = 1
            If Cells(42, para + 12).Value > 0 Then p = p + 1
            If Cells(43, para + 12).Value > 0 Then p = p + 1
            If Cells(44, para + 12).Value > 0 Then p = p + 1
            p = p + n + 1
        End If

        With ActiveChart.SeriesCollection(p)
            .ChartType = xlXYScatterLinesNoMarkers
            If strl(3) = "De" Then
                .Name = "='" & ActiveSheet.Name & "'!R1C" & (5 + (n * 3)) & ""
                .XValues = dataKeGraph.Offset(0, 0)
                .Values = dataKeGraph.Offset(0, 1)
            Else
                .Name = "='" & ActiveSheet.Name & "'!R1C" & (5 + (n * 3)) & ""
                .XValues = dataKeGraph.Offset(0, 1)
                .Values = dataKeGraph.Offset(0, 2)
            End If
            SourceRangeColor1 = .Border.Color
        End With
        
        If strl(1) = "Ke" And (strl(4) = "Ke" Or strl(4) = "Ek") Then
            ActiveSheet.ChartObjects(2).Activate
            If n > ncomp - 1 Then
               ActiveChart.SeriesCollection.NewSeries
            End If
            
            With ActiveChart.SeriesCollection(p)
                .ChartType = xlXYScatterLinesNoMarkers
                '.Name = Cells(1, 5 + (n * 3)).Value
                .Name = "='" & ActiveSheet.Name & "'!R1C" & (5 + (n * 3)) & ""
                .XValues = dataKeGraph
                .Values = dataKeGraph.Offset(0, 2)
                SourceRangeColor2 = .Border.Color
            End With
            Range(Cells(10, (4 + (n * 3))), Cells(10, ((4 + (n * 3))))).Interior.Color = SourceRangeColor2
            Range(Cells(9 + (imax), (4 + (n * 3))), Cells(9 + (imax), ((4 + (n * 3))))).Interior.Color = SourceRangeColor2
        Else
            Range(Cells(10, (4 + (n * 3))), Cells(10, ((4 + (n * 3))))).Interior.Color = SourceRangeColor1
            Range(Cells(9 + (imax), (4 + (n * 3))), Cells(9 + (imax), ((4 + (n * 3))))).Interior.Color = SourceRangeColor1
        End If

        Range(Cells(10, (5 + (n * 3))), Cells(10, ((5 + (n * 3))))).Interior.Color = SourceRangeColor1
        Range(Cells(9 + (imax), (5 + (n * 3))), Cells(9 + (imax), ((5 + (n * 3))))).Interior.Color = SourceRangeColor1
        
        If StrComp(strAna, "Graph_Norm", 1) = 0 Then
            strNorm = "Norm_"
        Else
            strNorm = vbNullString
        End If
        
        strTest = mid$(Cells(1, (5 + (n * 3))).Value, 1, Len(Cells(1, (5 + (n * 3))).Value) - 5)
        Cells(8 + (imax), (5 + (n * 3))).Value = Cells(1, (5 + (n * 3))).Value
        Cells(9 + (imax), (4 + (n * 3))).Value = strl(1) + strNorm + strTest
        Cells(9 + (imax), (5 + (n * 3))).Value = strl(2) + strNorm + strTest
        Cells(9 + (imax), (6 + (n * 3))).Value = strl(3) + strNorm + strTest
        n = n + 1
SkipOpen:
    Next Target
    
    fcmp = C3       ' peak parameters added
    ncmp = n        ' number of data added over cmp
    sBG = C4        ' peak BGs
    
    If strAna = "FitRatioAnalysis" Then
        sheetFit.Cells(3, para + 1).Value = peakNum   ' # of peaks final
    End If
End Sub

Sub descriptGraph()
    Dim strhighpe As String, imax As Integer
    
    strhighpe = ""
    Cells(2, 1).Value = "PE"
    Cells(3, 1).Value = "WF"
    Cells(4, 1).Value = "Char"
    Cells(5, 1).Value = "Start KE"
    Cells(6, 1).Value = "End KE"
    Cells(7, 1).Value = "Step KE"
    Cells(8, 1).Value = "# scan"
    
    If StrComp(Cells(2, 1).Value, "PE", 1) = 0 Then
        If UBound(highpe) > 0 Then
            For n = 1 To UBound(highpe)
                strhighpe = strhighpe & ";" & highpe(n)
            Next
            Cells(2, 3).Value = strhighpe & " eV"   'strhighpe
            [C3:C7].Value = "eV"
        Else
            [C2:C7].Value = "eV"
        End If
    End If
    
    Cells(10, 1).Value = "Ke"
    Cells(10, 2).Value = "Be"
    Cells(10, 3).Value = "In"
    g = 0
    Cells(1, 2).Value = g
    Cells(2, 2).Value = pe
    Cells(3, 2).Value = wf
    Cells(4, 2).Value = char
    Cells(5, 2).Value = startEk
    Cells(6, 2).Value = endEk
    Cells(7, 2).Value = stepEk
    Cells(8, 2).Value = scanNum
    Cells(8, 3).Value = "times"
    [B5:C8].Interior.Color = RGB(144, 202, 249)
    Cells(9, 1).Value = "Offset/multp"
    Cells(9, 2).Value = off
    Cells(9, 3).Value = multi
    
    Call descriptHidden1
    
    [A2:A4].Interior.Color = RGB(244, 67, 54)
    [B2:C4].Interior.Color = RGB(244, 143, 177)
    [A5:A8].Interior.Color = RGB(3, 169, 244)
    Range(Cells(9, 1), Cells(9, 1)).Interior.Color = RGB(139, 195, 74)
    Range(Cells(9, 2), Cells(9, 3)).Interior.Color = RGB(197, 225, 165)
    ReDim strl(3)
    imax = numData + 10
    
    If strMode = "PE/eV" Or strMode = "GE/eV" Then
        Cells(2, 2).Value = pe
        Cells(2, 1).Value = "PE shifts"
        Cells(5, 1).Value = "Start PE"
        Cells(6, 1).Value = "End PE"
        Cells(7, 1).Value = "Step PE"
        [C2:C7].Value = "eV"
        Range(Cells(3, 1), Cells(4, 3)).Clear
        
        Cells(10, 1).Value = "Pe"
        Cells(10, 2).Value = "+shift"
        Cells(10, 3).Value = "Ab"
        Cells(11, 2).FormulaR1C1 = "=R2C2 + RC[-1]"
        Cells(10 + (imax), 2).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
        strl(0) = "Photon energy (eV)"
        strl(1) = "Pe"
        strl(2) = "Sh"
        strl(3) = "Ab"
    ElseIf strMode = "QE/eV" Then
        Cells(2, 2).Value = pe
        Cells(2, 1).Value = "x offset"
        Cells(3, 2).Value = wf
        Cells(3, 1).Value = "x multiple"
        Cells(5, 1).Value = "Start"
        Cells(6, 1).Value = "End"
        Cells(7, 1).Value = "Step"
        [C2:C7].Value = "a.u."
        Range(Cells(4, 1), Cells(4, 3)).Clear
        Cells(10, 1).Value = "Mass"
        Cells(10, 2).Value = "+offset/multiple"
        Cells(10, 3).Value = "PP"
        Cells(11, 2).FormulaR1C1 = "=R3C2 * (R2C2 + RC[-1])"
        Cells(10 + (imax), 2).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
        strl(0) = "Mass (arb. unit)"
        strl(1) = "Po"
        strl(2) = "Pn"
        strl(3) = "Pp"
    ElseIf strMode = "BE/eV" Then
        Cells(2, 2).Value = pe
        Cells(2, 1).Value = "PE"
        Cells(3, 1).Value = "WF"
        Cells(4, 1).Value = "Char"
        Cells(5, 1).Value = "Start BE"
        Cells(6, 1).Value = "End BE"
        Cells(7, 1).Value = "Step BE"
        Cells(10, 1).Value = "Ek"   ' this is a trigger to handle getCompare correctly
        Cells(10, 2).Value = "Be"
        Cells(10, 3).Value = "In"
        Range(Cells(11, 2), Cells(11, 2).Offset(numData - 1, 0)).Value = Range(Cells(11, 1), Cells(11, 1).Offset(numData - 1, 0)).Value
        Cells(11, 1).FormulaR1C1 = "=R2C2 - RC[1] - R3C2"
        
        Cells(10 + (imax), 2).FormulaR1C1 = "=-R4C2 + R[-" & (imax - 1) & "]C"
        strl(0) = "Binding energy (eV)"
        strl(1) = "Ke"
        strl(2) = "Be"
        strl(3) = "In"
        
    ElseIf strMode = "AE/eV" Then
        Cells(2, 2).Value = pe
        Cells(2, 1).Value = "KE shifts"
        Cells(3, 2).Value = wf
        Cells(3, 1).Value = "Smoothing"
        Cells(5, 1).Value = "Start KE"
        Cells(6, 1).Value = "End KE"
        Cells(7, 1).Value = "Step KE"
        [C2:C7].Value = "eV"
        [A3:A3].Interior.ColorIndex = 45
        [B3:C3].Interior.ColorIndex = 44

        Range(Cells(4, 1), Cells(4, 3)).Clear
        Cells(1, 1).Value = "AES elec."
        If g = 0 Then
            g = 5
        ElseIf g = 10 Then
            strAES = "VG10kCrr4"
        End If
        
        Cells(1, 2).Value = g
        Cells(1, 3).Value = "keV"
        Cells(3, 3).Value = "points"
        Cells(10, 1).Value = "Ke"
        Cells(10, 2).Value = "Ae"
        Cells(10, 3).Value = "De"

        strl(0) = "Kinetic energy (eV)"
        strl(1) = "Ke"
        strl(2) = "Ae"
        strl(3) = "De"
    ElseIf strMode = "ME/eV" Then
        Cells(2, 2).Value = pe
        Cells(2, 1).Value = "Shifts"
        Cells(5, 1).Value = "Start"
        Cells(6, 1).Value = "End"
        Cells(7, 1).Value = "Step"
        [C2:C7].Value = "a.u."
        Range(Cells(3, 1), Cells(4, 3)).Clear
        
        Cells(10, 1).Value = "Po"
        Cells(10, 2).Value = "+shift"
        Cells(10, 3).Value = "Ab"
        Cells(11, 2).FormulaR1C1 = "=R2C2 + RC[-1]"
        Cells(10 + (imax), 2).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
        strl(0) = "Position (arb. unit)"
        strl(1) = "Po"
        strl(2) = "Sh"
        strl(3) = "Ab"
    Else
        Cells(11, 2).FormulaR1C1 = "=R2C2 - R3C2 - R4C2 - RC[-1]"
        Cells(10 + (imax), 2).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
        strl(0) = "Binding energy (eV)"
        strl(1) = "Ke"
        strl(2) = "Be"
        strl(3) = "In"
    End If
    
    If strl(3) = "De" Then
        Cells(10 + (imax), 1).FormulaR1C1 = "=R2C2 + R[-" & (imax - 1) & "]C"
        Range(Cells(10 + (imax), 1), Cells((2 * imax) - 1, 1)).FillDown
            
        Cells(10 + (imax), 2).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C) *R9C[1]"
        Range(Cells(10 + (imax), 2), Cells((2 * imax) - 1, 2)).FillDown
        Cells(10 + (imax), 3).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C) *R9C"
        Range(Cells(10 + (imax), 3), Cells((2 * imax) - 1, 3)).FillDown
        
        Set dataBGraph = Range(Cells(10 + (imax), 1), Cells(10 + (imax), 1).Offset(numData - 1, 1))
        Set dataKGraph = Union(Range(Cells(10 + (imax), 1), Cells(10 + (imax), 1).Offset(numData - 1, 0)), Range(Cells(10 + (imax), 3), Cells(10 + (imax), 3).Offset(numData - 1, 0)))
        Set dataKeGraph = Range(Cells(10 + (imax), 1), Cells(10 + (imax), 1).Offset(numData - 1, 0))
        Set dataBeGraph = dataKeGraph.Offset(, 1)
    Else
        If strMode = "BE/eV" Then
            Range(Cells(11, 1), Cells((10 + numData), 1)).FillDown
        Else
            Range(Cells(11, 2), Cells((10 + numData), 2)).FillDown
        End If
        
        Cells(10 + (imax), 1).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
        Range(Cells(10 + (imax), 1), Cells((2 * imax) - 1, 1)).FillDown
        Range(Cells(10 + (imax), 2), Cells((2 * imax) - 1, 2)).FillDown
        Cells(10 + (imax), 3).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C[-1]) *R9C"
        Range(Cells(10 + (imax), 3), Cells((2 * imax) - 1, 3)).FillDown
        
        Set dataBGraph = Range(Cells(10 + (imax), 2), Cells(10 + (imax), 2).Offset(numData - 1, 1))
        Set dataKGraph = Union(Range(Cells(10 + (imax), 1), Cells(10 + (imax), 1).Offset(numData - 1, 0)), Range(Cells(10 + (imax), 3), Cells(10 + (imax), 3).Offset(numData - 1, 0)))
        Set dataKeGraph = Range(Cells(10 + (imax), 1), Cells(10 + (imax), 1).Offset(numData - 1, 0))
        Set dataBeGraph = dataKeGraph.Offset(, 1)
        
        If strMode = "BE/eV" Then
            startEk = Cells(11, 1).Value
            endEk = Cells(10 + numData, 1).Value
        End If
    End If
End Sub

Sub descriptHidden1()
    Cells(1, 1).Value = "Grating"
    Cells(1, 3).Value = "lines/mm"
    Cells(1, 2).Value = g
    Cells(40, para + 10).Value = ver
    Cells(40, para + 9).Value = "Ver."
    Cells(41, para + 9).Value = "dblMin"
    Cells(42, para + 9).Value = "dblMax"
    Cells(43, para + 9).Value = "maxXPSFactor"
    Cells(44, para + 9).Value = "maxAESFactor"
    Cells(45, para + 9).Value = "ncomp"
    Cells(46, para + 9).Value = "XPS BE database:"
    Cells(47, para + 9).Value = "AES KE database:"
    Cells(41, para + 11).Value = "numData"
    Cells(42, para + 11).Value = "numChemFactors"
    Cells(43, para + 11).Value = "numXPSFactors"
    Cells(44, para + 11).Value = "numAESFactors"
    Cells(45, para + 11).Value = "Gnum"
    Cells(41, para + 12).Value = numData
    ncomp = 0
    Cells(45, para + 10).Value = ncomp
    Cells(45, para + 12).Value = 0
    Cells(46, para + 11).Value = strCasa
    Cells(47, para + 11).Value = strAES
    Cells(50, para + 11).Value = "Elem"
    Cells(50, para + 12).Value = "BE"
    Cells(50, para + 13).Value = "KE"
    Cells(50, para + 14).Value = "BEchar"
    Cells(50, para + 15).Value = "KEchar"
    Cells(50, para + 16).Value = "RSF"
    Cells(50, para + 17).Value = "Nom"
    Cells(50, para + 18).Value = "aes_dif"
    Cells(50, para + 19).Value = "beta"
    Cells(50, para + 20).Value = "atm_rto"
End Sub

Sub descriptHidden2()
    Cells(41, para + 10).Value = (dblMin / Cells(9, 3).Value) + Cells(9, 2).Value
    Cells(42, para + 10).Value = (dblMax / Cells(9, 3).Value) + Cells(9, 2).Value
    Cells(43, para + 10).Value = maxXPSFactor
    Cells(44, para + 10).Value = maxAESFactor
    Cells(42, para + 12).Value = 0      'numChemFactors
    Cells(43, para + 12).Value = numXPSFactors
    Cells(44, para + 12).Value = numAESFactors
    Cells(46, para + 11).Value = strCasa
    Cells(47, para + 11).Value = strAES
    Cells(51, para + 9).Value = ElemD
End Sub

Sub descriptFit()
    Dim tfa As Single, tfb As Single

    Cells(19, 101).Value = ver
    
    If strl(1) = "Po" Then
        Cells(1, 1).Value = "Polynominal"
        Cells(1, 2).Value = "BG"
        Cells(1, 3).Value = vbNullString
        Cells(2, 1).Value = "a0"
        Cells(3, 1).Value = "a1"
        Cells(4, 1).Value = "a2"
        Cells(5, 1).Value = "a3"
        Range(Cells(2, 2), Cells(5, 2)) = 0
        Range(Cells(3, 2), Cells(5, 2)).Font.Bold = "True"
    Else
        Cells(1, 1).Value = "Shirley"
        Cells(1, 2).Value = "BG"
        Cells(1, 3).Value = vbNullString
        Cells(2, 1).Value = "Tolerance"
        Cells(3, 1).Value = "Initial A"
        Cells(4, 1).Value = "Final A"
        Cells(5, 1).Value = "Iteration"
        Cells(2, 2).Value = 0.000001
        Cells(3, 2).Value = 0.001
    End If
    
    Cells(6 + sftfit2, 1).Value = "Solve BGS"
    Cells(7 + sftfit2, 1).Value = "Peak fit"
    Cells(8 + sftfit2, 1).Value = "# peaks"
    Cells(9 + sftfit2, 1).Value = "Solve LSM"
    Cells(10 + sftfit2, 1).Value = "Fit range"
    Cells(11 + sftfit2, 1).Value = "Start / eV"
    Cells(12 + sftfit2, 1).Value = "End / eV"
    Cells(13 + sftfit2, 1).Value = "Factors for N.Area"
    Cells(14 + sftfit2, 1).Value = "CAE"
    Cells(15 + sftfit2, 1).Value = "Grating"
    Cells(16 + sftfit2, 1).Value = "IMFP"
    Cells(17 + sftfit2, 1).Value = "a"
    Cells(18 + sftfit2, 1).Value = "b"
    Cells(19 + sftfit2, 1).Value = "theta"
    
    If IsNumeric(ver) Then
        If ver < 7.58 Then
            sftfit = 0
            sftfit2 = 0
        End If
    ElseIf IsNumeric(Left$(ver, 4)) Then
        If Left$(ver, 4) < 7.58 Then
            sftfit = 0
            sftfit2 = 0
        End If
    End If
    
    If strl(1) = "Pe" Then
        Cells(20 + sftfit, 1).Value = "PE / eV"
        Cells(20 + sftfit, 2).Value = "Ab"
    ElseIf strl(1) = "Po" Then
        Cells(20 + sftfit, 1).Value = "ME / eV"
        Cells(20 + sftfit, 2).Value = "Ab"
    Else
        Cells(20 + sftfit, 1).Value = "BE / eV"
        Cells(20 + sftfit, 2).Value = "In"
    End If
    
    Cells(15 + sftfit2, 2).Value = 0     ' Grating number, 0 means VersaProbe II
    
    If Cells(15 + sftfit2, 2).Value = 0 Then    ' VersaProbe II AlKa
        Cells(14 + sftfit2, 2).Value = 23.5     ' CAE must be setup by user
    Else
        Cells(14 + sftfit2, 2).Value = cae
    End If
                                            ' Inelastic mean free path parameter:
    Cells(16 + sftfit2, 2).Value = mfp      ' lambda is proportional to E^x, and x can be from 0.5 to 0.9.
    Cells(19 + sftfit2, 2).Value = 54.7356     ' Angle between x-ray and analyzer lens: magic angle 54.7356 deg.
    
    If Cells(15 + sftfit2, 2).Value = 0 Then    ' VersaProbe II AlKa
        tfa = 180.254
        tfb = 0.348
        Cells(19 + sftfit2, 2).Value = 45
    Else
        tfa = 1.35
        tfb = 0.35
        Cells(19 + sftfit2, 2).Value = 90
    End If
    
    Cells(17 + sftfit2, 2).Value = tfa     ' CLAM2 BL3.2Ua: 1.35, VersaProbe II: 180.2540
    Cells(18 + sftfit2, 2).Value = tfb     ' CLAM2 BL3.2Ua: 0.35, VersaProbe II: 0.3480
    Cells(2, 100).Value = "dblMin"
    Cells(3, 100).Value = "dblMax"
    Cells(4, 100).Value = "numXPSFactors"
    Cells(5, 100).Value = "numData"
    Cells(6, 100).Value = "startEb"
    Cells(7, 100).Value = "endEb"
    Cells(12, 100).Value = "pe/shift"
    Cells(13, 100).Value = "wf"
    Cells(14, 100).Value = "char"
    Cells(15, 100).Value = "nom. factor"
	Cells(16, 100).Value = vbNullString
    Cells(17, 100).Value = "Iteration limit"
    Cells(18, 100).Value = "Average data"
    Cells(19, 100).Value = "Ver."
	Cells(20, 100).Value = "BG type"
    Cells(8, 100).Value = "fit done"
    Cells(9, 100).Value = "#peak before"
    Cells(10, 100).Value = "Avg points"
    Cells(10, 101).Value = 10
    Cells(2, 101).Value = dblMin
    Cells(3, 101).Value = dblMax
    Cells(4, 101).Value = numXPSFactors
    Cells(5, 101).Value = numData
    Cells(6, 101).Value = startEb
    Cells(7, 101).Value = endEb
    Cells(12, 101).Value = pe
    Cells(13, 101).Value = wf
    Cells(14, 101).Value = char
    'Cells(16, 101).Value = "BG"
    Cells(17, 101).Value = 10       ' limit of iteration
    Cells(18, 101).FormulaR1C1 = "=Average(R31C2:R" & (30 + numData) & "C2)"
	Cells(20, 101).Value = "BG"
    Cells(8, 101).Value = 0         ' trigger to change the number of peaks
    Cells(2, 102).Value = "max FWHM1 limit"
    Cells(3, 102).Value = "min FWHM1 limit"
    Cells(4, 102).Value = "max FWHM2 limit"
    Cells(5, 102).Value = "min FWHM2 limit"
    Cells(6, 102).Value = "max shape limit"
    Cells(7, 102).Value = "min shape limit"
    Cells(8, 102).Value = "factor additional peaks" ' peak BE to be added with this value/#peaks
    Cells(9, 102).Value = "GL form"
    
    If IsEmpty(Cells(9, 103).Value) Then
        If WorksheetFunction.Round(Cells(12, 101).Value, 1) = 1486.6 Then
            Cells(9, 103).Value = "MultiPak"
        Else
            Cells(9, 103).Value = "Sum"
        End If
    ElseIf LCase(Cells(9, 103).Value) = "multipak" Then
        Cells(9, 103).Value = "MultiPak"
    ElseIf LCase(Cells(9, 103).Value) = "product" Then
        Cells(9, 103).Value = "Product"
    Else
        Cells(9, 103).Value = "Sum"
    End If
    
    If mid$(Cells(25 + sftfit2, 1).Value, 1, 1) = "M" Then   ' manual set
        Cells(2, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value)       ' max FWHM1 limit
        Cells(3, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / 100      ' min FWHM1 limit
        Cells(4, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value)        ' max FWHM2 limit
        Cells(5, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / 100      ' min FWHM2 limit
        Cells(6, 103).Value = 0.999       ' max shape limit
        Cells(7, 103).Value = 0.001       ' min shape limit
'        Cells(10, 101).Value = 5          ' average points for poly BG
        If numData > 50 Then
            Cells(10, 101).Value = 10          ' average points for poly BG
        Else
            Cells(10, 101).Value = Application.Ceiling(numData / 10, 1)
        End If
        Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (100)
    ElseIf Cells(15 + sftfit2, 2).Value = 1 Then   ' grating #1
        Cells(2, 103).Value = 2       ' max FWHM1 limit
        Cells(3, 103).Value = 0.1       ' min FWHM1 limit
        Cells(4, 103).Value = 2       ' max FWHM2 limit
        Cells(5, 103).Value = 0.1       ' min FWHM2 limit
        Cells(6, 103).Value = 0.999       ' max shape limit
        Cells(7, 103).Value = 0.001       ' min shape limit
        Cells(10, 101).Value = 20          ' average points for poly BG
        If strl(1) = "Pe" Then             ' additional BE step
            Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (20)
            'Cells(2, 103).Value = 1       ' max FWHM1 limit
        Else
            Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (100)
        End If
    Else        ' grating #2, 3, G = 0 for AlKa XPS
        Cells(2, 103).Value = 10       ' max FWHM1 limit
        Cells(3, 103).Value = 0.5       ' min FWHM1 limit
        Cells(4, 103).Value = 10       ' max FWHM2 limit
        Cells(5, 103).Value = 0.5       ' min FWHM2 limit
        Cells(6, 103).Value = 0.999       ' max shape limit
        Cells(7, 103).Value = 0.001       ' min shape limit
'        Cells(10, 101).Value = 10          ' average points for poly BG
        If numData > 50 Then
            Cells(10, 101).Value = 10          ' average points for poly BG
        Else
            Cells(10, 101).Value = Application.Ceiling(numData / 10, 1)
        End If
        
        If strl(1) = "Pe" Then             ' additional BE step
            Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (4)
        Else
            Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (10)
        End If
    End If
    
    Cells(13, 102).Value = "+1.5 IQR"
    Cells(12, 102).Value = "Q3"
    Cells(16, 102).Value = "median"
    Cells(15, 102).Value = "Q1"
    Cells(14, 102).Value = "-1.5 IQR"
    Cells(17, 102).Value = "avg"
    Cells(18, 102).Value = "min"
    Cells(19, 102).Value = "max"
    Cells(13, 103).FormulaR1C1 = "=R12C103 + (R12C103 - R15C103)*1.5"
    Cells(12, 103).FormulaR1C1 = "=PERCENTILE(R" & (21 + sftfit) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (20 + sftfit + Cells(5, 101).Value) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ",0.75) "
    Cells(16, 103).FormulaR1C1 = "=PERCENTILE(R" & (21 + sftfit) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (20 + sftfit + Cells(5, 101).Value) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ",0.5) "
    Cells(15, 103).FormulaR1C1 = "=PERCENTILE(R" & (21 + sftfit) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (20 + sftfit + Cells(5, 101).Value) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ",0.25) "
    Cells(14, 103).FormulaR1C1 = "=R15C103 - (R12C103 - R15C103)*1.5"
    Cells(17, 103).FormulaR1C1 = "=Average(R" & (21 + sftfit) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (20 + sftfit + Cells(5, 101).Value) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ") "
    'Cells(18, 103).FormulaR1C1 = "=Stdevp(R" & (21 + sftfit) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (20 + sftfit + Cells(5, 101).Value) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ") "
    Cells(18, 103).FormulaR1C1 = "=PERCENTILE(R" & (21 + sftfit) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (20 + sftfit + Cells(5, 101).Value) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ",0) "
    Cells(19, 103).FormulaR1C1 = "=PERCENTILE(R" & (21 + sftfit) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (20 + sftfit + Cells(5, 101).Value) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ",1) "
    Cells(11, 103).Value = DateValue(Now) & ", " & TimeValue(Now)
    Cells(11, 104).Value = DateValue(Now) + 1
    Cells(12, 104).FormulaR1C1 = "=R12C103"
    Cells(13, 104).FormulaR1C1 = "=R13C103"
    Cells(14, 104).FormulaR1C1 = "=R14C103"
    Cells(15, 104).FormulaR1C1 = "=R15C103"
    Cells(16, 104).FormulaR1C1 = "=R16C103"
    [A2:A5].Interior.Color = RGB(156, 204, 101)    '43
    [B2:B5].Interior.Color = RGB(197, 225, 165)    '35
    Range(Cells(6 + sftfit2, 1), Cells(6 + sftfit2, 1)).Interior.Color = RGB(102, 187, 106) 'RGB(128, 203, 196) ' RGB(156, 204, 101)    '43
    Range(Cells(6 + sftfit2, 2), Cells(6 + sftfit2, 2)).Interior.Color = RGB(165, 214, 167) 'RGB(178, 223, 219) ' RGB(197, 225, 165)    '35
    Range(Cells(8 + sftfit2, 1), Cells(9 + sftfit2, 1)).Interior.Color = RGB(255, 160, 0) '45
    Range(Cells(8 + sftfit2, 2), Cells(9 + sftfit2, 2)).Interior.Color = RGB(255, 202, 40)     '44
    Range(Cells(11 + sftfit2, 1), Cells(12 + sftfit2, 1)).Interior.Color = RGB(186, 104, 200)  '39
    Range(Cells(11 + sftfit2, 2), Cells(12 + sftfit2, 2)).Interior.Color = RGB(225, 190, 231)  '38
    Range(Cells(14 + sftfit2, 1), Cells(19 + sftfit2, 1)).Interior.Color = RGB(161, 136, 127)   '16
    Range(Cells(14 + sftfit2, 2), Cells(19 + sftfit2, 2)).Interior.Color = RGB(188, 170, 164)  '15
    Range(Cells(1, 4), Cells(15 + sftfit2, 4)).Interior.Color = RGB(77, 208, 225)  '33
    Range(Cells(15 + sftfit2 + 1, 4), Cells(15 + sftfit2 + 4, 4)).Interior.Color = RGB(176, 190, 197)
    Range(Cells(1, 5), Cells(15 + sftfit2, 5)).Interior.Color = RGB(178, 235, 242) '34
    Range(Cells(15 + sftfit2 + 1, 5), Cells(15 + sftfit2 + 4, 5)).Interior.Color = RGB(207, 216, 220)
End Sub

Sub descriptInitialFit()
    Cells(20 + sftfit, 3).Value = "BG"
    Cells(20 + sftfit, 4).Value = "In-BG"
    Cells(1, 4).Value = "Name"
    Cells(2, 4).Value = "BE"
    Cells(3, 4).Value = "KE"
    Cells(4, 4).Value = "FWHM1"
    Cells(5, 4).Value = "FWHM2"
    Cells(6, 4).Value = "Amplitude"
    Cells(7, 4).Value = "Shape"
    
    If sftfit2 >= 5 Then
        Cells(8, 4).Value = "Option a"
        Cells(9, 4).Value = "Option b"
        Cells(10, 4).Value = "Option c"
        Cells(11, 4).Value = "Form"
        Cells(12, 4).Value = "beta"
    End If
    
    Cells(16 + sftfit2, 4).Value = "T.I. Area"
    Cells(17 + sftfit2, 4).Value = "S.I. Area"
    Cells(18 + sftfit2, 4).Value = "N.I. Area"
    Cells(19 + sftfit2, 4).Value = "Corr. RSF"
    Cells(7 + sftfit2, 1).Value = "Peak Fit"
    Cells(8 + sftfit2, 4).Value = "Amp+BG"
    Cells(9 + sftfit2, 4).Value = "RSF"
    Cells(10 + sftfit2, 4).Value = "P. Area"
    Cells(11 + sftfit2, 4).Value = "S. Area"
    Cells(12 + sftfit2, 4).Value = "N. Area"
    Cells(13 + sftfit2, 4).Value = "Asym"
    Cells(14 + sftfit2, 4).Value = "Amp. rat."
    Cells(15 + sftfit2, 4).Value = "BE diff."
    Cells(11, 103).Value = DateValue(Now) & ", " & TimeValue(Now)   ' time stamp
    Cells(13, 103).FormulaR1C1 = "=R12C103 + (R12C103 - R15C103)*1.5"
    Cells(12, 103).FormulaR1C1 = "=PERCENTILE(R" & (startR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (endR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ",0.75) "
    Cells(16, 103).FormulaR1C1 = "=PERCENTILE(R" & (startR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (endR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ",0.5) "
    Cells(15, 103).FormulaR1C1 = "=PERCENTILE(R" & (startR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (endR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ",0.25) "
    Cells(14, 103).FormulaR1C1 = "=R15C103 - (R12C103 - R15C103)*1.5"
    Cells(17, 103).FormulaR1C1 = "=Average(R" & (startR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (endR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ") "
    Cells(18, 103).FormulaR1C1 = "=PERCENTILE(R" & (startR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (endR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ",0) "
    Cells(19, 103).FormulaR1C1 = "=PERCENTILE(R" & (startR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ":R" & (endR) & "C" & (7 + Cells(8 + sftfit2, 2).Value) & ",1) "
    Cells(21 + sftfit + numData, 4) = IntegrationTrapezoid(Range(Cells(21 + sftfit, 1), Cells(20 + sftfit + numData, 1)), Range(Cells(21 + sftfit, 4), Cells(20 + sftfit + numData, 4)))
    
    For q = 1 To j
        Cells(16 + sftfit2, 4 + q) = IntegrationTrapezoid(Range(Cells(21 + sftfit, 1), Cells(20 + sftfit + numData, 1)), Range(Cells(21 + sftfit, 4 + q), Cells(20 + sftfit + numData, 4 + q)))
        Cells(17 + sftfit2, 4 + q).FormulaR1C1 = "= R" & (16 + sftfit2) & "C / (R" & (9 + sftfit2) & "C)"
        Cells(18 + sftfit2, 4 + q).FormulaR1C1 = "= R" & (16 + sftfit2) & "C / R" & (19 + sftfit2) & "C"
        
        If strl(1) = "Pe" Or strl(1) = "Po" Then
            Cells(19 + sftfit2, 4 + q).Value = 1        ' CorrRSF
        Else
            Cells(19 + sftfit2, 4 + q).FormulaR1C1 = "= (R15C101 * (1 - (0.25 * R" & (7 + sftfit2) & "C)*(3 * (cos(3.14*R24C2/180))^2 - 1)) * R" & (9 + sftfit2) & "C * ((R3C)^(R" & (16 + sftfit2) & "C2)) * R" & (14 + sftfit2) & "C2 * (((R" & (17 + sftfit2) & "C2^2)/((R" & (17 + sftfit2) & "C2^2)+((R3C)/(R" & (14 + sftfit2) & "C2))^2))^R" & (18 + sftfit2) & "C2))"
        End If
    Next
    
    Cells(21 + sftfit + numData, 5 + j) = IntegrationTrapezoid(Range(Cells(21 + sftfit, 1), Cells(20 + sftfit + numData, 1)), Range(Cells(21 + sftfit, 5 + j), Cells(20 + sftfit + numData, 5 + j)))
    Range(Cells(11, 104), Cells(16, 104)).ClearContents
    If ActiveSheet.ChartObjects.Count = 2 Then GoTo SkipBarPlot
    
    ActiveSheet.ChartObjects(3).Activate
    With ActiveSheet.ChartObjects(3)
        With .Chart.Axes(xlValue, xlPrimary)
            .MinimumScale = ActiveSheet.ChartObjects(2).Chart.Axes(xlValue, xlSecondary).MinimumScale
            .MaximumScale = ActiveSheet.ChartObjects(2).Chart.Axes(xlValue, xlSecondary).MaximumScale
        End With
    End With
    
SkipBarPlot:
End Sub

Sub ShirleyBG()
    Cells(1, 1).Value = "Shirley"
    Cells(1, 2).Value = "BG"
    Cells(1, 3).Value = vbNullString
    Cells(2, 1).Value = "Tolerance"
    Cells(3, 1).Value = "Initial A"
    Cells(4, 1).Value = "Final A"
    Cells(5, 1).Value = "Iteration"
    Cells(20, 101).Value = "Shirley"
    Cells(20, 102).Value = Cells(1, 2).Value
    Cells(20, 103).Value = vbNullString
    
    If Cells(8, 101).Value = 0 Then 'Or Cells(9, 101).Value > 0 Then
        Cells(2, 2).Value = 0.000001
        If Cells(3, 2).Value > 0.1 Or Cells(3, 2).Value <= 0.0000001 Then Cells(3, 2).Value = 0.001
    ElseIf Cells(3, 2).Value >= 0.1 Or Cells(3, 2).Value <= 0.000001 Then
        Cells(3, 2).Value = 0.001
    End If

    Cells(4, 2).Value = Cells(3, 2).Value
    Cells(startR, 98).FormulaR1C1 = "= (2 * RC1 - (R" & startR & "C1 + R" & endR & "C1))/(R" & endR & "C1 - R" & startR & "C1)" ' CT
    Range(Cells(startR, 98), Cells(endR, 98)).FillDown

    If Cells(20 + sftfit, 2).Value = "Ab" Then ' for PE
        If Cells(startR, 1).Value = Cells(21 + sftfit, 1).Value Then
            Cells(startR - 1, 3).FormulaR1C1 = "=AVERAGE(R[1]C2:R[" & (ns) & "]C2)"
            Cells(startR, 3).FormulaR1C1 = "=AVERAGE(R[1]C2:R[" & (ns) & "]C2)"
            Cells(startR - 1, 3).Value = Cells(startR - 1, 3).Value
            Cells(startR, 3).Value = Cells(startR, 3).Value
        ElseIf Cells(startR + Int(ns / 2), 1).Value > Cells(11 + sftfit2, 2).Value And Cells(startR, 1).Value <= Cells(11 + sftfit2, 2).Value Then
            Cells(startR - 1, 3).FormulaR1C1 = "=AVERAGE(RC2:R[" & (ns - 1) & "]C2)"
            Cells(startR, 3).FormulaR1C1 = "=AVERAGE(RC2:R[" & (ns - 1) & "]C2)"
        ElseIf Cells(startR + Int(ns / 2), 1).Value <= Cells(11 + sftfit2, 2).Value Then
            Cells(startR - 1, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (Int(ns / 2)) & "]C2:R[" & (Int(ns / 2) + 1) & "]C2)"
            Cells(startR, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (Int(ns / 2)) & "]C2:R[" & (Int(ns / 2) + 1) & "]C2)"
        End If
        
        Cells(startR, 99).FormulaR1C1 = "= ABS(RC2 - R[-1]C3)"  ' CU
        Cells(startR, 99).Value = Cells(startR, 99).Value
        For k = startR + 1 To endR Step 1
            Cells(k, 99).FormulaR1C1 = "= ABS(R[-1]C2 - R[-1]C3)"
            Cells(k, 3).FormulaR1C1 = "=R" & (startR) & "C + R4C2 * SUM(R[-1]C99:R" & (startR) & "C99)"
        Next
    Else        ' for BE
        If Cells(endR, 1).Value = Cells(Cells(5, 101).Value + 20 + sftfit, 1).Value Then
            Cells(endR + 1, 3).FormulaR1C1 = "=AVERAGE(R[-1]C2:R[" & (-ns) & "]C2)"
            Cells(endR, 3).FormulaR1C1 = "=AVERAGE(R[-1]C2:R[" & (-ns) & "]C2)"
            Cells(endR + 1, 3).Value = Cells(endR + 1, 3).Value
            Cells(endR, 3).Value = Cells(endR, 3).Value
        ElseIf Cells(Cells(5, 101).Value + 20 + sftfit - Int(ns / 2), 1).Value > Cells(11 + sftfit2, 2).Value And Cells(Cells(5, 101).Value + 20 + sftfit, 1).Value <= Cells(11 + sftfit2, 2).Value Then
            Cells(endR + 1, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (ns - (Cells(5, 101).Value + 20 + sftfit - endR)) & "]C2:R[" & (Cells(5, 101).Value + 20 + sftfit - endR - 1) & "]C2)"
            Cells(endR, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (ns - (Cells(5, 101).Value + 20 + sftfit - endR)) & "]C2:R[" & (Cells(5, 101).Value + 20 + sftfit - endR - 1) & "]C2)"
            Cells(endR + 1, 3).Value = Cells(endR + 1, 3).Value
            Cells(endR, 3).Value = Cells(endR, 3).Value
        ElseIf Cells(Cells(5, 101).Value + 20 + sftfit - Int(ns / 2), 1).Value <= Cells(11 + sftfit2, 2).Value Then
            Cells(endR + 1, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (Int(ns / 2) - 1) & "]C2:R[" & (Int(ns / 2)) & "]C2)"
            Cells(endR, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (Int(ns / 2) - 1) & "]C2:R[" & (Int(ns / 2)) & "]C2)"
            Cells(endR + 1, 3).Value = Cells(endR + 1, 3).Value
            Cells(endR, 3).Value = Cells(endR, 3).Value
        End If
        
        Cells(endR, 99).FormulaR1C1 = "= Abs(RC2 - R[1]C3)"

        For k = endR - 1 To startR Step -1
            Cells(k, 99).FormulaR1C1 = "= Abs(R[1]C2 - R[1]C3)"
            Cells(k, 3).FormulaR1C1 = "=R" & (endR + 1) & "C + R4C2 * SUM(R[1]C99:R" & (endR) & "C99)"
        Next
    End If

    Cells(startR, 100).FormulaR1C1 = "=((RC2 - RC3)^2)/(abs(RC3))" ' CV     ' added abs to solve sonvergence if negative data
    Range(Cells(startR, 100), Cells(endR, 100)).FillDown
    If ns <= 0 Then
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=AVERAGE(R" & startR & "C100:R" & endR & "C100)"
    Else
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    End If
    
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Cells(4, 2)
    SolverAdd CellRef:=Cells(4, 2), Relation:=1, FormulaText:=1 ' max
    SolverAdd CellRef:=Cells(4, 2), Relation:=3, FormulaText:=-1 ' min
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
End Sub

Sub VictoreenBG()
    Cells(1, 1).Value = "Victoreen"
    Cells(1, 2).Value = "BG"
    Cells(1, 3).Value = vbNullString
    Cells(2, 1).Value = "a0: offset"
    Cells(3, 1).Value = "a1: slope"
    Cells(4, 1).Value = "a2: 2nd poly"
    Cells(5, 1).Value = "a3: 3rd poly"
    Cells(6, 1).Value = "a4: 4th poly"
    Cells(7, 1).Value = "Edge"
    Cells(8, 1).Value = "Pre-edge"
    Cells(9, 1).Value = "Post-edge"
    Cells(20, 101).Value = "Victoreen"
    Cells(20, 102).Value = Cells(1, 3).Value
    Cells(20, 103).Value = vbNullString
    
    For k = 2 To 6
        If Cells(k, 2).Font.Bold = "True" Then
        End If
    Next
    
    If Cells(8, 101).Value = 0 Then
        Cells(8, 2).Value = Cells(11 + sftfit2, 2).Value + (Cells(12 + sftfit2, 2).Value - Cells(11 + sftfit2, 2).Value) / 20
        Cells(9, 2).Value = Cells(12 + sftfit2, 2).Value - (Cells(12 + sftfit2, 2).Value - Cells(11 + sftfit2, 2).Value) / 20
        Cells(2, 2).Value = dblMin
        Cells(3, 2).Value = ((dblMax - dblMin) / (Cells(12 + sftfit2, 2).Value - Cells(11 + sftfit2, 2).Value)) / 2
        Cells(4, 2).Value = 0
        Cells(5, 2).Value = 0
        Cells(6, 2).Value = 0
        Cells(5, 2).Font.Bold = "True"
        Cells(6, 2).Font.Bold = "True"
    End If
    
    Cells(startR, 98).FormulaR1C1 = "= RC1 - R8C2"
    Range(Cells(startR, 98), Cells(endR, 98)).FillDown
    Cells(startR, 99).FormulaR1C1 = "= (2 * (RC1-R8C2) - (R" & startR & "C1 + R" & endR & "C1 -2*R8C2))/(R" & endR & "C1 - R" & startR & "C1)" ' PE
    Range(Cells(startR, 99), Cells(endR, 99)).FillDown
    Cells(startR, 3).FormulaR1C1 = "= R2C2 + (R3C2 * RC98) + (R4C2 * (RC98^2)) + (R5C2 * (RC98^3)) + (R6C2 * (RC98^4))"
    Range(Cells(startR, 3), Cells(endR, 3)).FillDown
    Cells(startR, 100).FormulaR1C1 = "=((RC2 - RC3)^2)/(abs(RC3))" ' CV
    Range(Cells(startR, 100), Cells(endR, 100)).FillDown
    
    If Cells(8, 2).Value = vbNullString Then    ' the same as polynoial BG
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=AVERAGE(R" & startR & "C100:R" & endR & "C100)"
        Cells(8, 1).Value = "No edge"
    ElseIf Cells(8, 2).Value < Cells(12 + sftfit2, 2).Value And Cells(8, 2).Value > Cells(11 + sftfit2, 2).Value Then
        If Cells(20 + sftfit, 2).Value = "Ab" Then ' for PE
            iRow = startR + CInt(Abs(Cells(8, 2).Value - Cells(11 + sftfit2, 2).Value) / Abs(Cells(startR + 1, 1).Value - Cells(startR, 1).Value))
            Cells(6 + sftfit2, 2).FormulaR1C1 = "=AVERAGE(R" & startR & "C100:R" & iRow & "C100)"
        Else
            iRow = endR - CInt(Abs(Cells(8, 2).Value - Cells(11 + sftfit2, 2).Value) / Abs(Cells(startR + 1, 1).Value - Cells(startR, 1).Value))
            Cells(6 + sftfit2, 2).FormulaR1C1 = "=AVERAGE(R" & iRow & "C100:R" & endR & "C100)"
        End If
    Else
        Cells(8, 1).Value = "Both ends"
        If ns <= 0 Then
            Cells(6 + sftfit2, 2).FormulaR1C1 = "=AVERAGE(R" & startR & "C100:R" & endR & "C100)"
        Else
            Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
        End If
    End If
        
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(6, 2))
    SolverAdd CellRef:=Range(Cells(3, 2), Cells(4, 2)), Relation:=3, FormulaText:=0
    SolverAdd CellRef:=Cells(3, 2), Relation:=1, FormulaText:=1 ' max
    SolverAdd CellRef:=Range(Cells(4, 2), Cells(6, 2)), Relation:=1, FormulaText:=1 ' max
    SolverAdd CellRef:=Range(Cells(4, 2), Cells(6, 2)), Relation:=3, FormulaText:=-1 ' min
    
    For k = 2 To 6
        If Cells(k, 2).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
        End If
    Next

    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
    [A2:A6].Interior.Color = RGB(156, 204, 101)    '43
    [B2:B6].Interior.Color = RGB(197, 225, 165)    '35
    [A7:A9].Interior.Color = RGB(159, 168, 218)
    [B7:B9].Interior.Color = RGB(197, 202, 233)
End Sub

Sub PolynominalShirleyBG()
    Cells(1, 1).Value = "Polynominal"
    Cells(1, 2).Value = "Shirley"
    Cells(1, 3).Value = "ABG"   ' active only
    Cells(2, 1).Value = "Tolerance"
    Cells(3, 1).Value = "Initial A"
    Cells(4, 1).Value = "Final A"
    Cells(5, 1).Value = "Iteration"
    Cells(6, 1).Value = "Ratio S:P"
    Cells(7, 1).Value = "0th poly"
    Cells(8, 1).Value = "1st poly"
    Cells(9, 1).Value = "2nd poly"
    Cells(10, 1).Value = "3rd poly"
    Cells(20, 101).Value = "Polynominal"
    Cells(20, 102).Value = "Shirley"
    Cells(20, 103).Value = Cells(1, 3).Value
    
    If Cells(8, 101).Value = 0 Then 'Or Cells(9, 101).Value > 0 Then
        Cells(2, 2).Value = 0.000001
        If Cells(3, 2).Value > 0.1 Or Cells(3, 2).Value <= 0.0000001 Then Cells(3, 2).Value = 0.001
    ElseIf Cells(3, 2).Value >= 0.1 Or Cells(3, 2).Value <= 0.000001 Then
        Cells(3, 2).Value = 0.001
    End If
    
    Cells(4, 2).Value = Cells(3, 2).Value
    
    For k = 2 To 10
        If Cells(k, 2).Font.Bold = "True" Then
        ElseIf k = 2 Then
            If Cells(8, 101).Value = 0 Then Cells(2, 2).Value = 0.000001
            If Cells(3, 2).Value > 0.1 Or Cells(3, 2).Value <= 0.0000001 Then Cells(3, 2).Value = 0.001
        ElseIf k = 3 Then
            If Cells(3, 2).Value >= 0.1 Or Cells(3, 2).Value <= 0.000001 Then Cells(3, 2).Value = 0.001
        ElseIf k = 4 Then
            Cells(4, 2).Value = Cells(3, 2).Value
        ElseIf k = 6 Then
            If Cells(8, 101).Value = 0 Or IsEmpty(Cells(k, 2).Value) Then Cells(6, 2).Value = 0.5
        ElseIf k = 7 Then
            If Cells(8, 101).Value = 0 Or IsEmpty(Cells(k, 2).Value) Then Cells(7, 2).Value = Cells(2, 101).Value
        ElseIf k = 8 Then
            If Abs(Cells(k, 2).Value) > Abs(Cells(7, 2).Value) Or IsEmpty(Cells(k, 2).Value) Then Cells(k, 2).Value = 0
        ElseIf k = 9 Then
            If Abs(Cells(k, 2).Value) > Abs(Cells(7, 2).Value) Or IsEmpty(Cells(k, 2).Value) Then Cells(k, 2).Value = 0
        ElseIf k = 10 Then
            If Abs(Cells(k, 2).Value) > Abs(Cells(7, 2).Value) Or IsEmpty(Cells(k, 2).Value) Then Cells(k, 2).Value = 0
        End If
    Next
    
    Cells(startR, 98).FormulaR1C1 = "= (2 * RC1 - (R" & startR & "C1 + R" & endR & "C1))/(R" & endR & "C1 - R" & startR & "C1)"
    Range(Cells(startR, 98), Cells(endR, 98)).FillDown
    
    If Cells(20 + sftfit, 2).Value = "Ab" Then ' for PE
        If Cells(startR, 1).Value = Cells(21 + sftfit, 1).Value Then
            Cells(startR - 1, 3).FormulaR1C1 = "=AVERAGE(R[1]C2:R[" & (ns) & "]C2)"
            Cells(startR, 3).FormulaR1C1 = "=AVERAGE(R[1]C2:R[" & (ns) & "]C2)"
            Cells(startR - 1, 3).Value = Cells(startR - 1, 3).Value
            Cells(startR, 3).Value = Cells(startR, 3).Value
        ElseIf Cells(startR + Int(ns / 2), 1).Value > Cells(11 + sftfit2, 2).Value And Cells(startR, 1).Value <= Cells(11 + sftfit2, 2).Value Then
            Cells(startR - 1, 3).FormulaR1C1 = "=AVERAGE(RC2:R[" & (ns - 1) & "]C2)"
            Cells(startR, 3).FormulaR1C1 = "=AVERAGE(RC2:R[" & (ns - 1) & "]C2)"
        ElseIf Cells(startR + Int(ns / 2), 1).Value <= Cells(11 + sftfit2, 2).Value Then
            Cells(startR - 1, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (Int(ns / 2)) & "]C2:R[" & (Int(ns / 2) + 1) & "]C2)"
            Cells(startR, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (Int(ns / 2)) & "]C2:R[" & (Int(ns / 2) + 1) & "]C2)"
        End If
        
        Cells(startR, 99).FormulaR1C1 = "= ABS(RC2 - R[-1]C3)"
        Cells(startR, 99).Value = Cells(startR, 99).Value
        
        If Cells(8, 101).Value = 0 Then
            For k = startR + 1 To endR Step 1
                Cells(k, 99).FormulaR1C1 = "= ABS(R[-1]C2 - R[-1]C3)"
                Cells(k, 3).FormulaR1C1 = "=(R6C2 *(R" & (startR) & "C + ( R4C2 * SUM(R[-1]C99:R" & (startR) & "C99)))) + ((1-R6C2) * (R7C2 + (R8C2 * RC98) + (R9C2 * (RC98)^2) + (R10C2 * (RC98)^3)))"
            Next
        End If
    Else        ' for BE
        If Cells(endR, 1).Value = Cells(Cells(5, 101).Value + 20 + sftfit, 1).Value Then
            Cells(endR + 1, 3).FormulaR1C1 = "=AVERAGE(R[-1]C2:R[" & (-ns) & "]C2)"
            Cells(endR, 3).FormulaR1C1 = "=AVERAGE(R[-1]C2:R[" & (-ns) & "]C2)"
            Cells(endR + 1, 3).Value = Cells(endR + 1, 3).Value
            Cells(endR, 3).Value = Cells(endR, 3).Value
        ElseIf Cells(Cells(5, 101).Value + 20 + sftfit - Int(ns / 2), 1).Value > Cells(11 + sftfit2, 2).Value And Cells(Cells(5, 101).Value + 20 + sftfit, 1).Value <= Cells(11 + sftfit2, 2).Value Then
            Cells(endR + 1, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (ns - (Cells(5, 101).Value + 20 + sftfit - endR)) & "]C2:R[" & (Cells(5, 101).Value + 20 + sftfit - endR - 1) & "]C2)"
            Cells(endR, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (ns - (Cells(5, 101).Value + 20 + sftfit - endR)) & "]C2:R[" & (Cells(5, 101).Value + 20 + sftfit - endR - 1) & "]C2)"
            Cells(endR + 1, 3).Value = Cells(endR + 1, 3).Value
            Cells(endR, 3).Value = Cells(endR, 3).Value
        ElseIf Cells(Cells(5, 101).Value + 20 + sftfit - Int(ns / 2), 1).Value <= Cells(11 + sftfit2, 2).Value Then
            Cells(endR + 1, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (Int(ns / 2) - 1) & "]C2:R[" & (Int(ns / 2)) & "]C2)"
            Cells(endR, 3).FormulaR1C1 = "=AVERAGE(R[" & -1 * (Int(ns / 2) - 1) & "]C2:R[" & (Int(ns / 2)) & "]C2)"
            Cells(endR + 1, 3).Value = Cells(endR + 1, 3).Value
            Cells(endR, 3).Value = Cells(endR, 3).Value
        End If
        
        Cells(endR, 99).FormulaR1C1 = "= ABS(RC2 - R[1]C3)"
        
        If Cells(8, 101).Value = 0 Then
            For k = endR - 1 To startR Step -1
                Cells(k, 99).FormulaR1C1 = "= ABS(R[1]C2 - R[1]C3)"
                'Cells(k, 3).FormulaR1C1 = "=R" & (endR + 1) & "C + (R3C2 *( R2C2 * SUM(R[1]C99:R" & (endR) & "C99))) + ((1-R3C2) * (R4C2 * RC98) + (R3C3 * (RC98)^2) + (R5C3 * (RC98)^3))"
                Cells(k, 3).FormulaR1C1 = "=(R6C2 * (R" & (endR + 1) & "C + ( R4C2 * SUM(R[1]C99:R" & (endR) & "C99)))) + ((1-R6C2) * (R7C2 +(R8C2 * RC98) + (R9C2 * (RC98)^2) + (R10C2 * (RC98)^3)))"
            Next
        End If
    End If
    
    Cells(startR, 100).FormulaR1C1 = "=((RC2 - RC3)^2)/(abs(RC3))" ' CV
    Range(Cells(startR, 100), Cells(endR, 100)).FillDown
    If ns <= 0 Then
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=AVERAGE(R" & startR & "C100:R" & endR & "C100)"
    Else
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    End If
    
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(4, 2), Cells(10, 2))
    SolverAdd CellRef:=Cells(6, 2), Relation:=1, FormulaText:=1 ' max ratio
    SolverAdd CellRef:=Cells(6, 2), Relation:=3, FormulaText:=0 ' min ratio
    SolverAdd CellRef:=Cells(4, 2), Relation:=1, FormulaText:=1 ' max A
    SolverAdd CellRef:=Cells(4, 2), Relation:=3, FormulaText:=0 ' min A
    SolverAdd CellRef:=Cells(5, 2), Relation:=2, FormulaText:=1 '
    
    For k = 4 To 10
        If Cells(k, 2).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
        ElseIf k = 8 Then
            SolverAdd CellRef:=Cells(8, 2), Relation:=1, FormulaText:=Cells(7, 2) / 1
        ElseIf k = 9 Then
            SolverAdd CellRef:=Cells(9, 2), Relation:=1, FormulaText:=Cells(8, 2) / 1
        ElseIf k = 10 Then
            SolverAdd CellRef:=Cells(10, 2), Relation:=1, FormulaText:=Cells(9, 2) / 1
        End If
    Next
    
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
    Range(Cells(6, 1), Cells(10, 1)).Interior.Color = RGB(156, 204, 101)   '43
    Range(Cells(6, 2), Cells(10, 2)).Interior.Color = RGB(197, 225, 165)   '35
End Sub

Sub TangentArcBG()
    Cells(1, 1).Value = "Arctan"
    Cells(1, 2).Value = "BG"
    Cells(1, 3).Value = vbNullString
    Cells(20, 101).Value = "Arctan"
    Cells(20, 102).Value = Cells(1, 2).Value
    Cells(20, 103).Value = vbNullString
    
    For k = 2 To 7
        If Cells(k, 2).Font.Bold = "True" Then
        End If
    Next
    
    Cells(2, 1).Value = "Const. BG"
    Cells(3, 1).Value = "Step height"
    Cells(4, 1).Value = "Inflection"
    Cells(5, 1).Value = "Step width"
    Cells(6, 1).Value = "Slope"
    Cells(7, 1).Value = "ratio A:L"
    
    If Cells(8, 101).Value = 0 Then
        Cells(6, 2).Value = 0.4
        Cells(3, 2).Value = (Cells(3, 101).Value - Cells(2, 101).Value) / 2
        Cells(4, 2).Value = Cells(11 + sftfit2, 2).Value + (Cells(12 + sftfit2, 2).Value - Cells(11 + sftfit2, 2).Value) / 4
        Cells(5, 2).Value = 2
    End If
    
    Cells(startR, 3).FormulaR1C1 = "=R2C2 + (1-R7C2) * (R6C2 * (RC1 - R4C2)) + R7C2 * (R3C2 * ((0.5) + (1/3.14) * ATAN((RC1 - R4C2)/(R5C2 / 2))))"
    Range(Cells(startR, 3), Cells(endR, 3)).FillDown
    Cells(startR, 100).FormulaR1C1 = "=((RC2 - RC3)^2)/(abs(RC3))" ' CV
    Range(Cells(startR, 100), Cells(endR, 100)).FillDown
    If ns <= 0 Then
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=AVERAGE(R" & startR & "C100:R" & endR & "C100)"
    Else
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    End If

    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(7, 2))
    SolverAdd CellRef:=Cells(4, 2), Relation:=3, FormulaText:=Cells(11 + sftfit2, 2).Value        ' This is a point to control the position of inflection
    SolverAdd CellRef:=Cells(4, 2), Relation:=1, FormulaText:=Cells(12 + sftfit2, 2).Value
    SolverAdd CellRef:=Cells(5, 2), Relation:=3, FormulaText:=1 'step width minimum
    SolverAdd CellRef:=Cells(5, 2), Relation:=1, FormulaText:=(Cells(12 + sftfit2, 2).Value - Cells(11 + sftfit2, 2).Value)
    SolverAdd CellRef:=Cells(3, 2), Relation:=3, FormulaText:=0
    SolverAdd CellRef:=Cells(3, 2), Relation:=1, FormulaText:=(Cells(3, 101).Value - Cells(2, 101).Value)
    SolverAdd CellRef:=Cells(2, 2), Relation:=3, FormulaText:=0
    SolverAdd CellRef:=Cells(6, 2), Relation:=3, FormulaText:=-1
    SolverAdd CellRef:=Cells(6, 2), Relation:=1, FormulaText:=1
    SolverAdd CellRef:=Cells(7, 2), Relation:=3, FormulaText:=0
    SolverAdd CellRef:=Cells(7, 2), Relation:=1, FormulaText:=1
    
    For k = 2 To 7
        If Cells(k, 2).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
        End If
    Next

    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
    Range(Cells(6, 1), Cells(7, 1)).Interior.Color = RGB(156, 204, 101)  '43
    Range(Cells(6, 2), Cells(7, 2)).Interior.Color = RGB(197, 225, 165)  '35
End Sub

Sub PolynominalBG()
    If StrComp(strl(1), "Po", 1) = 0 Then
    Else
        For k = 2 To 5
            If Cells(k, 2).Font.Bold = "True" Then
            ElseIf Cells(8, 101).Value = 0 Then
                If k = 2 Then
                    'If Abs(Cells(2, 2).Value) > 10 Or Abs(Cells(2, 2).Value) < 1 Then Cells(2, 2).Value = 5
                    Cells(2, 2).Value = (Cells(2, 101).Value)
                ElseIf k = 3 Then
                    If Abs(Cells(3, 2).Value) > Abs(Cells(2, 2).Value) Then Cells(3, 2).Value = Cells(2, 2).Value / 2 '0.1
                ElseIf k = 4 Then
                    If Abs(Cells(4, 2).Value) > Abs(Cells(2, 2).Value) Then Cells(4, 2).Value = Cells(2, 2).Value / 5 '0.01
                ElseIf k = 5 Then
                    If Abs(Cells(5, 2).Value) > Abs(Cells(2, 2).Value) Then Cells(5, 2).Value = Cells(2, 2).Value / 10 '0.001
                End If
            End If
        Next
    End If
    
    Cells(1, 1).Value = "Polynominal"
    
    If StrComp(mid$(LCase(Cells(1, 2).Value), 1, 1), "a", 1) = 0 Then
        Cells(1, 2).Value = "ABG"
    Else
        Cells(1, 2).Value = "BG"
    End If
    
    Cells(1, 3).Value = vbNullString
    Cells(2, 1).Value = "a0"
    Cells(3, 1).Value = "a1"
    Cells(4, 1).Value = "a2"
    Cells(5, 1).Value = "a3"
    Cells(20, 101).Value = "Polynominal"
    Cells(20, 102).Value = Cells(1, 2).Value
    Cells(20, 103).Value = vbNullString
    Cells(startR, 99).FormulaR1C1 = "= (2 * RC1 - (R" & startR & "C1 + R" & endR & "C1))/(R" & endR & "C1 - R" & startR & "C1)"
    Range(Cells(startR, 99), Cells(endR, 99)).FillDown
    Cells(startR, 3).FormulaR1C1 = "=R2C2 + (R3C2 * RC99) + (R4C2 * (RC99)^2) + (R5C2 * (RC99)^3)"
    Range(Cells(startR, 3), Cells(endR, 3)).FillDown
    
    If Cells(2, 2).Value = 0 Or Cells(startR, 3).Value = 0 Then
        Cells(startR, 100).FormulaR1C1 = "=(RC2 - RC3)^2" ' CV this is the case for RC3 = 0
    Else
        Cells(startR, 100).FormulaR1C1 = "=((RC2 - RC3)^2)/(abs(RC3))" ' CV
    End If
    
    Range(Cells(startR, 100), Cells(endR, 100)).FillDown
    If ns <= 0 Then
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=AVERAGE(R" & startR & "C100:R" & endR & "C100)"
    Else
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    End If
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(5, 2))
    
    For k = 2 To 5
        If Cells(k, 2).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
        End If
    Next
    
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
End Sub

Sub TougaardBG()
    Dim pnpara As String
    
    If StrComp(mid$(Cells(3, 1).Value, 1, 6), "C (C'=", 1) = 0 And IsNumeric(mid$(Cells(3, 1).Value, 7, 2)) = True Then
        p = mid$(Cells(3, 1).Value, 7, 2)
        If p = 1 Then
            pnpara = "+1"
        ElseIf p = -1 Then
            pnpara = "-1"
        Else
            p = 1
            pnpara = "+1"
        End If
    Else
        p = 1
        pnpara = "+1"
    End If
    
    Cells(1, 1).Value = "Tougaard"
    Cells(1, 2).Value = "BG"
    Cells(1, 3).Value = vbNullString
    Cells(2, 1).Value = "B"
    Cells(3, 1).Value = "C (C'=" & pnpara & ")"
    Cells(4, 1).Value = "D"
    Cells(5, 1).Value = "Norm"
    Cells(6, 1).Value = "Offset"
    Cells(20, 101).Value = "Tougaard"
    Cells(20, 102).Value = Cells(1, 2).Value
    Cells(20, 103).Value = vbNullString
    
    For k = 2 To 5
        If Cells(k, 2).Font.Bold = "True" Then
        ElseIf k = 2 Then
            Cells(2, 2).Value = 2866    '2866 or 1840 or 736
        ElseIf k = 3 Then
            Cells(3, 2).Value = 1643    '1643 or 1000 or 400
        ElseIf k = 4 Then
            Cells(4, 2).Value = 1       ' 1 default
        ElseIf k = 5 Then
            Cells(5, 2).Value = 1
        End If
    Next
    
    stepEk = Cells(21 + sftfit, 1).Value - Cells(22 + sftfit, 1).Value
    
    If Cells(20 + sftfit, 2).Value = "Ab" Then ' for PE
        If ns < 2 Then
            Cells(startR, 3).Value = Cells(startR, 2).Value
            Cells(startR, 99).Value = Cells(startR, 3).Value
        Else
            Cells(startR, 3).FormulaR1C1 = "=Average(RC2:R[" & (ns - 1) & "]C2)"
            Cells(startR, 99).Value = Cells(startR, 3).Value
        End If
        
        For k = startR - 1 To endR - 1 Step 1
            Cells(k + 1, 99).FormulaR1C1 = "= ((RC2 * R2C2 * (" & ((startR - k + 1) * stepEk) & " ))/((R3C2 + " & p & " * (" & ((startR - k + 1) * stepEk) & ")^2)^2 + R4C2 * ((" & ((startR - k + 1) * stepEk) & " )^2)))"
            Cells(k + 1, 3).FormulaR1C1 = "=R5C2 * (R6C2 + SUM(R[-1]C99:R" & (startR) & "C99))"
        Next
    Else
        If ns < 2 Then
            Cells(endR, 3).Value = Cells(endR, 2).Value
            Cells(endR, 99).Value = Cells(endR, 3).Value
        Else
            Cells(endR, 3).FormulaR1C1 = "=Average(RC2:R[" & (ns - 1) & "]C2)"
            Cells(endR, 99).Value = Cells(startR, 3).Value
        End If
        
        For k = endR + 1 To startR + 1 Step -1
            Cells(k - 1, 99).FormulaR1C1 = "= ((RC2 * R2C2 * (" & ((endR - k + 1) * stepEk) & " ))/((R3C2 + " & p & " * (" & ((endR - k + 1) * stepEk) & ")^2)^2 + R4C2 * ((" & ((endR - k + 1) * stepEk) & " )^2)))"
            Cells(k - 1, 3).FormulaR1C1 = "=R5C2 * (R6C2 + SUM(R[1]C99:R" & (endR) & "C99))"
        Next
    End If
    
    Cells(20 + sftfit, 99).Value = "Toug"

    Cells(startR, 100).FormulaR1C1 = "=((RC2 - RC3)^2)/(abs(RC3))" ' CV
    Range(Cells(startR, 100), Cells(endR, 100)).FillDown
    
    If ns <= 0 Then
        Cells(1, 2).Value = "ABG"
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=AVERAGE(R" & startR & "C100:R" & endR & "C100)"
    Else
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    End If
    
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(6, 2))
    SolverAdd CellRef:=Range(Cells(2, 2), Cells(3, 2)), Relation:=1, FormulaText:=5000
    SolverAdd CellRef:=Range(Cells(2, 2), Cells(3, 2)), Relation:=3, FormulaText:=0.001
    SolverAdd CellRef:=Cells(4, 2), Relation:=1, FormulaText:=100
    SolverAdd CellRef:=Cells(5, 2), Relation:=1, FormulaText:=10
    SolverAdd CellRef:=Cells(6, 2), Relation:=1, FormulaText:=Cells(2, 101)
    
    For k = 2 To 6
        If Cells(k, 2).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
        End If
    Next
    
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
    
    [A2:A6].Interior.Color = RGB(156, 204, 101)    '43
    [B2:B6].Interior.Color = RGB(197, 225, 165)    '35
End Sub

Sub PolynominalTougaardBG()
    Dim pnpara As String
    
    If StrComp(mid$(Cells(3, 1).Value, 1, 6), "C (C'=", 1) = 0 And IsNumeric(mid$(Cells(3, 1).Value, 7, 2)) = True Then
        p = mid$(Cells(3, 1).Value, 7, 2)
        If p = 1 Then
            pnpara = "+1"
        ElseIf p = -1 Then
            pnpara = "-1"
        Else
            p = 1
            pnpara = "+1"
        End If
    Else
        p = 1
        pnpara = "+1"
    End If
    
    If Cells(10, 101).Value < 1 Then Cells(10, 101).Value = 1
    Cells(1, 1).Value = "Polynominal"
    Cells(1, 2).Value = "Tougaard"
    Cells(1, 3).Value = "ABG"   ' active only
    Cells(2, 1).Value = "B"
    Cells(3, 1).Value = "C (C'=" & pnpara & ")"
    Cells(4, 1).Value = "D"
    Cells(5, 1).Value = "Norm"
    Cells(6, 1).Value = "Offset"
    Cells(7, 1).Value = "ratio T:P"
    Cells(8, 1).Value = "1st poly"
    Cells(9, 1).Value = "2nd poly"
    Cells(10, 1).Value = "3rd poly"
    Cells(20, 101).Value = "Polynominal"
    Cells(20, 102).Value = "Tougaard"
    Cells(20, 103).Value = Cells(1, 3).Value
    
    For k = 2 To 10
        If Cells(8, 101).Value = 0 And k >= 7 Then
            Cells(k, 2).Font.Bold = "False"
        End If
        
        If Cells(k, 2).Font.Bold = "True" Then
        ElseIf k = 2 Then
            Cells(2, 2).Value = 2866    '2866 or 1840 or 736
        ElseIf k = 3 Then
            Cells(3, 2).Value = 1643    '1643 or 1000 or 400
        ElseIf k = 4 Then
            Cells(4, 2).Value = 1       ' 1 default
        ElseIf k = 5 Then
            Cells(5, 2).Value = 1
        ElseIf k = 6 Then
            Cells(6, 2).Value = Cells(2, 101).Value
        ElseIf k = 7 Then
            Cells(7, 2).Value = 0.9 ' ratio for Toug to Poly BG
        ElseIf k = 8 Then
            Cells(k, 2).Value = 0   ' 1st poly
        ElseIf k = 9 Then
            Cells(k, 2).Value = 0   ' 2nd poly
        ElseIf k = 10 Then
            Cells(k, 2).Value = 0   ' 3rd poly
        End If
    Next
    
    Cells(startR, 98).FormulaR1C1 = "= (2 * RC1 - (R" & startR & "C1 + R" & endR & "C1))/(R" & endR & "C1 - R" & startR & "C1)"
    Range(Cells(startR, 98), Cells(endR, 98)).FillDown
    stepEk = Cells(21 + sftfit, 1).Value - Cells(22 + sftfit, 1).Value
    
    If Cells(20 + sftfit, 2).Value = "Ab" Then ' for PE
        If ns < 2 Then
            Cells(startR, 3).Value = Cells(startR, 2).Value
            Cells(startR, 99).Value = 0
        Else
            Cells(startR, 3).FormulaR1C1 = "=Average(RC2:R[" & (ns - 1) & "]C2)"
            Cells(startR, 99).Value = 0
        End If

        For k = startR To endR - 1 Step 1
            Cells(k + 1, 99).FormulaR1C1 = "= ((RC2 * R2C2 * (" & ((startR - k + 1) * stepEk) & " ))/((R3C2 + " & p & " * (" & ((startR - k + 1) * stepEk) & ")^2)^2 + R4C2 * ((" & ((startR - k + 1) * stepEk) & " )^2)))"
            Cells(k + 1, 3).FormulaR1C1 = "=R7C2 * (R" & startR & "C + SUM(R[1]C99:R" & (startR) & "C99)) + ((1-R7C2) * (R6C2 + (R8C2 * (RC98) + (R9C2 * (RC98)^2) + (R10C2 * (RC98)^3))))"
        Next
    Else
        If ns < 2 Then
            Cells(endR, 3).Value = Cells(endR, 2).Value
            Cells(endR, 99).Value = 0
        Else
            Cells(endR, 3).FormulaR1C1 = "=Average(RC2:R[" & (ns - 1) & "]C2)"
            Cells(endR, 99).Value = 0
        End If
        
        For k = endR To startR + 1 Step -1
            Cells(k - 1, 99).FormulaR1C1 = "= ((RC2 * R2C2 * (" & ((endR - k + 1) * stepEk) & " ))/((R3C2 + " & p & " * (" & ((endR - k + 1) * stepEk) & ")^2)^2 + R4C2 * ((" & ((endR - k + 1) * stepEk) & " )^2)))"
            Cells(k - 1, 3).FormulaR1C1 = "=R7C2 * (R" & endR & "C + SUM(R[1]C99:R" & (endR) & "C99)) + ((1-R7C2) * (R6C2 + (R8C2 * (RC98) + (R9C2 * (RC98)^2) + (R10C2 * (RC98)^3))))"
        Next
    End If
    
    Cells(startR, 100).FormulaR1C1 = "=((RC2 - RC3)^2)/(abs(RC3))" ' CV
    Range(Cells(startR, 100), Cells(endR, 100)).FillDown
    If ns <= 0 Then
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=AVERAGE(R" & startR & "C100:R" & endR & "C100)"
    Else
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    End If
    
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(10, 2))
    SolverAdd CellRef:=Range(Cells(2, 2), Cells(3, 2)), Relation:=1, FormulaText:=5000
    SolverAdd CellRef:=Range(Cells(2, 2), Cells(3, 2)), Relation:=3, FormulaText:=0.001
    SolverAdd CellRef:=Cells(4, 2), Relation:=1, FormulaText:=100
    SolverAdd CellRef:=Cells(5, 2), Relation:=1, FormulaText:=10
    SolverAdd CellRef:=Cells(6, 2), Relation:=1, FormulaText:=Cells(2, 101)
    
    For k = 2 To 10
        If Cells(k, 2).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
        ElseIf k = 7 Then
            SolverAdd CellRef:=Cells(k, 2), Relation:=3, FormulaText:=0
            SolverAdd CellRef:=Cells(k, 2), Relation:=1, FormulaText:=1
        End If
    Next
    
    If Not Cells(10, 2).Font.Bold = "True" And p = 1 Then SolverAdd CellRef:=Cells(10, 2), Relation:=2, FormulaText:=0
    
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1

    [A2:A10].Interior.Color = RGB(156, 204, 101)    '43
    [B2:B10].Interior.Color = RGB(197, 225, 165)    '35
End Sub

Sub SolverSetup()
    SolverReset ' Error due to the Solver installation! Check the Solver function correctly installed.
    SolverOptions MaxTime:=100, Iterations:=32767, Precision:=0.000001, AssumeLinear _
        :=False, StepThru:=False, Estimates:=1, Derivatives:=1, SearchOption:=1, _
        IntTolerance:=5, Scaling:=False, Convergence:=(0.0001 / Cells(3, 101).Value), AssumeNonNeg:=False
End Sub

Function ShowTrial(Reason As Integer)
    MsgBox "Reason = " & Reason
    ShowTrial = 0
End Function

Sub descriptEFfit1(fcmp As Range)
    Cells(1, 1).Value = "EF"
    Cells(1, 2).Value = "fit"
    Cells(1, 3).Value = vbNullString
    Cells(2, 1).Value = "Int. DOS"
    Cells(3, 1).Value = "Slope DOS"
    Cells(4, 1).Value = "Int. BG"
    Cells(5, 1).Value = "Slope BG"
    Cells(6, 1).Value = "Poly2nd"
    Cells(7, 1).Value = "Poly3rd"
    Cells(8, 1).Value = "Norm (FD)"
    Cells(5 + sftfit2, 1).Value = "Solve FD"
    Cells(6 + sftfit2, 1).Value = "Solve GC"
    Cells(7 + sftfit2, 1).Value = "EF range"
    Cells(8 + sftfit2, 1).Value = "BE min"
    Cells(9 + sftfit2, 1).Value = "BE max"
    
    If Cells(8, 101).Value = 0 Then
        Cells(8 + sftfit2, 2).Value = -0.5
        Cells(9 + sftfit2, 2).Value = 0.5
        Cells(4, 5).Value = 300
        Cells(4, 5).Font.Bold = "True"
        Cells(2, 5).Value = 0
        Cells(6, 5).Value = 1
        Cells(8, 5).Value = 1
    ElseIf Cells(8, 101).Value = -1 Then
        Range(Cells(2, 5), Cells(8, 5)) = fcmp
    End If
    
    Cells(2, 2).Value = dblMax
    Cells(3, 2).Value = 1
    Cells(4, 2).Value = dblMin
    Cells(5, 2).Value = 0
'    Cells(6, 2).Value = 0
'    Cells(7, 2).Value = 0
    Cells(8, 2).Value = 0
    Cells(20 + sftfit, 3).Value = "FitEF (FD)"
    Cells(20 + sftfit, 4).Value = "Least fits (FD)"
    Cells(20 + sftfit, 5).Value = "Residual (FD)"
    Cells(20 + sftfit, 6).Value = "FitEF (GC)"
    Cells(20 + sftfit, 7).Value = "Least fits (GC)"
    Cells(20 + sftfit, 8).Value = "Residual (GC)"
    Cells(8, 101).Value = 0     ' 7.45: revised from "-1"
    Cells(20, 101).Value = "EF"
	Cells(20, 102).Value = "Fit"
	Cells(20, 103).Value = vbNullString
End Sub

Sub descriptEFfit2()
    Dim rng As Range, dataFit As Range, dataCheck As Range, dataIntCheck As Range
    
    j = 1
    Cells(1, 4).Value = "Name"
    Cells(2, 4).Value = "BE"
    Cells(3, 4).Value = "KE"
    Cells(4, 4).Value = "Temp"
    Cells(5, 4).Value = "Width(FD)"
    Cells(6, 4).Value = "Width(GC)"
    Cells(7, 4).Value = "Total"
    Cells(8, 4).Value = "Norm (GC)"
    Cells(1, 5).Value = "EF"
    Cells(3, 5).FormulaR1C1 = "=(" & (pe - wf - char) & " - R2C)" ' KE
    Cells(5, 5).FormulaR1C1 = "=(4.39 * R4C/11604)" ' Width     ' kT = 0.02585 eV at 300 K, 10-90% of electrons in 4.39 kT
    Cells(7, 5).FormulaR1C1 = "=sqrt(R5C5^2 + R6C5^2)" ' Width
    Cells(1, 4).Interior.Color = RGB(77, 150, 200)    '33
    Range(Cells(2, 4), Cells(8, 4)).Interior.Color = RGB(77, 208, 225)    '33
    Cells(1, 5).Interior.Color = RGB(77, 182, 172)
    Range(Cells(2, 5), Cells(8, 5)).Interior.Color = RGB(178, 235, 242)   '34
    Range(Cells(6, 1), Cells(8, 1)).Interior.Color = RGB(156, 204, 101)   '43
    Range(Cells(6, 2), Cells(8, 2)).Interior.Color = RGB(197, 225, 165)   '35
    Cells(5 + sftfit2, 1).Interior.Color = RGB(102, 187, 106) 'RGB(128, 203, 196) ' RGB(156, 204, 101)    '43
    Cells(5 + sftfit2, 2).Interior.Color = RGB(165, 214, 167) 'RGB(178, 223, 219) ' RGB(197, 225, 165)    '35
    
    Set rng = Range(Cells(startR, 1), Cells(endR, 1))
    Set dataFit = Range(Cells(p, 1), Cells(q, 1))
    
    Cells(13, 103).FormulaR1C1 = "=R12C103 + (R12C103 - R15C103)*1.5"
    Cells(12, 103).FormulaR1C1 = "=PERCENTILE(R" & (p) & "C8:R" & (q) & "C8,0.75) "
    Cells(16, 103).FormulaR1C1 = "=PERCENTILE(R" & (p) & "C8:R" & (q) & "C8,0.5) "
    Cells(15, 103).FormulaR1C1 = "=PERCENTILE(R" & (p) & "C8:R" & (q) & "C8,0.25) "
    Cells(14, 103).FormulaR1C1 = "=R15C103 - (R12C103 - R15C103)*1.5"
    Cells(17, 103).FormulaR1C1 = "=Average(R" & (p) & "C8:R" & (q) & "C8) "
    Cells(18, 103).FormulaR1C1 = "=PERCENTILE(R" & (p) & "C8:R" & (q) & "C8,0) "
    Cells(19, 103).FormulaR1C1 = "=PERCENTILE(R" & (p) & "C8:R" & (q) & "C8,1) "
    
    Range(Cells(11, 104), Cells(16, 104)).ClearContents '.Delete
    If ActiveSheet.ChartObjects.Count = 2 Then GoTo SkipBarPlotEF
    
    ActiveSheet.ChartObjects(3).Activate
    With ActiveSheet.ChartObjects(3)
        With .Chart.Axes(xlValue, xlPrimary)
            .MinimumScale = ActiveSheet.ChartObjects(2).Chart.Axes(xlValue, xlPrimary).MinimumScale
            .MaximumScale = ActiveSheet.ChartObjects(2).Chart.Axes(xlValue, xlPrimary).MaximumScale
        End With
    End With
    
SkipBarPlotEF:
    
    Set dataData = rng
    Set dataIntData = dataFit
End Sub

Sub ProfileAnalyzer()
    Dim coef0 As Integer, coef1, coef2, coef3, coef4, coef5, strShape As String
    
    strShape = Cells(11, (4 + n)).Value
    
    If IsNumeric(Cells(7, (4 + n)).Value) Then
        If Cells(7, (4 + n)).Value = 0 Then
            coef1 = 0
        ElseIf Cells(7, (4 + n)).Value = 1 Then
            coef1 = 1
        Else
            coef1 = 2
        End If
    Else
        If Cells(7, (4 + n)).Value = "Gauss" Then
            coef1 = 0
        ElseIf Cells(7, (4 + n)).Value = "Lorentz" Then
            coef1 = 1
        Else
            coef1 = 2
        End If
    End If
    
    If Cells(7, (4 + n)).Font.Italic Then
        coef2 = 1
    Else
        coef2 = 0
    End If
    
    If Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleNone Then
        coef3 = 0
    ElseIf Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleSingle Then
        coef3 = 1
    ElseIf Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleDouble Then
        coef3 = 2
    End If
    
    If LCase(Cells(9, 103).Value) = "multipak" Then
        coef4 = 1
        coef5 = 1
    ElseIf LCase(Cells(9, 103).Value) = "product" Then
        coef4 = -1
        coef5 = 0
    Else
        coef4 = 1      ' "sum"
        coef5 = 0
    End If
    
    If coef1 = 2 And (strShape = "G" Or strShape = "SGL") Then
        strShape = "SGL"
        Cells(9, 103).Value = "Sum"
        coef4 = 1      ' "sum"
        coef5 = 0
    End If
    coef0 = (1000 * coef5 + 100 * coef1 + 10 * coef2 + coef3) * coef4
    
    If strShape = "G" And coef0 <> 0 Then
        Cells(7, (4 + n)).Value = "Gauss"
        Cells(7, (4 + n)).Font.Italic = "False"
        Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleNone
    ElseIf strShape = "L" And Abs(coef0) <> 100 Then
        Cells(7, (4 + n)).Value = "Lorentz"
        Cells(7, (4 + n)).Font.Italic = "False"
        Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleNone
    ElseIf strShape = "SGL" And coef0 <> 200 Then
        If Not 0 < Cells(7, (4 + n)).Value < 1 Or IsNumeric(Cells(7, (4 + n)).Value) = False Then Cells(7, (4 + n)).Value = 0.2
        Cells(7, (4 + n)).Font.Italic = "False"
        If coef0 > 1000 Then
            Cells(9, 103).Value = "MultiPak"
        Else
            Cells(9, 103).Value = "Sum"
        End If
        Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleNone
    ElseIf strShape = "TSGL" And coef0 <> 1201 Then
        If Not 0 < Cells(7, (4 + n)).Value < 1 Or IsNumeric(Cells(7, (4 + n)).Value) = False Then Cells(7, (4 + n)).Value = 0.2
        Cells(7, (4 + n)).Font.Italic = "False"
        Cells(9, 103).Value = "MultiPak"
        Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleSingle
    ElseIf strShape = "GL" And coef0 <> 1200 Then
        If Not 0 < Cells(7, (4 + n)).Value < 1 Or IsNumeric(Cells(7, (4 + n)).Value) = False Then Cells(7, (4 + n)).Value = 0.2
        Cells(7, (4 + n)).Font.Italic = "False"
        Cells(9, 103).Value = "MultiPak"
        Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleNone
    End If
End Sub

Sub FitEquations()
    Dim rng As Range, imax As Integer, npa As Integer, pts As Points, pt As Point
    
    Set rng = Range(Cells(startR, 1), Cells(endR, 1))
    
    If Cells(15 + sftfit2, 2).Value = 1 Then    ' normalized factors for each grating by gold reference measurement
        Cells(15, 101).Value = 0.01
    ElseIf Cells(15 + sftfit2, 2).Value = 2 Then
        Cells(15, 101).Value = 0.0002
    ElseIf Cells(15 + sftfit2, 2).Value = 3 Then
        Cells(15, 101).Value = 0.0001
    Else
        Cells(15, 101).Value = 0.001    ' VersaProbe II AlKa mode normalized by 1000
    End If
    
    imax = 0    '# of iteration for asymmetric voigt fit
    npa = Cells(8 + sftfit2, 2).Value
    j = npa
    q = Cells(9, 101).Value
    Range(Cells(1, (5 + npa)), Cells(15 + sftfit2 + 4, 55)).Clear
    Range(Cells(20 + sftfit, 5), Cells((2 * numData + 22 + sftfit), 55)).ClearContents
    
    Range(Cells(1, 5), Cells(15 + sftfit2, (4 + npa))).Interior.Color = RGB(178, 235, 242) '34
    Range(Cells(15 + sftfit2 + 1, 5), Cells(15 + sftfit2 + 4, (4 + npa))).Interior.Color = RGB(207, 216, 220)
    
    If q < j Then
        If (j - q) Mod 2 = 0 And j Mod 2 = 0 Then
            For n = 1 To (j - q) Step 2
                Range(Cells(1, (4 + q + n)), Cells(9 + sftfit2, (4 + q + n + 1))).Value = Range(Cells(1, (4 + q - 1)), Cells(9 + sftfit2, (4 + q))).Value
                Range(Cells(14 + sftfit2, (4 + q + n)), Cells(15 + sftfit2, (4 + q + n + 1))).Value = Range(Cells(14 + sftfit2, (4 + q - 1)), Cells(15 + sftfit2, (4 + q))).Value
                Cells(1, (4 + q + n)).Value = Cells(1, 5).Value + "_" + CStr((4 + q + n - 5) / 2)
                Cells(1, (4 + q + n + 1)).Value = Cells(1, 6).Value + "_" + CStr((4 + q + n + 1 - 6) / 2)
                Cells(2, (4 + q + n)).Value = Cells(2, (4 + q - 1)).Value + n * (Cells(8, 103).Value / Cells(8 + sftfit2, 2).Value)
                Cells(2, (4 + q + n + 1)).Value = Cells(2, (4 + q)).Value + n * (Cells(8, 103).Value / Cells(8 + sftfit2, 2).Value)
                If Cells(4, 5).Font.Bold = True Then
                    Cells(4, (4 + q + n)).Font.Bold = True
                End If
                
                If Cells(4, 6).Font.Bold = True Then
                    Cells(4, (4 + q + n + 1)).Font.Bold = True
                End If
            Next
        Else
            For n = 1 To (j - q)
                Range(Cells(1, (4 + q + n)), Cells(9 + sftfit2, (4 + q + n))).Value = Range(Cells(1, (4 + q)), Cells(9 + sftfit2, (4 + q))).Value
                Cells(1, (4 + q + n)).Value = Cells(1, 5).Value + "s" + CStr((4 + q + n) - 5)
                Cells(2, (4 + q + n)).Value = Cells(2, (4 + q)).Value + n * (Cells(8, 103).Value / Cells(8 + sftfit2, 2).Value)
            Next
        End If
        Cells(9, 101).Value = j
    ElseIf q > j Then
        Cells(9, 101).Value = j
    End If

    For n = 1 To j
        Call ProfileAnalyzer
        
        If IsEmpty(Cells(7, (4 + n))) = True Then Cells(7, (4 + n)) = 0
        If IsNumeric(Cells(7, (4 + n))) = False Then
            If Cells(7, (4 + n)) = "Gauss" Then
                Cells(7, (4 + n)) = 0
            ElseIf Cells(7, (4 + n)) = "Lorentz" Then
                Cells(7, (4 + n)) = 1
            ElseIf Cells(7, (4 + n)) = "Voigt" Then
                Cells(7, (4 + n)) = 0.5
            Else
                Cells(7, (4 + n)) = 0
            End If
        Else
            If Cells(7, (4 + n)) < 0 Or Cells(7, (4 + n)) > 1 Then Cells(7, (4 + n)) = 0
        End If
        
        If Cells(7, (4 + n)) = 0 Then
            Cells(startR, (4 + n)).FormulaR1C1 = "=R6C * EXP(-(1/2)*((RC[" & (-3 - n) & "]-R2C)/(R4C/2.35))^2)"
            Range(Cells(startR, (4 + n)), Cells(endR, (4 + n))).FillDown
            Cells(10 + sftfit2, (4 + n)).FormulaR1C1 = "=SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)"  ' Area Gauss
            Cells(11 + sftfit2, (4 + n)).FormulaR1C1 = "=R15C / R14C" ' Area Gauss
            Cells(12 + sftfit2, (4 + n)).FormulaR1C1 = "=R15C / R24C" ' Area Gauss
            Cells(11, (4 + n)).Value = "G"
        ElseIf Cells(7, (4 + n)) = 1 Then
            Cells(startR, (4 + n)).FormulaR1C1 = "= R6C * (((R4C/2)^2)/((RC[" & (-3 - n) & "]-R2C)^2 + (R4C/2)^2))"
            Range(Cells(startR, (4 + n)), Cells(endR, (4 + n))).FillDown
            Cells(10 + sftfit2, (4 + n)).FormulaR1C1 = "=(R6C * (R4C/2) * 3.14)" ' Area Lorentz
            Cells(11 + sftfit2, (4 + n)).FormulaR1C1 = "=R15C / R14C"  ' Area Lorentz
            Cells(12 + sftfit2, (4 + n)).FormulaR1C1 = "=R15C / R24C"  ' Area Lorentz
            Cells(11, (4 + n)).Value = "L"
        ElseIf 0 < Cells(7, (4 + n)).Value < 1 And Cells(9, 103).Value = "Sum" Then    ' GL sum form: SGL
            Cells(5, (4 + n)).Value = Cells(4, (4 + n)).Value
            Cells(startR, (4 + n)).FormulaR1C1 = "=R6C * ((R7C)*((((R5C)/2)^2)/((RC[" & (-3 - n) & "]-R2C)^2 + ((R5C)/2)^2)) + (1- R7C)*(EXP(-(1/2)*((RC[" & (-3 - n) & "]-R2C)/(R4C/2.35))^2)))"
            Range(Cells(startR, (4 + n)), Cells(endR, (4 + n))).FillDown
            Cells(10 + sftfit2, (4 + n)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R5C/2) * 3.14))) "
            Cells(11 + sftfit2, (4 + n)).FormulaR1C1 = "=R15C / R14C"
            Cells(12 + sftfit2, (4 + n)).FormulaR1C1 = "=R15C / R24C"
            Cells(11, (4 + n)).Value = "SGL"
        ElseIf 0 < Cells(7, (4 + n)).Value < 1 And Cells(9, 103).Value = "MultiPak" Then    ' GL multipak form: GL and TSGL
            Cells(5, (4 + n)).Value = Cells(4, (4 + n)).Value
            If Cells(7, (4 + n)).Font.Italic = "False" And Cells(7, (4 + n)).Font.Underline = xlUnderlineStyleSingle Then   ' exponential asymmetric blend based Voigt (GL multipak)
                Cells(8, (4 + n)).Value = 0.35
                Cells(9, (4 + n)).Value = 10
                Debug.Print "non-italic underline multipak"
                For k = 1 To numData        ' ' R8C: Tail coefficient, R9C: Half Tail length at half maximum
                    If Cells((startR - 1 + k), 1).Value >= Cells(2, (4 + n)).Value Then
                        Cells((startR - 1 + k), (4 + n)).FormulaR1C1 = "=R6C * ((R7C)*((((R4C)/2)^2)/((RC[" & (-3 - n) & "]-R2C)^2 + ((R4C)/2)^2)) + (1- R7C)*(EXP(-(1/2)*((RC[" & (-3 - n) & "]-R2C)/(R4C/2.35))^2)) + (R8C * (1 - EXP(-(1/2)*((RC[" & (-3 - n) & "]-R2C)/(R4C/2.35))^2)) * exp((-6.9/R9C) * (2 * (RC[" & (-3 - n) & "] - R2C))/R4C)))"
                    Else
                        Cells((startR - 1 + k), (4 + n)).FormulaR1C1 = "=R6C * ((R7C)*((((R4C)/2)^2)/((RC[" & (-3 - n) & "]-R2C)^2 + ((R4C)/2)^2)) + (1- R7C)*(EXP(-(1/2)*((RC[" & (-3 - n) & "]-R2C)/(R4C/2.35))^2)))"
                    End If
                Next
                
                Cells(10 + sftfit2, (4 + n)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R4C/2) * 3.14))) "
                Cells(11 + sftfit2, (4 + n)).FormulaR1C1 = "=R15C / R14C"
                Cells(12 + sftfit2, (4 + n)).FormulaR1C1 = "=R15C / R24C"
                Cells(11, (4 + n)).Value = "TSGL"
            Else
                Cells(startR, (4 + n)).FormulaR1C1 = "=R6C * ((R7C)*((((R4C)/2)^2)/((RC[" & (-3 - n) & "]-R2C)^2 + ((R4C)/2)^2)) + (1- R7C)*(EXP(-(1/2)*((RC[" & (-3 - n) & "]-R2C)/(R4C/2.35))^2)))"
                Range(Cells(startR, (4 + n)), Cells(endR, (4 + n))).FillDown
                Cells(10 + sftfit2, (4 + n)).FormulaR1C1 = "=SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)"
                Cells(11 + sftfit2, (4 + n)).FormulaR1C1 = "=R15C / R14C"
                Cells(12 + sftfit2, (4 + n)).FormulaR1C1 = "=R15C / R24C"
                Cells(11, (4 + n)).Value = "GL"     ' MultiPak GL sum form with a single FWHM for G and L
            End If
        End If
        
        Cells(20 + sftfit, (4 + n)).FormulaR1C1 = "=R1C" ' Peak name
        Cells(3, (4 + n)).FormulaR1C1 = "=(R12C101 - R13C101 - R14C101 - R2C)" ' KE " & (pe - wf - char) & "
        Cells((numData + 23 + sftfit), (4 + n)).FormulaR1C1 = "=R[" & (-numData - 2) & "]C + R[" & (-numData - 2) & "]C[" & -n - 1 & "]"      ' Peak + BG
        Range(Cells((numData + 23 + sftfit), (4 + n)), Cells((2 * numData + 22 + sftfit), (4 + n))).FillDown
        Cells((numData + 22 + sftfit), (4 + n)).FormulaR1C1 = "=R1C" ' Peak name"
        Cells(8 + sftfit2, (4 + n)).FormulaR1C1 = "=R6C + " & dblMin
    Next

    Cells(startR, (5 + j)).FormulaR1C1 = "=SUM(RC[" & -j & "]:RC[-1])"      ' sum of peaks
    Range(Cells(startR, (5 + j)), Cells(endR, (5 + j))).FillDown
    Cells((numData + 23 + sftfit), (5 + j)).FormulaR1C1 = "=R[" & (-numData - 2) & "]C + R[" & (-numData - 2) & "]C[" & -j - 2 & "]"    ' Sum of Peaks + BG
    Range(Cells((numData + 23 + sftfit), (4 + n)), Cells((2 * numData + 22 + sftfit), (4 + n))).FillDown
    Cells((numData + 22 + sftfit), (4 + n)).Value = "peaks+BG"
    Cells(startR, (6 + j)).FormulaR1C1 = "=((RC2 - R[" & (2 + numData) & "]C[-1])^2)/(abs(R[" & (2 + numData) & "]C[-1]))"     ' Least fits 2, added abs to converge if the data in negative
    Range(Cells(startR, (6 + j)), Cells(endR, (6 + j))).FillDown
    Cells(9 + sftfit2, 2).FormulaR1C1 = "=(SUM(R" & (21 + sftfit) & "C" & (6 + j) & ":R" & (20 + sftfit + numData) & "C" & (6 + j) & ")) /(" & (endR - startR + 1) & ")" 'Sum of LS4
    Cells(20 + sftfit, (5 + j)).Value = "SUM fits"
    Cells(20 + sftfit, (6 + j)).Value = "Least fits"
    Cells(20 + sftfit, (7 + j)).Value = "Residual"
    Cells(startR, (7 + j)).FormulaR1C1 = "=(RC2 - R[" & (2 + numData) & "]C[-2])"    ' percentage
    Range(Cells(startR, (7 + j)), Cells(endR, (7 + j))).FillDown
    
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(3)
        .ChartType = xlXYScatterLinesNoMarkers
        .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit + numData + 2) & "C" & (5 + j) & ""
        .XValues = rng
        .Values = rng.Offset((numData + 2), (4 + j))
        .Border.ColorIndex = 41
        .Format.Line.Weight = 3
    End With
    
    For n = 1 To j
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(3 + n)
            .ChartType = xlXYScatterLinesNoMarkers
            .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C" & (4 + n) & ""
            .XValues = rng
            .Values = rng.Offset((numData + 2), (3 + n))
            .Format.Line.Weight = 1
        End With
    Next
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(4 + j)
        .ChartType = xlXYScatter
        .Name = "Peaks"
        .XValues = Range(Cells(2, 5), Cells(2, 5).Offset(0, (j - 1)))
        .Values = Range(Cells(8 + sftfit2, 5), Cells(8 + sftfit2, 5).Offset(0, (j - 1)))
        .MarkerStyle = 2
        .MarkerSize = 10
        .Format.Fill.ForeColor.RGB = RGB(255, 0, 0)
        .HasDataLabels = True
        n = 0
        Set pts = .Points
        For Each pt In pts
            n = n + 1
            With pt.DataLabel
                .Text = Range(Cells(1, 5), Cells(1, 5).Offset(0, (j - 1))).Cells(n).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 12
            End With
        Next
    End With
    
    For n = 1 To j
        Cells(1, (4 + n)).Interior.Color = ActiveChart.SeriesCollection(n + 3).Border.Color
        Cells(1, (4 + n)).Font.ColorIndex = 2
    Next
    
    If ActiveSheet.ChartObjects.Count = 1 Then Exit Sub
    
    ActiveSheet.ChartObjects(2).Activate
    k = ActiveChart.SeriesCollection.Count
    For n = k To 2 Step -1
        ActiveChart.SeriesCollection(n).Delete
    Next
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(2)
        .ChartType = xlXYScatterLinesNoMarkers
        .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C" & (5 + j) & ""
        .XValues = rng
        .Values = rng.Offset(0, (4 + j))
        .AxisGroup = xlPrimary
        .Border.ColorIndex = 33
        .Format.Line.Weight = 2
    End With
    
    For n = 1 To j
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(2 + n)
            .ChartType = xlXYScatterLinesNoMarkers
            .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C" & (4 + n) & ""
            .XValues = rng
            .Values = rng.Offset(0, (3 + n))
            .AxisGroup = xlPrimary
            .Format.Line.Weight = 1
            .Border.LineStyle = xlDashDot
            .HasDataLabels = False
        End With
    Next
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(3 + j)
        .ChartType = xlXYScatter
        .Name = "Peaks"
        .XValues = Range(Cells(2, 5), Cells(2, 5).Offset(0, (j - 1)))
        .Values = Range(Cells(6, 5), Cells(6, 5).Offset(0, (j - 1)))
        .AxisGroup = xlPrimary
        .MarkerStyle = 2
        .MarkerSize = 10
        .MarkerForegroundColor = RGB(255, 0, 0)
        .Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .HasDataLabels = True
        n = 0
        Set pts = .Points
        For Each pt In pts
            n = n + 1
            With pt.DataLabel
                .Text = Range(Cells(1, 5), Cells(1, 5).Offset(0, (j - 1))).Cells(n).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 12
            End With
        Next
    End With
    
    For n = 1 To j
        ActiveChart.SeriesCollection(n + 2).Border.Color = Cells(1, (4 + n)).Interior.Color
    Next
    
'   ----- *Residual display* -----------
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart
        With .SeriesCollection(4 + j)
            .ChartType = xlXYScatterLinesNoMarkers
            .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C" & (6 + n) & ""
            .XValues = rng
            .Values = rng.Offset(, (6 + j))
            .AxisGroup = xlSecondary
            .Border.ColorIndex = 44
            .Format.Line.Weight = 2
            .HasDataLabels = False
        End With
    End With
    
    ActiveChart.HasAxis(xlCategory, xlSecondary) = True
    With ActiveChart.Axes(xlCategory, xlSecondary)
        If StrComp(strl(1), "Pe", 1) = 0 Then
            .MinimumScale = startEb
            .MaximumScale = endEb
            .ReversePlotOrder = False
            .Crosses = xlMaximum
        Else
            .MinimumScale = endEb
            .MaximumScale = startEb
            .ReversePlotOrder = True
            .Crosses = xlMinimum
        End If
    End With

    ActiveChart.HasAxis(xlCategory, xlSecondary) = False

    With ActiveChart.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = "Residual"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .MajorGridlines.Border.LineStyle = xlDot
    End With
    
    k = ActiveChart.SeriesCollection.Count
    For n = k To 5 + j Step -1
        ActiveChart.SeriesCollection(n).Delete
    Next
    
    ActiveSheet.ChartObjects(1).Activate
End Sub

Sub numMajorUnitsCheck(ByRef startEk As Single, endEk As Single)
    If Abs(startEk - endEk) >= 500 Then
        numMajorUnit = 100
    ElseIf Abs(startEk - endEk) > 100 And Abs(startEk - endEk) < 500 Then
        numMajorUnit = 50 * windowSize
    ElseIf Abs(startEk - endEk) <= 100 And Abs(startEk - endEk) > 50 Then
        numMajorUnit = 4 * windowSize
    ElseIf Abs(startEk - endEk) <= 50 And Abs(startEk - endEk) > 20 Then
        numMajorUnit = 2 * windowSize
    ElseIf Abs(startEk - endEk) <= 20 And Abs(startEk - endEk) > 1 Then
        numMajorUnit = 1 * windowSize
    ElseIf Abs(startEk - endEk) <= 1 Then
        numMajorUnit = 0
    End If
    
    If numData >= 1500 Then numMajorUnit = 0
End Sub

Sub scalecheck()
    Dim dataIntGraph As Range, chkRef As Object
    
    If StrComp(mid$(ActiveSheet.Name, 1, 4), "Fit_", 1) = 0 Then
        Set dataIntGraph = dataKeGraph.Offset(, 1)
    Else
        Set dataIntGraph = dataKeGraph.Offset(, 2)
    End If
    
    With Application
        If strl(3) = "De" Then
            j = 1
        Else
            j = 0
        End If
        
        startEb = Cells(20 + (numData), 2 - j).Value
        endEb = Cells(20 + (numData), 2 - j).Offset(numData - 1, 0).Value

        Call numMajorUnitsCheck(startEb, endEb)
        
        If strl(3) = "Pp" Then numMajorUnit = 5 ' for RGA Qmass
    
        If strl(1) = "Pe" Or strl(3) = "De" Or strl(1) = "Po" Then
            If numMajorUnit = 0 Then
                
            ElseIf startEb < 0 Then
                startEb = .Ceiling(startEb, (-1 * numMajorUnit))
            Else
                startEb = .Floor(startEb, numMajorUnit)
            End If
        ElseIf startEb > 0 And numMajorUnit <> 0 Then
            startEb = .Ceiling(startEb, numMajorUnit)
        ElseIf numMajorUnit <> 0 Then
            startEb = .Floor(startEb, (-1 * numMajorUnit))
        End If

        If strl(1) = "Pe" Or strl(3) = "De" Or strl(1) = "Po" Then
            If numMajorUnit = 0 Then
            
            ElseIf endEb < 0 Then
                endEb = .Floor(endEb, (-1 * numMajorUnit))
            Else
                endEb = .Ceiling(endEb, numMajorUnit)
            End If
        ElseIf endEb > 0 And numMajorUnit <> 0 Then
            endEb = .Floor(endEb, numMajorUnit)
        ElseIf numMajorUnit <> 0 Then
            endEb = .Ceiling(endEb, (-1 * numMajorUnit))
        End If
        
        dblMax = .Max(dataIntGraph)
        dblMin = .Min(dataIntGraph)
        
        If strl(3) = "De" Then
            dblMax = .Max(dataBeGraph)
            dblMin = .Min(dataBeGraph)
            chkMax = .Max(dataIntGraph)
            chkMin = .Min(dataIntGraph)
            If chkMax = 0 Or chkMin = 0 Then
                strErr = "err0"
            Else
                If InStr(1, chkMax, ".") Then
                    j = Len(mid$(chkMax, 1, InStr(1, chkMax, ".", 1) - 1))
                Else
                    j = Len(chkMax)
                End If
                
                chkMax = .Ceiling(chkMax, 2 * (10 ^ (j - 1)))
                
                If InStr(1, chkMin, ".") Then
                    j = Len(mid$(chkMin, 1, InStr(1, Abs(chkMin), ".", 1) - 1))
                Else
                    j = Len(chkMin) - 1
                End If
                
                chkMin = .Ceiling(chkMin, -2 * (10 ^ (j - 1)))
            End If
        End If
    End With
End Sub

Sub Initial()
    numRun = numRun + 1
    strLabel = ""
    strAna = ""
    strCasa = ""
    strAES = ""
    strErr = ""
    strErrX = ""

    pe = 0
    off = 0
    multi = 1
    startR = 0
    endR = 0
    g = 0
    cmp = -1
    numXPSFactors = 0
    numAESFactors = 0

    ReDim Preserve highpe(0)
    ReDim strl(3)
    Debug.Print numRun
    'If numRun = 1 And backSlash = "/" Then Call requestFileAccess

    On Error Resume Next
    If ActiveWorkbook.Charts.Count > 0 Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Charts.Delete
        Application.DisplayAlerts = True
    End If
    
    If Err.Number > 0 Then
        If Err.Number = 91 Then Call debugAll       ' if no workbook is open, go to debugAll process!
        End
    End If
    
    If backSlash <> "/" Then
        If InStr(ActiveWorkbook.Name, ".txt") > 0 And InstanceCount > 1 Then
            Call ExcelRenew ' regenerate xlsx file from text opened with different insstance
        ElseIf StrComp(ActiveSheet.Name, "renew", 1) = 0 Then
            testMacro = "debug"
            ElemX = ElemD
            ActiveSheet.Name = ActiveWorkbook.Name
        End If
    End If
    
    With Application.AddIns
    For n = 1 To .Count
        If LCase(.Item(n).Name) = "solver.xlam" Then
            If Len(.Item(n).FullName) > 10 Then
                If AddIns("Solver Add-In").Installed = True Then
                    Exit For
                Else
                    MsgBox "No solver installed in Excel Add-in!" & vbCrLf & " Go to Excel Options - Add-Ins - Go Manage - Solver to be checked."
                    End
                End If
            End If
        ElseIf n = .Count And LCase(.Item(n).Name) <> "solver.xlam" Then
            MsgBox "No solver found in Excel Add-in!" & vbCrLf & " Go to Excel Options - Add-Ins - Go Manage - Solver.xlam to be browsed."
            End
        End If
    Next n
    End With
End Sub

Function InstanceCount() As Integer
    Dim objList As Object, objType As Object, strObj$
    
    strObj = "Excel.exe"
    Set objType = GetObject("winmgmts:").ExecQuery("select * from win32_process where name='" & strObj & "'")
    InstanceCount = objType.Count
' http://www.mrexcel.com/forum/excel-questions/400446-visual-basic-applications-check-if-excel-already-open.html
End Function

Sub ExcelRenew()
    Dim xlApp As Object, nxlApp As Object, wb As String, Fname As String

    Application.DisplayAlerts = False
    On Error Resume Next
    
    Set xlApp = GetObject(ActiveWorkbook.FullName).Application
    wb = mid$(xlApp.ActiveWorkbook.Name, 1, Len(xlApp.ActiveWorkbook.Name) - 4) + ".xlsx"
    Fname = xlApp.ActiveWorkbook.Path + backSlash + wb
    xlApp.ActiveSheet.Name = "renew"
    xlApp.ActiveWorkbook.SaveAs fileName:=Fname, FileFormat:=xlOpenXMLWorkbook
    xlApp.ActiveWorkbook.Close SaveChanges:=False
    xlApp.Quit
    xlApp.Visible = False
    
    Set nxlApp = GetObject(, "excel.application")
    
    If nxlApp Is Nothing Or xlApp Is nxlApp Then
        Set xlApp = Nothing
        Set nxlApp = Nothing
        Application.DisplayAlerts = True
        End
    End If
    
    With nxlApp
        .Applocation.Visible = True
        .UserControl = True
        .Workbooks.Open Fname
        .Run ("PERSONAL.XLSB!CLAM2")
    End With
    
    Set xlApp = Nothing
    Set nxlApp = Nothing
    
    Application.DisplayAlerts = True
    End
' http://www.access-programmers.co.uk/forums/showthread.php?t=253555
End Sub

Sub GetNormalize()
    Dim C1 As Variant, C2 As Variant, C3 As Variant
    Dim SourceRangeColor1 As Single
    Dim rng As Range
    Dim imax As Integer, strTest As String
    
    If Cells(1, 1).Value = "norm" Then
        strSheetAnaName = "Norm_" + strSheetDataName
    Else
        strSheetAnaName = "Diff_" + strSheetDataName
    End If
    
    strSheetGraphName = "Graph_" + strSheetDataName

    If ExistSheet(strSheetAnaName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetAnaName).Delete
        Application.DisplayAlerts = True
    End If
        
    Worksheets.Add().Name = strSheetAnaName
    Set sheetAna = Worksheets(strSheetAnaName)
    Set sheetGraph = Worksheets(strSheetGraphName)

    wb = ActiveWorkbook.Name
    sheetGraph.Activate
    
    If Cells(1, 1).Value = "norm" Or Cells(1, 1).Value = "diff" Then
        n = 1   ' means data to be generated on third set of data column
        k = 1   ' means data to be normalized on first set of data column
        
        off = Cells(9, (5 + (n * 3)))
        multi = Cells(9, (6 + (n * 3)))
        If multi = 0 Then
            multi = 1
        End If
        
        sheetGraph.Range(Cells(1, (4 + (n * 3))), Cells((2 * (numData + 10)) - 1, (6 + (n * 3)))).Clear
        Set rng = Range(Cells(11, (k + 1 + ((0) * 3))), Cells(11, (k + 1 + (0 * 3))).End(xlDown))
        numData = Application.CountA(rng)
        Set rng = Range(Cells(11, (k + 1 + ((1) * 3))), Cells(11, (k + 1 + (1 * 3))).End(xlDown))
        iCol = Application.CountA(rng)

        C1 = sheetGraph.Range(Cells(11 + numData + 9, (k + 1 + (0 * 3))), Cells(11 + (numData * 2) + 8, (k + 2 + (0 * 3))))
        C2 = sheetGraph.Range(Cells(11 + iCol + 9, (2 + (n * 3))), Cells(11 + (iCol * 2) + 8, (3 + (n * 3))))
        C3 = sheetGraph.Range(Cells(11, (1 + ((n + 1) * 3))), Cells(10 + numData, (3 + ((n + 1) * 3))))
        stepEk = Cells(7, (k + 1 + (0 * 3))).Value
        endEk = Cells(7, (k + 1 + (1 * 3))).Value

        p = 1
        For q = 1 To numData
            For j = 1 To iCol
                If C1(q, 1) > C2(j, 1) - (endEk / 2) And C1(q, 1) < C2(j, 1) + (endEk / 2) Then
                    C3(p, 1) = C1(q, 1)
                    If C2(j, 2) <> 0 Then
                        If Cells(1, 1).Value = "norm" Then
                            C3(p, 3) = C1(q, 2) / C2(j, 2) ' here is normalized
                        Else
                            C3(p, 3) = C1(q, 2) - C2(j, 2) ' difference
                        End If
                    Else
                        C3(p, 3) = "NaN"
                    End If
                    p = p + 1
                    Exit For
                End If
            Next

            If j = iCol + 1 And endEk < stepEk Then
                For j = 1 To iCol
                    If C1(q, 1) > C2(j, 1) - (stepEk / 2) And C1(q, 1) < C2(j, 1) + (stepEk / 2) Then
                        C3(p, 1) = C1(q, 1)
                        If C2(j, 2) <> 0 Then
                            If Cells(1, 1).Value = "norm" Then
                                C3(p, 3) = C1(q, 2) / C2(j, 2) ' here is normalized
                            Else
                                C3(p, 3) = C1(q, 2) - C2(j, 2) ' difference
                            End If
                        Else
                            C3(p, 3) = "NaN"
                        End If
                        p = p + 1
                        Exit For
                    End If
                Next
            End If
        Next
        
        numData = p - 1
        imax = numData + 10
        sheetGraph.Range(Cells(11, (1 + ((n + 1) * 3))), Cells(10 + numData, (3 + ((n + 1) * 3)))) = C3
        
        If LCase(Cells(10, 1).Value) = "pe" Then
            strl(1) = "Pe"
            strl(2) = "Sh"
            strl(3) = "Ab"
        Else
            strl(1) = "Ke"
            strl(2) = "Be"
            strl(3) = "In"
        End If
        
        If Cells(1, 1).Value = "norm" Then
            strTest = strSheetDataName + "_norm"
        Else
            strTest = strSheetDataName + "_diff"
        End If
        
        Cells(1, (5 + (n * 3))).Value = strTest
        Cells(8 + (imax), (5 + (n * 3))).Value = strTest
        Cells(9 + (imax), (4 + (n * 3))).Value = strl(1) + strTest
        Cells(9 + (imax), (5 + (n * 3))).Value = strl(2) + strTest
        Cells(9 + (imax), (6 + (n * 3))).Value = strl(3) + strTest
        
        If LCase(Cells(10, 1).Value) = "pe" Then
            Cells(2, ((4 + (n * 3)))).Value = UCase(strl(1)) & " shifts"
            Cells(2, ((5 + (n * 3)))).Value = 0
            Cells(2, ((6 + (n * 3)))).Value = "eV"
            Cells(5, ((4 + (n * 3)))).Value = "Start " & UCase(strl(1))
            Cells(6, ((4 + (n * 3)))).Value = "End " & UCase(strl(1))
            Cells(7, ((4 + (n * 3)))).Value = "Step " & UCase(strl(1))
        Else
            Cells(2, ((4 + (n * 3)))).Value = UCase(strl(2)) & " shifts"
            Cells(2, ((5 + (n * 3)))).Value = 0
            Cells(2, ((6 + (n * 3)))).Value = "eV"
            Cells(5, ((4 + (n * 3)))).Value = "Start " & UCase(strl(2))
            Cells(6, ((4 + (n * 3)))).Value = "End " & UCase(strl(2))
            Cells(7, ((4 + (n * 3)))).Value = "Step " & UCase(strl(2))
        End If
        
        Cells(5, ((5 + (n * 3)))).Value = Cells(11, 7).Value
        Cells(6, ((5 + (n * 3)))).Value = Cells(10 + numData, 7).Value
        Cells(7, ((5 + (n * 3)))).Value = Cells(12, 7).Value - Cells(11, 7).Value
        Range(Cells(5, 9), Cells(7, 9)) = "eV"

        Cells(9, ((4 + (n * 3)))).Value = "Offset/multp"
        Cells(9, ((5 + (n * 3)))).Value = off
        Cells(9, ((6 + (n * 3)))).Value = multi
        Cells(10, ((4 + (n * 3)))).Value = strl(1)
        Cells(10, ((5 + (n * 3)))).Value = strl(2)
        Cells(10, ((6 + (n * 3)))).Value = strl(3)
        
        Range(Cells(5, (4 + (n * 3))), Cells(7, (4 + (n * 3)))).Interior.ColorIndex = 41
        Range(Cells(5, (5 + (n * 3))), Cells(7, (6 + (n * 3)))).Interior.ColorIndex = 37
        Range(Cells(2, (4 + (n * 3))), Cells(2, (4 + (n * 3)))).Interior.ColorIndex = 3
        Range(Cells(2, (5 + (n * 3))), Cells(2, (6 + (n * 3)))).Interior.ColorIndex = 38
        Range(Cells(9, (4 + (n * 3))), Cells(9, ((4 + (n * 3))))).Interior.ColorIndex = 43
        Range(Cells(9, (5 + (n * 3))), Cells(9, ((6 + (n * 3))))).Interior.ColorIndex = 35
    
        Cells(11, (5 + (n * 3))).FormulaR1C1 = "=R2C + RC[-1]"
        Cells(10 + (imax), (5 + (n * 3))).FormulaR1C1 = "=R2C + R[-" & (imax - 1) & "]C[-1]"
        Range(Cells(11, (5 + (n * 3))), Cells((imax), (5 + (n * 3)))).FillDown
        
        Cells(10 + (imax), (4 + (n * 3))).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
        Range(Cells(10 + (imax), (4 + (n * 3))), Cells((2 * imax) - 1, (4 + (n * 3)))).FillDown
        Range(Cells(10 + (imax), (5 + (n * 3))), Cells((2 * imax) - 1, (5 + (n * 3)))).FillDown
        Cells(10 + (imax), (6 + (n * 3))).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C[-1])*R9C"
        Range(Cells(10 + (imax), (6 + (n * 3))), Cells((2 * imax) - 1, (6 + (n * 3)))).FillDown
        
        Set dataKeGraph = Range(Cells(10 + (imax), (4 + (n * 3))), Cells((2 * imax - 1), (4 + (n * 3))))
        
        ActiveSheet.ChartObjects(1).Activate
        p = ActiveChart.SeriesCollection.Count
        For j = 1 To p
            If ActiveChart.SeriesCollection(j).Name = Cells(1, 5 + (n * 3)).Value Then
                ActiveChart.SeriesCollection(j).Delete
                p = p - 1
                Exit For
            End If
        Next
        
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(p + n)
            .ChartType = xlXYScatterLinesNoMarkers
            .Name = Cells(1, 5 + (n * 3)).Value
            .XValues = dataKeGraph.Offset(0, 1)
            .Values = dataKeGraph.Offset(0, 2)
            SourceRangeColor1 = .Border.Color
        End With
        
        Range(Cells(10, (4 + (n * 3))), Cells(10, ((4 + (n * 3))))).Interior.Color = SourceRangeColor1
        Range(Cells(9 + (imax), (4 + (n * 3))), Cells(9 + (imax), ((4 + (n * 3))))).Interior.Color = SourceRangeColor1
        Range(Cells(10, (5 + (n * 3))), Cells(10, ((5 + (n * 3))))).Interior.Color = SourceRangeColor1
        Range(Cells(9 + (imax), (5 + (n * 3))), Cells(9 + (imax), ((5 + (n * 3))))).Interior.Color = SourceRangeColor1
        
        sheetGraph.Range(Cells(11 + numData + 8, (5 + (n * 3))), Cells(11 + (numData * 2) + 8, (6 + (n * 3)))).Copy
        sheetAna.Cells(1, 1 + ((n - 1) * 2)).PasteSpecial Paste:=xlValues

        If strl(1) = "Pe" Then
            sheetAna.Cells(1, 1).Value = "PE/eV"
        Else
            sheetAna.Cells(1, 1).Value = "BE/eV"
        End If
        
        sheetGraph.Activate
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    Application.CutCopyMode = False
    strErr = "skip"
End Sub

Sub CombineLegend() ' no k is used because from GetCompare Sub
    Dim spr As String, strSheetSampleName As String, strSheetTargetName As String, strSeriesName As String
    Dim sheetSample As Worksheet, sheetTarget As Worksheet, icur As Integer, kcur As Integer
    
    If mid$(Results, 1, 1) = "n" Then
        icur = CInt(mid$(Results, 2, Len(Results) - InStr(1, Results, "c"))) ' number of comp in each comp
        kcur = CInt(mid$(Results, InStr(1, Results, "c") + 1, Len(Results) - InStr(1, Results, "c")))            ' position of comp from 0
        Results = vbNullString
    Else
        icur = -1
        kcur = -1
    End If

    If Cells(40, para + 9).Value = "Ver." Then        ' check para is the same as that in the file analyzed previously
    Else
        For n = 1 To 500
            If StrComp(Cells(40, n + 9).Value, "Ver.", 1) = 0 Then Exit For
        Next
        para = n
    End If
        
    spr = ": "
    strSheetTargetName = ActiveSheet.Name
    strSheetSampleName = "samples"
    
    Set sheetTarget = Worksheets(strSheetTargetName)
    ncomp = sheetTarget.Cells(45, para + 10).Value
    
    If ncomp > 0 Then
        If ExistSheet(strSheetSampleName) = False Then
            Worksheets.Add().Name = strSheetSampleName
            Set sheetSample = Worksheets(strSheetSampleName)
            Cells(1, 1).Value = "No."
            Cells(1, 2).Value = "Name"
            Cells(1, 3).Value = "Sep."
            Cells(1, 4).Value = "File name"

            For n = 0 To ncomp
                Cells(2 + n, 1).Value = n + 1
                If n > 0 And InStr(1, sheetTarget.Cells(1, 2 + n * 3), spr) > 0 Then
                    Cells(2 + n, 2).Value = mid$(sheetTarget.Cells(1, 2 + n * 3).Value, 1, InStr(1, sheetTarget.Cells(1, 2 + n * 3).Value, spr) - 1)
                    Cells(2 + n, 4).Value = mid$(sheetTarget.Cells(1, 2 + n * 3).Value, InStr(1, sheetTarget.Cells(1, 2 + n * 3).Value, spr) + Len(spr), Len(sheetTarget.Cells(1, 2 + n * 3).Value))
                ElseIf n = 0 Then
                    sheetTarget.Activate
                    If ActiveSheet.ChartObjects.Count > 0 Then
                        ActiveSheet.ChartObjects(1).Activate
                        strSeriesName = ActiveChart.SeriesCollection(1).Name
                        If InStr(1, ActiveChart.SeriesCollection(1).Name, spr) > 0 Then
                            sheetSample.Activate
                            Cells(2 + n, 2).Value = mid$(strSeriesName, 1, InStr(1, strSeriesName, spr) - 1)
                            Cells(2 + n, 4).Value = mid$(strSeriesName, InStr(1, strSeriesName, spr) + Len(spr), Len(strSeriesName))
                        Else
                            sheetSample.Activate
                            Cells(2 + n, 2).Value = "no." & n + 1
                            Cells(2 + n, 4).Value = strSeriesName
                        End If
                    End If
                Else
                    Cells(2 + n, 2).Value = "no." & n + 1
                    Cells(2 + n, 4).Value = sheetTarget.Cells(1, 2 + n * 3).Value
                End If
                Cells(2 + n, 3).Value = spr
            Next
        Else
            Set sheetSample = Worksheets(strSheetSampleName)
            sheetSample.Activate

            If ncomp + 2 > sheetSample.UsedRange.Rows.Count Then
                For n = sheetSample.UsedRange.Rows.Count - 1 To ncomp
                    Cells(2 + n, 1).Value = n + 1
                    Cells(2 + n, 2).Value = "no." & n + 1
                    Cells(2 + n, 3).Value = spr
                    Cells(2 + n, 4).Value = sheetTarget.Cells(1, 2 + n * 3).Value
                Next
            ElseIf kcur >= 0 And kcur + 3 < sheetSample.UsedRange.Rows.Count Then
                For n = kcur + 1 To icur   ' comp in the middle until number of comp (icur) from kcur
                    Cells(2 + n, 1).Value = n + 1
                    Cells(2 + n, 2).Value = "no." & n + 1
                    Cells(2 + n, 3).Value = spr
                    Cells(2 + n, 4).Value = sheetTarget.Cells(1, 2 + n * 3).Value
                Next
            End If
        End If
        
        Set sheetSample = Worksheets(strSheetSampleName)
        sheetTarget.Activate
                
        For n = 0 To ncomp - 1
            sheetTarget.Cells(1, 5 + n * 3) = sheetSample.Cells(n + 3, 2).Value & spr & sheetSample.Cells(n + 3, 4).Value
        Next
        
        If ActiveSheet.ChartObjects.Count > 0 Then
            For n = 0 To ActiveSheet.ChartObjects.Count - 1
                ActiveSheet.ChartObjects(1 + n).Activate
                With ActiveChart.SeriesCollection(1)
                    .Name = sheetSample.Cells(2, 2).Value & spr & sheetSample.Cells(2, 4).Value
                End With
            Next
        End If
    End If
    
    sheetTarget.Activate
    If mid$(strSheetGraphName, 1, 6) = "Graph_" Then
        Cells(1, 1).Value = "Grating"
    Else
        Cells(1, 1).Value = vbNullString
    End If
End Sub

Sub descriptGConv()
    For k = 1 To (endR - startR + 1)
        Cells(startR, 100 + k).FormulaR1C1 = "=RC3 * Exp(-(1/2)*((RC1-R" & (startR + k - 1) & "C1)/(R6C5/2.35))^2)" ' CV
        Range(Cells(startR, 100 + k), Cells(endR, 100 + k)).FillDown
        Cells(startR + k - 1, 100).FormulaR1C1 = "=Sum(R" & (startR) & "C" & (100 + k) & ":R" & (endR) & "C" & (100 + k) & ")"
    Next k
End Sub

Sub debugAll()      ' multiple file analysis in sequence
    Dim be4all() As Variant, am4all() As Variant, fw4all() As Variant, wbX As String, shgX As Worksheet, shfX As Worksheet, strSheetDataNameX As String, numpeakX As Integer
    Dim Target As Variant, C1 As Variant, C2 As Variant, OpenFileName As Variant, debugMode As String, seriesnum As Integer, SourceRangeColor1 As Long, rng As Range, strNorm As String
    Dim debugcp As Integer, shf As Worksheet, strTest As String
    
    If mid$(testMacro, 1, 5) = "debug" Then
        modex = -1
        If testMacro = "debugGraph" Then
            debugMode = "debugGraph"
        ElseIf testMacro = "debugGraphn" Then
            debugMode = "debugGraphn"
        ElseIf testMacro = "debugFit" Then
            debugMode = "debugFit"
        ElseIf testMacro = "debugShift" Then
            debugMode = "debugShift"
        ElseIf testMacro = "debugPara" Then
            debugMode = "debugPara"
            modex = -2
        ElseIf testMacro = "debugCopy" Then ' fit the spectrum based on the fitted sheet
            debugMode = "debugCopy"
            modex = -3
        End If
    Else
        modex = 1
    End If
    
    strErrX = ""
    
    If modex <= -2 Then
        If backSlash = "/" Then
            OpenFileName = Select_File_Or_Files_Mac("xlsx")
        Else
            ChDrive mid$(ActiveWorkbook.Path, 1, 1)
            ChDir ActiveWorkbook.Path
            OpenFileName = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Please select a file", MultiSelect:=True)
        End If
    Else
        If backSlash = "/" Then
            OpenFileName = Select_File_Or_Files_Mac("csv")
        Else
            If modex <= -1 Then
                ChDrive mid$(ActiveWorkbook.Path, 1, 1)
                ChDir ActiveWorkbook.Path
            End If
            OpenFileName = Application.GetOpenFilename(FileFilter:="Text Files (*.txt), *.txt,MultiPak Files (*.csv), *.csv", Title:="Please select a file", MultiSelect:=True)
        End If
    End If
    
    If IsArray(OpenFileName) Then
        If UBound(OpenFileName) > 1 And backSlash = "\" Then
            ' http://www.cpearson.com/excel/SortingArrays.aspx
            ' put the array values on the worksheet
            Workbooks.Add
            
            Set rng = ActiveSheet.Range("A1").Resize(UBound(OpenFileName) - LBound(OpenFileName) + 1, 1)
            rng = Application.Transpose(OpenFileName)
            ' sort the range
            rng.Sort key1:=rng, order1:=xlAscending, MatchCase:=False
            
            ' load the worksheet values back into the array
            For q = 1 To rng.Rows.Count
                OpenFileName(q) = rng(q, 1)
            Next q
            
            ActiveWorkbook.Close SaveChanges:=False
        End If
    Else
        Exit Sub
    End If
    
    If modex <= -1 Then
        wb = ActiveWorkbook.Name
        wbX = wb
        If StrComp(mid$(strSheetDataName, 1, 5), "Norm_", 1) = 0 Then
            strNorm = "Norm_"
        Else
            strNorm = vbNullString
        End If
        strSheetDataNameX = strSheetDataName
        Set shgX = Workbooks(wbX).Sheets("Graph_" + strSheetDataNameX)
        peX = Workbooks(wb).Sheets("Graph_" + strSheetDataName).Cells(2, 2).Value
        If debugMode = "debugFit" Or debugMode = "debugShift" Then
            Set shfX = Workbooks(wbX).Sheets("Fit_" + strSheetDataNameX)
            numpeakX = Workbooks(wb).Sheets("Fit_" + strSheetDataName).Cells(8 + sftfit2, 2).Value
        ElseIf debugMode = "debugPara" Then
            Set shfX = Workbooks(wbX).Sheets("Fit_" + strSheetDataNameX)
            C1 = Workbooks(wb).Sheets("Fit_" + strSheetDataName).Range(Cells(14 + sftfit2, 1), Cells(19 + sftfit2, 2)).Value
        ElseIf debugMode = "debugCopy" Then
            Set shfX = Workbooks(wbX).Sheets("Fit_" + strSheetDataNameX)
        End If
    End If
    
    If modex = -1 Then
        ElemX = Workbooks(wbX).Sheets("Graph_" + strSheetDataName).Cells(51, para + 9).Value
    ElseIf modex <= -2 Then
    ElseIf Len(ElemX) > 0 Then
        Debug.Print ElemX, "ElemX", Len(ElemX)
    Else
        ElemX = Application.InputBox(Title:="Input atomic elements", Prompt:="Example:C,O,Co,etc ... without space!", Default:="C,O,Au", Type:=2)
    End If
    
    If modex <= -2 Then
    ElseIf modex = 1 Then
        If ElemX <> "False" Then
        Else
            Call GetOut
            Exit Sub
        End If
    End If
    
    idebug = 0
    
    For Each Target In OpenFileName
        If ActiveWorkbook Is Nothing Then
        Else
            If StrComp(Target, ActiveWorkbook.FullName, 1) = 0 Or StrComp(mid$(Target, 1, Len(Target) - 4), mid$(ActiveWorkbook.FullName, 1, Len(ActiveWorkbook.FullName) - 5), 1) = 0 Then GoTo SkipOpenDebug
        End If
        
        strTest = mid$(Target, InStrRev(Target, backSlash) + 1, Len(Target) - InStrRev(Target, backSlash))
        
        If Not WorkbookOpen(strTest) Then
            Workbooks.Open Target
        Else
            Workbooks(strTest).Activate
            strLabel = ActiveSheet.Name
        End If
        
        If modex = -2 Then
            Application.DisplayAlerts = False
            strSheetDataName = strNorm + mid$(Target, InStrRev(Target, backSlash) + 1, Len(Target) - InStrRev(Target, backSlash) - 5)
            Workbooks(ActiveWorkbook.Name).Sheets("Fit_" + strSheetDataName).Range(Cells(14 + sftfit2, 1), Cells(19 + sftfit2, 2)) = C1
            Workbooks(ActiveWorkbook.Name).Sheets("Fit_" + strSheetDataName).Cells(19 + sftfit2, 4) = "Corr. RSF"
            j = Workbooks(ActiveWorkbook.Name).Sheets("Fit_" + strSheetDataName).Cells(8 + sftfit2, 2).Value

            For q = 1 To j
                Cells(11 + sftfit2, (4 + q)).FormulaR1C1 = "=R15C / R14C"
                Cells(12 + sftfit2, (4 + q)).FormulaR1C1 = "=R15C / R24C"
                Cells(17 + sftfit2, 4 + q).FormulaR1C1 = "= R21C / R14C"
                Cells(18 + sftfit2, 4 + q).FormulaR1C1 = "= R21C / R24C"
                Cells(19 + sftfit2, 4 + q).FormulaR1C1 = "= (R15C101 * (1 - (0.25 * R12C)*(3 * (cos(3.14*R24C2/180))^2 - 1)) * R14C * ((R3C)^(R21C2)) * R19C2 * (((R22C2^2)/((R22C2^2)+((R3C)/(R19C2))^2))^R23C2))"
            Next
            Workbooks(ActiveWorkbook.Name).Close SaveChanges:=True
            Application.DisplayAlerts = True
            GoTo SkipOpenDebug
        ElseIf modex = -3 Then
            Application.DisplayAlerts = False
            strSheetDataName = strNorm + mid$(Target, InStrRev(Target, backSlash) + 1, Len(Target) - InStrRev(Target, backSlash) - 5)
            If Len(strSheetDataName) > 25 Then strSheetDataName = mid$(strSheetDataName, 1, 25)
            Debug.Print strSheetDataName
            
            Set shf = Workbooks(strTest).Sheets("Fit_" + strSheetDataName)
            shf.Activate
            
            If shf.Cells(8, 101).Value = 0 Then
                testMacro = "debug"     ' This is a trigger to run the debugAll code in sequence
                Call CLAM2              ' This is a main code. First run makes Graph, Fit, and Check sheets
                ' Code until here
    
                ' Error handling process here
                If StrComp(strErrX, "skip", 1) = 0 Then
                    Workbooks(ActiveWorkbook.Name).Close SaveChanges:=False
                    Debug.Print "strErrX"
                    Exit Sub
                End If
                'Error handling process end
            End If
            
            shfX.Activate
            debugcp = shfX.Cells(8 + sftfit2, 2).Value
            shfX.Range(Cells(1, 1), Cells(24 + sftfit2, 4 + debugcp)).Copy
            
            shf.Activate
            shf.Paste Destination:=shf.Range(Cells(1, 1), Cells(24 + sftfit2, 4 + debugcp))
            
            If ActiveSheet.ChartObjects.Count > 3 Then
                For q = ActiveSheet.ChartObjects.Count To 4 Step -1
                    ActiveSheet.ChartObjects(q).Delete      ' delete the chart copied from the source, no idea how to remove it!
                Next
            End If
            
            testMacro = "debug"     ' This is a trigger to run the debugAll code in sequence
            Call CLAM2              ' This is a main code. First run makes Graph, Fit, and Check sheets
            ' Code until here
            
            ' Error handling process here
            If StrComp(strErrX, "skip", 1) = 0 Then
                Workbooks(ActiveWorkbook.Name).Close SaveChanges:=False
                Debug.Print "strErrX"
                Exit Sub
            End If
            ' Error handling process end
            
            Workbooks(ActiveWorkbook.Name).Close SaveChanges:=True
            Application.DisplayAlerts = True
            Set shf = Nothing
            GoTo SkipOpenDebug
        End If
        
        ' 1st Code to run in each Target
        testMacro = "debug"     ' This is a trigger to run the debugAll code in sequence
        Call CLAM2              ' This is a main code. First run makes Graph, Fit, and Check sheets
        ' Code until here
        
        ' Error handling process here
        'If StrComp(strErrX, "skip", 1) = 0 Then
        '    Workbooks(ActiveWorkbook.Name).Close SaveChanges:=False
        '    Exit Sub
        'End If
        ' Error handling process end
        
        If modex = -1 Then
            testMacro = "debug"     ' This is a trigger to run the debugAll code in sequence
            sheetGraph.Activate     ' activate Graph sheet
            shgX.Activate
            
            If debugMode = "debugGraphn" Then
                Set rng = [D:D]
                numpeakX = (Application.CountA(rng) - 8) / 2
                shgX.Range(Cells(1, 4), Cells(2 * numpeakX + 19, 6)).Copy   ' C1 =
                sheetGraph.Activate
                sheetGraph.Cells(1, 4).PasteSpecial
                sheetGraph.Cells(1, 1).Value = "norm"
                sheetGraph.Cells(45, para + 10).Value = 1
                ActiveSheet.ChartObjects(1).Activate
                ActiveChart.SeriesCollection.NewSeries
                seriesnum = ActiveChart.SeriesCollection.Count
                
                With ActiveChart.SeriesCollection(seriesnum)
                    .ChartType = xlXYScatterLinesNoMarkers
                    .Name = "='" & ActiveSheet.Name & "'!R1C5"
                    .XValues = sheetGraph.Range(Cells(20 + numpeakX, 4), Cells(19 + 2 * numpeakX, 4))
                    .Values = sheetGraph.Range(Cells(20 + numpeakX, 6), Cells(19 + 2 * numpeakX, 6))
                    SourceRangeColor1 = .Border.Color
                End With
                
                With ActiveChart.Axes(xlValue)
                    .MinimumScaleIsAuto = True
                    .MaximumScaleIsAuto = True
                End With
                
                With ActiveSheet.ChartObjects(1)
                    .Top = 150
                End With
        
                Cells(1, 1).Select
            Else
                C1 = shgX.Range(Cells(1, 1), Cells(10, 3))                  ' basic parameters
                C2 = shgX.Range(Cells(46, para + 11), Cells(47, para + 11)) ' database
                sheetGraph.Activate
                sheetGraph.Range(Cells(1, 1), Cells(10, 3)) = C1
                sheetGraph.Range(Cells(46, para + 11), Cells(47, para + 11)) = C2
            End If
            
            Call CLAM2

            If StrComp(strErrX, "skip", 1) = 0 Then
                Workbooks(ActiveWorkbook.Name).Close SaveChanges:=False
                Exit Sub
            End If
            
            If debugMode = "debugGraphn" Then
                Workbooks(ActiveWorkbook.Name).Close SaveChanges:=True
                GoTo SkipOpenDebug
            ElseIf debugMode = "debugFit" Or debugMode = "debugShift" Then
                testMacro = "debug"     ' This is a trigger to run the debugAll code in sequence

                sheetFit.Activate       ' activate fit sheet for fitting with Shirley BG
                Call CLAM2

                If StrComp(strErrX, "skip", 1) = 0 Then
                    Workbooks(ActiveWorkbook.Name).Close SaveChanges:=False
                    Exit Sub
                End If
                
                shfX.Activate
                C1 = shfX.Range(Cells(1, 1), Cells(19 + sftfit2, 3))
                C2 = shfX.Range(Cells(2, 103), Cells(9, 103))
                sheetFit.Activate
                sheetFit.Range(Cells(1, 1), Cells(19 + sftfit2, 3)) = C1
                sheetFit.Range(Cells(2, 103), Cells(9, 103)) = C2
                
                testMacro = "debug"     ' This is a trigger to run the debugAll code in sequence
                Call CLAM2
                
                shfX.Activate
                If debugMode = "debugShift" And idebug = 0 Then
                    ReDim Preserve be4all(j)
                    ReDim Preserve am4all(j)
                    ReDim Preserve fw4all(j)
                    For k = 0 To j - 1
                        be4all(k) = Cells(2, k + 5).Value   ' record first fitted BEs for second fit in the first file
                        am4all(k) = Cells(6, k + 5).Value
                        fw4all(k) = Cells(4, k + 5).Value
                    Next
                End If
                
                C1 = shfX.Range(Cells(1, 5), Cells(15 + sftfit2 + 4, numpeakX + 4))
                sheetFit.Activate
                sheetFit.Range(Cells(1, 5), Cells(15 + sftfit2 + 4, numpeakX + 4)) = C1
                shfX.Activate
                shfX.Range(Cells(1, 5), Cells(15 + sftfit2 + 4, numpeakX + 4)).Copy
                sheetFit.Activate
                sheetFit.Cells(1, 5).PasteSpecial (xlPasteFormats)
                Application.CutCopyMode = False
                
                testMacro = "debug"     ' This is a trigger to run the debugAll code in sequence
                
                If debugMode = "debugShift" And idebug >= 0 Then
                    For k = 0 To j - 1
                        Cells(2, 5 + k).Value = be4all(k)             ' BE: 4f7/2, 5/2
                        Cells(6, 5 + k).Value = am4all(k)             ' AM: 4f7/2, 5/2
                        Cells(4, 5 + k).Value = fw4all(k)             ' FW: 4f7/2, 5/2
                    Next
                End If
                
                Call CLAM2
                
                If debugMode = "debugShift" Then
                    For k = 0 To j - 1
                        be4all(k) = Cells(2, k + 5).Value       ' record BEs for fitting in the next openning file
                        am4all(k) = Cells(6, k + 5).Value
                        fw4all(k) = Cells(4, k + 5).Value
                    Next
                End If
                
                If StrComp(strErrX, "skip", 1) = 0 Then
                    Workbooks(ActiveWorkbook.Name).Close SaveChanges:=False
                    Exit Sub
                End If
            End If
        End If
        
        On Error GoTo SkipOpenDebug
        Workbooks(ActiveWorkbook.Name).Close SaveChanges:=False
SkipOpenDebug:
        idebug = idebug + 1
    Next Target
End Sub

Sub UDsamples()         ' user defined database examples
    Sheets("Sheet1").Name = "XPS"
    Sheets("Sheet2").Name = "AES"
    Sheets("Sheet3").Name = "Notes"
    
    Dim xpsdb() As String, aesdb() As String
    xpsdb = Split("Element Orbit BE(eV) ASF C 1s 284.6 1 O 1s 532 2.93 Au 4f5/2 87.6 7.54 Au 4f7/2 84 9.58", " ")
    aesdb = Split("Element Auger KE(eV) RSF C KLL 266 0.6 O KLL 506 0.96", " ")
    
    Sheets("XPS").Activate
    k = 0
    For n = 1 To UBound(xpsdb) / 4
        For j = 1 To 4
            Cells(n, j).Value = xpsdb(k)
            k = k + 1
        Next
    Next

    Sheets("AES").Activate
    k = 0
    For n = 1 To UBound(aesdb) / 4
        For j = 1 To 4
            Cells(n, j).Value = aesdb(k)
            k = k + 1
        Next
    Next
End Sub


' "Ctrl+Q" is a set of VBA codes based on Windows Excel 2007 for
' soft x-ray XPS/XAS data analysis working with a bunch of database files
'
' Copyright (C) 2012 - 2016 Hideki NAKAJIMA
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program. If not, see <http://www.gnu.org/licenses/>.




Function WorkbookOpen(WorkBookName As String) As Boolean
' returns TRUE if the workbook is open
    WorkbookOpen = False
    On Error GoTo WorkBookNotOpen
    If Len(Application.Workbooks(WorkBookName).Name) > 0 Then
        WorkbookOpen = True
        Exit Function
    End If
WorkBookNotOpen:
End Function

Function ExistSheet(sheetName) As Boolean
    Dim r As Integer, cnt As Integer
    
    cnt = Sheets.Count
    ExistSheet = False
    For r = 1 To cnt
        If Sheets(r).Name = sheetName Then
            ExistSheet = True
            Exit For
        End If
    Next
End Function

Function IntegrationTrapezoid(KnownXs As Variant, KnownYs As Variant) As Variant
    'Calculates the area under a curve using the trapezoidal rule.
    'KnownXs and KnownYs are the known (x,y) points of the curve.
    'By Christos Samaras : http://www.myengineeringworld.net
    Dim n As Integer, rng As Range
    
    If Not TypeName(KnownXs) = "Range" Then    'Check if the X values are range.
        IntegrationTrapezoid = "Xs range is not valid"
        Exit Function
    End If
    
    If Not TypeName(KnownYs) = "Range" Then    'Check if the Y values are range.
        IntegrationTrapezoid = "Ys range is not valid"
        Exit Function
    End If
    
    IntegrationTrapezoid = 0
    
    For Each rng In KnownYs.Cells
        If IsNumeric(rng.Value) = False Then
            rng.Value = vbNullString
        End If
    Next
    
    For n = 1 To KnownXs.Rows.Count - 1
        IntegrationTrapezoid = IntegrationTrapezoid + Abs(0.5 * (KnownXs.Cells(n + 1, 1) _
        - KnownXs.Cells(n, 1)) * (KnownYs.Cells(n, 1) + KnownYs.Cells(n + 1, 1)))
    Next n
End Function

Sub SolverInstall1()
    On Error Resume Next
    Dim wb As Workbook, SolverPath As String
    
    Set wb = ActiveWorkbook ' Set a Reference to the workbook that will hold Solver
    SolverPath = Application.LibraryPath & "\SOLVER\SOLVER.XLAM"
    
    With AddIns("Solver Add-In")
        .Installed = False
        .Installed = True
    End With
    
    'Solver itself has 'focus' at this point.
    'Make sure you point to the correct Workbook for Solver
    wb.VBProject.References.AddFromFile SolverPath
    ' http://www.pcreview.co.uk/threads/vba-code-to-add-a-reference-to-solver.973572/
End Sub

Sub SolverInstall2()
    Dim wb As Workbook  '// Dana DeLouis
    
    On Error Resume Next
    Set wb = ActiveWorkbook ' Set a Reference to the workbook that will hold Solver
    
    With wb.VBProject.References
        .Remove.Item ("SOLVER")
    End With
    
    With AddIns("Solver Add-In")
        .Installed = False
        .Installed = True
        wb.VBProject.References.AddFromFile .FullName
    End With
    
    Application.Run "Solver.xlam!Solver.Solver2.Auto_open"    ' initialize Solver
End Sub

Sub requestFileAccess()
    Dim fileAccessGranted As Boolean, filePermissionCandidates, p As Integer, strt As String
    'Create an array with file paths for the permissions that are needed.
    Workbooks.Open fileName:=direc & "file_list.xlsx"
    numGrant = Cells(5, 5).End(xlDown).Row - 4
    'filePermissionCandidates = Application.Transpose(Range(Cells(5, 5), Cells(numGrant + 4, 5)).Value)
    strt = vbNullString
    For p = 1 To numGrant
        strt = strt & "," & Cells(4 + p, 5).Value
    Next p
    filePermissionCandidates = Array(strt)
    'Request access from user.
    fileAccessGranted = GrantAccessToMultipleFiles(filePermissionCandidates)
    ActiveWorkbook.Close SaveChanges:=False
    'Returns true if access is granted; otherwise, false.
    'MsgBox "Access granted on " & numGrant & " files.", vbInformation
End Sub

Function Select_File_Or_Files_Mac(ext As String) As Variant
    'Select files in Mac Excel with the format that you want
    'Working in Mac Excel 2011 and 2016
    'Ron de Bruin, 20 March 2016
    Dim MyPath As String, MyScript As String, MyFiles As String, MySplit As Variant
    Dim Fname As String, OneFile As Boolean, FileFormat As String

    'In this example you can only select xlsx files
    'See my webpage how to use other and more formats.
    If ext = "xlsx" Then
        FileFormat = "{""org.openxmlformats.spreadsheetml.sheet""}"
    ElseIf ext = "csv" Then
        FileFormat = "{""public.plain-text"","" public.comma-separated-values-text""}"
    End If

    ' Set to True if you only want to be able to select one file
    ' And to False to be able to select one or more files
    OneFile = False

    On Error Resume Next
    MyPath = MacScript("return (path to desktop folder) as String")
    'Or use A full path with as separator the :
    'MyPath = "HarddriveName:Users:<UserName>:Desktop:YourFolder:"

    'Building the applescript string, do not change this
    'This is Mac Excel 2016
    If OneFile = True Then
        MyScript = _
            "set theFile to (choose file of type" & _
            " " & FileFormat & " " & _
            "with prompt ""Please select a file"" default location alias """ & _
            MyPath & """ without multiple selections allowed) as string" & vbNewLine & _
            "return posix path of theFile"
    Else
        MyScript = _
            "set theFiles to (choose file of type" & _
            " " & FileFormat & " " & _
            "with prompt ""Please select a file or files"" default location alias """ & _
            MyPath & """ with multiple selections allowed)" & vbNewLine & _
            "set thePOSIXFiles to {}" & vbNewLine & _
            "repeat with aFile in theFiles" & vbNewLine & _
            "set end of thePOSIXFiles to POSIX path of aFile" & vbNewLine & _
            "end repeat" & vbNewLine & _
            "set {TID, text item delimiters} to {text item delimiters, ASCII character 10}" & vbNewLine & _
            "set thePOSIXFiles to thePOSIXFiles as text" & vbNewLine & _
            "set text item delimiters to TID" & vbNewLine & _
            "return thePOSIXFiles"
    End If

    MyFiles = MacScript(MyScript)
    On Error GoTo 0

    'If you select one or more files MyFiles is not empty
    'We can do things with the file paths now like I show you below
    If MyFiles <> "" Then
        Select_File_Or_Files_Mac = Split(MyFiles, Chr(10))
    End If
End Function



