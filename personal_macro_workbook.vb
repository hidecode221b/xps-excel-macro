Option Explicit
    
    Dim FSO As Object, flag As Boolean
    Dim iniTime As Date, finTime As Date, startTime As Date, endTime As Date, checkTime As Date, checkDate As Date, TimeC1 As Date, TimeC2 As Date
    Dim Fname As Variant, en As Variant, C As Variant, A As Variant, b As Variant, D As Variant, tmp As Variant, highpe() As Variant
    Dim OpenFileName As Variant, Target As Variant, Record As Variant, U As Variant, ratio() As Variant, bediff() As Variant
    Dim j As Integer, k As Integer, q As Integer, p As Integer, iRow As Integer, iCol As Integer, imax As Integer, startR As Integer, endR As Integer
    Dim numMajorUnit As Integer, ns As Integer, modex As Integer, g As Integer, Gnum As Integer, cae As Integer, NumSheets As Integer, ncomp As Integer
    Dim scanNum As Integer, scanNumR As Integer, numXPSFactors As Integer, numAESFactors As Integer, numChemFactors As Integer, fileNum As Integer
    Dim para As Integer, oriXPSFactors As Integer, numscancheck As Integer, graphexist As Integer, idebug As Integer, spacer As Integer
    Dim sftfit As Integer, sftfit2 As Integer
    Dim sh As String, wb As String, ver As String, verchk As String, wbc As String, wbp As String, NoCheck As String, strhighpe As String
    Dim strSheetDataName As String, strSheetGraphName As String, strSheetCheckName As String, strSheetFitName As String, strSheetAnaName As String
    Dim strSheetXPSFactors As String, strSheetAESFactors As String, strSheetPICFactors As String, strSheetChemFactors As String, Results As String
    Dim strTest As String, strLabel As String, strCpa As String, ElemD As String, Elem As String, strscanNum As String, strscanNumR As String
    Dim strSheetAvgName As String, strList As String, strCasa As String, strAES As String, strErr As String, strSheetCheckName2 As String
    Dim ElemX As String, strErrX As String, TimeCheck As String, asf As String
    Dim str1 As String, str2 As String, str3 As String, str4 As String, strAna, direc As String, testMacro As String
    Dim sheetData As Worksheet, sheetGraph As Worksheet, sheetCheck As Worksheet, sheetFit As Worksheet, sheetAvg As Worksheet, sheetCheck2 As Worksheet
    Dim sheetXPSFactors As Worksheet, sheetAESFactors As Worksheet, sheetPICFactors As Worksheet, sheetChemFactors As Worksheet, sheetAna As Worksheet
    Dim dataData As Range, dataKeData As Range, dataIntData As Range, dataBGraph As Range, dataKGraph, dataKeGraph As Range, dataBeGraph As Range
    Dim dataIntGraph As Range, dataCheck As Range, dataKeCheck As Range, dataIntCheck As Range, dataToFit As Range, dataFit As Range, rng As Range
    Dim mySeries As Series, pts As Points, pt As Point, myChartOBJ As ChartObject
    Dim pe As Single, wf As Single, char As Single, off As Single, multi As Single, nomfac As Single, windowSize As Single, startEk As Single
    Dim endEk As Single, startEb As Single, endEb As Single, stepEk As Single, peX As Single, dblMax As Single, dblMin As Single, dblMax1 As Single
    Dim dblMin1 As Single, chkMax As Single, chkMin As Single, gamma As Single, lambda As Single, dblMax2 As Single, dblMin2 As Single
    Dim chkMax2 As Single, chkMin2 As Single, windowRatio As Single, maxXPSFactor As Single, maxAESFactor As Single
    Dim a0 As Single, a1 As Single, a2 As Single, a3 As Single, qe As Single, trans As Single, fitLimit As Single, mfp As Single
    Dim i As Long, numData As Long, numDataN As Long, iniRow As Long, endRow As Long, totalDataPoints As Long
    Dim SourceRangeColor1 As Long, SourceRangeColor2 As Long
    
Sub CLAM2()
    ver = "8.00p"                             ' Version of this code.
    'direc = "E:\DATA\hideki\XPS\"            ' a  directory location of database (this is for PC with SSD storage.)
    direc = "D:\DATA\hideki\XPS\"            ' this is for PC with HDD storage.
    'direc = "C:\Users\Public\Data\"         ' this is for BOOTCAMP on MacBookAir.
    
    windowSize = 1.7          ' 1 for large, 2 for small display, and so on. Larger number, smaller graph plot.
    windowRatio = 4 / 3     ' window width / height, "2/1" for eyes or "4/3" for ppt
    NoCheck = "OFF"         ' "OFF" for normal, "ON" for no checking function to reduce CPU load. "Obb" for fluence analysis.
    ElemD = "C,O"           ' Default elements to be shown up in the element analysis.
    TimeCheck = "No"        ' "yes" to display the progress time, "No" only iteration results in fitting, numeric value to suppress any display.
    
    a0 = -0.00044463        ' Undulator parameters for harmonics or
    a1 = 1.0975             ' B vs gap equation
    a2 = -0.02624           ' B = A0 + A1 * Exp(A2 * gap)
    gamma = 1.2             ' An electron energy: GeV
    lambda = 6              ' A magnetic period: cm
    fitLimit = 250          ' Maximum fit range: eV
    
    mfp = 0.6               ' Inelastic mean free path formula: E^(mfp), and mfp can be from 0.5 to 0.9.
    
    qe = 0.063162           ' Averaged quantum efficiency of gold evaluated in Ip and photodiode measurements (beam size 1mm^2 to be assumed)
                            ' qe = 0.063162 for PE:39.5 eV; qe = 0.041824 for PE:60.0 eV  (units: # of electron per photon)
                            ' qe = 0.018186 for PE:80.0 eV; qe = 0.020104 for PE:120.0 eV (20130514)
    trans = 0.65            ' Gold mesh transmission to evaluate flux based on Ip: 65%
    
    para = 100              ' position of parameters in the graph sheet with higher version of 6.56.
                            ' the limit of compared spectra depends on (para/3).
    spacer = 4              ' spacer between data tables for each parameter in FitRatioAnalysis, but it should be more than 3
    sftfit = 10             ' 10
    sftfit2 = 5             ' 5
    
    Call SheetNameAnalysis
    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    
    Call TargetDataAnalysis
End Sub

Sub SheetNameAnalysis()
    If mid$(direc, Len(direc), 1) <> "\" Then direc = direc & "\"
    direc = Replace(direc, "/", "\")
    direc = Replace(direc, "*", "")
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If FSO.FolderExists(direc) Then
        If Len(Dir(direc + "UD.xlsx")) = 0 Then
            Application.DisplayAlerts = False
            Workbooks.Add
            Call UDsamples
            ActiveWorkbook.SaveAs Filename:=direc & "UD.xlsx", FileFormat:=51
            ActiveWorkbook.Close
            Application.DisplayAlerts = True
        End If
    Else
        If InStr(1, ActiveSheet.Name, "Fit_") > 0 Then
        'TimeCheck = MsgBox("Database Not Found in " + direc + "!", vbExclamation, "Database error")
            TimeCheck = MsgBox("Database Not Found in " + direc + "!" + vbCrLf + "Would you like to continue?", 4, "Database error")
            If TimeCheck = 6 Then
                testMacro = "debug"
                ElemX = ""
            
                'On Error GoTo DeadInTheWater1
                tmp = Split(direc, "\")
                For q = 1 To UBound(tmp) - 1
                    tmp(q) = tmp(q - 1) & "\" & tmp(q)
                    Debug.Print tmp(q)
                    FSO.CreateFolder tmp(q)
                Next q
                
                Workbooks.Add
                Call UDsamples
                ActiveWorkbook.SaveAs Filename:=direc & "UD.xlsx", FileFormat:=51
                ActiveWorkbook.Close
            Else
                End 'Call GetOut
DeadInTheWater1:
                MsgBox "A folder could not be created in the following path: " & direc & "." & vbCrLf & "Create directory manually and try again."
                End
            End If
        Else
            TimeCheck = MsgBox("Database Not Found in " + direc + "!" + vbCrLf + "Would you like to continue and create directory?", 4, "Database error")
            If TimeCheck = 6 Then
                'On Error GoTo DeadInTheWater2
                tmp = Split(direc, "\")
                For q = 1 To UBound(tmp) - 1
                    tmp(q) = tmp(q - 1) & "\" & tmp(q)
                    Debug.Print tmp(q)
                    FSO.CreateFolder tmp(q)
                Next q
                
                Workbooks.Add
                Call UDsamples
                ActiveWorkbook.SaveAs Filename:=direc & "UD.xlsx", FileFormat:=51
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
    
    If StrComp(testMacro, "debug", 1) = 0 Then
        TimeCheck = 0
    End If
    
    Call Initial
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic    ' revised for Office 2010
    
    graphexist = 0
    sh = ActiveSheet.Name
    
    If InStr(1, sh, "Graph_") > 0 Then
        strSheetDataName = mid$(sh, 7, (Len(sh) - 6))
        graphexist = 1       ' i for trigger for Graph sheet
        
        If IsEmpty(Cells(1, 2).Value) = False Then
            If IsNumeric(Cells(1, 2).Value) = True Then
                Gnum = Cells(1, 2).Value
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
                If Cells(2, 2).Value > 1500 Then
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
                strhighpe = Cells(2, 3).Value   ' Higher order/ghost light effects
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
        strscanNum = Cells(8, 2).Value
        strscanNumR = Cells(8, 2).Value
        
        If Cells(40, para + 9).Value = "Ver." Then
        Else
            For q = 1 To 500
                If StrComp(Cells(40, q + 9).Value, "Ver.", 1) = 0 Then Exit For
            Next
            para = q
        End If
        
        strCasa = Cells(46, para + 11).Value
        strAES = Cells(47, para + 11).Value
        
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
            Call ExportCmp
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf StrComp(LCase(Cells(1, 1).Value), "norm", 1) = 0 Then
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
        strCpa = Cells(1, (4 + (3 * k))).Value
        Call Gcheck
        If StrComp(strAna, "ana", 1) = 0 And StrComp(TimeCheck, "yes", 1) = 0 Then TimeCheck = vbNullString
    ElseIf InStr(1, sh, "Check_") > 0 Then
        strSheetDataName = mid$(sh, 7, (Len(sh) - 6))
        If StrComp(LCase(Cells(1, 1).Value), "exp", 1) = 0 Or StrComp(LCase(Cells(1, 1).Value), "exr", 1) = 0 Then
            strSheetAnaName = "Eck_" + strSheetDataName
            strSheetCheckName = "Check_" + strSheetDataName
            Call ExportChk
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
    ElseIf InStr(1, sh, "Photo_") > 0 Then
        strSheetDataName = mid$(sh, 7, (Len(sh) - 6))
        If StrComp(LCase(Cells(1, 1).Value), "exp", 1) = 0 Or StrComp(LCase(Cells(1, 1).Value), "exr", 1) = 0 Then
            strSheetAnaName = "Eck_" + strSheetDataName
            strSheetCheckName = "Photo_" + strSheetDataName
            Call ExportChk
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
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
        Else
            strSheetAnaName = "Exc_" + strSheetDataName
            strSheetGraphName = "Cmp_" + strSheetDataName
            ncomp = Range(Cells(10, 1), Cells(10, 1).End(xlToRight)).Columns.Count / 3
            Call ExportCmp
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
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
            g = Cells(8 + sftfit2, 2).Value
            C = Workbooks(wb).Sheets("Fit_" + strSheetDataName).Range(Cells(1, 5), Cells(12 + sftfit2, 4 + g))
            b = Workbooks(wb).Sheets("Fit_" + strSheetDataName).Range(Cells(1, 1), Cells(1, 3))
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
                strTest = "Do fit range"
            Else
                strTest = "Do fit"
            End If
            Call FitCurve
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
    ElseIf InStr(1, sh, "Ana_") > 0 Then
        strSheetDataName = mid$(sh, 5, (Len(sh) - 4))
        wb = ActiveWorkbook.Name
        If StrComp(Cells(1, para + 1).Value, "Parameters", 1) = 0 Then
        Else
            For i = 1 To 500
                If Cells(1, i).Value = "Parameters" Then
                    Exit For
                ElseIf i = 500 Then
                    MsgBox "Ana sheet has no parameters to be compared."
                    End
                End If
            Next
            para = i
        End If
        Call FitRatioAnalysis
        
        Application.CutCopyMode = False
        End
            
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    Else
        If InStr(ActiveWorkbook.Name, ".") < 1 Then
            TimeCheck = MsgBox("Save the file with the extention: xlsx and try it again!", vbExclamation)
            End 'Call GetOut
        Else
            strTest = mid$(ActiveWorkbook.Name, 1, InStrRev(ActiveWorkbook.Name, ".") - 1)
            strTest = mid$(strTest, 1, 25)
        End If
        
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
        
        strCasa = "User Defined"
        strAES = "User Defined"
    End If
    
    strSheetGraphName = "Graph_" + strSheetDataName
    strSheetCheckName = "Check_" + strSheetDataName
    strSheetFitName = "Fit_" + strSheetDataName
    
    If Not ExistSheet(strSheetDataName) Then
        TimeCheck = MsgBox("Data sheet " & strSheetDataName & " is not found.", vbExclamation)
        End
    End If
    Set sheetData = Worksheets(strSheetDataName)
    Worksheets(strSheetDataName).Activate
    
    If StrComp(mid$(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") + 1, 3), "txt", 1) = 0 Or StrComp(mid$(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") + 1, 3), "csv", 1) = 0 Then
        wb = mid$(ActiveWorkbook.Name, 1, InStrRev(ActiveWorkbook.Name, ".") - 1) + ".xlsx"
    
        Application.DisplayAlerts = False
        
        If Len(ActiveWorkbook.Path) < 2 Then
            Application.Dialogs(xlDialogSaveAs).Show
        Else
            On Error GoTo Error1
            ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path + "\" + wb, FileFormat:=51
        End If
        Application.DisplayAlerts = True
    ElseIf StrComp(mid$(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") + 1, 4), "xlsx", 1) = 0 Then
        wb = mid$(ActiveWorkbook.Name, 1, InStrRev(ActiveWorkbook.Name, ".") - 1) + ".xlsx"
    ElseIf StrComp(mid$(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") + 1, 3), "xls", 1) = 0 Then
        wb = mid$(ActiveWorkbook.Name, 1, InStrRev(ActiveWorkbook.Name, ".") - 1) + ".xlsx"
    Else
Error2:
        Application.Dialogs(xlDialogSaveAs).Show
        wb = ActiveWorkbook.Name
    End If
    Err.Clear
    Exit Sub
Error1:
    Err.Clear
    wb = mid$(ActiveWorkbook.Name, 1, InStr(ActiveWorkbook.Name, ".") - 1) + "_bk.xlsx"
    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path + "\" + wb, FileFormat:=51
End Sub

Sub GetAutoScale()
    Dim numDataT As Integer
    Dim iniRow1 As Single
    Dim iniRow2 As Single
    Dim endRow1 As Single
    Dim endRow2 As Single
    ' "autop" to run the previous auto command
    If StrComp(LCase(Cells(1, 1).Value), "autop", 1) = 0 And IsEmpty(Cells(40, para + 11).Value) = False Then Cells(1, 1).Value = Cells(40, para + 11).Value
    k = 0
    For i = 0 To ncomp
        Set rng = Range(Cells(11, (3 + (i * 3))), Cells(11, (3 + (i * 3))).End(xlDown))
        numDataT = Application.CountA(rng)
        If Len(Cells(1, 1).Value) > 4 Then
            If StrComp(mid$(Cells(1, 1).Value, 5, 1), "(", 1) = 0 And StrComp(mid$(Cells(1, 1).Value, Len(Cells(1, 1)), 1), ")", 1) = 0 Then
                If IsNumeric(mid$(Cells(1, 1).Value, 6, InStr(6, Cells(1, 1), ",", 1) - 6)) And IsNumeric(mid$(Cells(1, 1).Value, InStr(6, Cells(1, 1), ",", 1) + 1, Len(Cells(1, 1)) - InStr(6, Cells(1, 1), ",", 1) - 1)) Then
                    p = Application.Floor(mid$(Cells(1, 1).Value, 6, InStr(6, Cells(1, 1), ",", 1) - 6), 1)
                    q = Application.Ceiling(mid$(Cells(1, 1).Value, InStr(6, Cells(1, 1), ",", 1) + 1, Len(Cells(1, 1)) - InStr(6, Cells(1, 1), ",", 1) - 1), 1)
                    
                    If p >= 1 And q > p Then
                    Else
                        p = 1
                        q = 10
                    End If
                Else
                    p = 1
                    q = 10
                End If
                
                Set rng = Range(Cells(11 + numDataT - q, (3 + (i * 3))), Cells(11 + numDataT - p, (3 + (i * 3))))
                Set dataData = Range(Cells(10 + p, (3 + (i * 3))), Cells(10 + q, (3 + (i * 3))))
                
                If Application.WorksheetFunction.Average(dataData) > Application.WorksheetFunction.Average(rng) Then  ' PES mode
                    Cells(9, 3 * i + 2).Value = Application.WorksheetFunction.Average(rng)
                    Cells(9, 3 * i + 3).Value = 1 / Abs(Application.WorksheetFunction.Average(dataData) - Cells(9, 3 * i + 2).Value)
                Else ' XAS mode
                    Cells(9, 3 * i + 2).Value = Application.WorksheetFunction.Average(dataData)
                    Cells(9, 3 * i + 3).Value = 1 / Abs(Application.WorksheetFunction.Average(rng) - Cells(9, 3 * i + 2).Value)
                End If
            ElseIf StrComp(mid$(Cells(1, 1).Value, 5, 1), "[", 1) = 0 And StrComp(mid$(Cells(1, 1).Value, Len(Cells(1, 1)), 1), "]", 1) = 0 Then
                If IsNumeric(mid$(Cells(1, 1).Value, 6, InStr(6, Cells(1, 1), ":", 1) - 6)) And IsNumeric(mid$(Cells(1, 1).Value, InStr(6, Cells(1, 1), ",", 1) + 1, Len(Cells(1, 1)) - InStr(InStr(6, Cells(1, 1), ",", 1) + 1, Cells(1, 1), ":", 1) - 1)) Then
                    stepEk = Abs(Cells(7, 3 * i + 2).Value)
                    If mid$(Cells(1, 1).Value, 6, InStr(6, Cells(1, 1), ":", 1) - 6) < 0 Then
                        iniRow1 = Application.Floor(mid$(Cells(1, 1).Value, 6, InStr(6, Cells(1, 1), ":", 1) - 6), -1 * stepEk)
                    Else
                        iniRow1 = Application.Floor(mid$(Cells(1, 1).Value, 6, InStr(6, Cells(1, 1), ":", 1) - 6), stepEk)
                    End If
                    If mid$(Cells(1, 1).Value, InStr(6, Cells(1, 1), ",", 1) + 1, Len(Cells(1, 1)) - InStr(InStr(6, Cells(1, 1), ",", 1) + 1, Cells(1, 1), ":", 1) - 1) < 0 Then
                        iniRow2 = Application.Floor(mid$(Cells(1, 1).Value, InStr(6, Cells(1, 1), ",", 1) + 1, Len(Cells(1, 1)) - InStr(InStr(6, Cells(1, 1), ",", 1) + 1, Cells(1, 1), ":", 1) - 1), -1 * stepEk)
                    Else
                        iniRow2 = Application.Floor(mid$(Cells(1, 1).Value, InStr(6, Cells(1, 1), ",", 1) + 1, Len(Cells(1, 1)) - InStr(InStr(6, Cells(1, 1), ",", 1) + 1, Cells(1, 1), ":", 1) - 1), stepEk)
                    End If
                    If mid$(Cells(1, 1).Value, InStr(6, Cells(1, 1), ":", 1) + 1, InStr(InStr(6, Cells(1, 1), ":", 1) + 1, Cells(1, 1), ",", 1) - InStr(6, Cells(1, 1), ":", 1) - 1) < 0 Then
                        endRow1 = Application.Ceiling(mid$(Cells(1, 1).Value, InStr(6, Cells(1, 1), ":", 1) + 1, InStr(InStr(6, Cells(1, 1), ":", 1) + 1, Cells(1, 1), ",", 1) - InStr(6, Cells(1, 1), ":", 1) - 1), -1 * stepEk)
                    Else
                        endRow1 = Application.Ceiling(mid$(Cells(1, 1).Value, InStr(6, Cells(1, 1), ":", 1) + 1, InStr(InStr(6, Cells(1, 1), ":", 1) + 1, Cells(1, 1), ",", 1) - InStr(6, Cells(1, 1), ":", 1) - 1), stepEk)
                    End If
                    If mid$(Cells(1, 1).Value, InStr(InStr(6, Cells(1, 1), ",", 1) + 1, Cells(1, 1), ":", 1) + 1, Len(Cells(1, 1)) - InStr(InStr(6, Cells(1, 1), ",", 1) + 1, Cells(1, 1), ":", 1) - 1) < 0 Then
                        endRow2 = Application.Ceiling(mid$(Cells(1, 1).Value, InStr(InStr(6, Cells(1, 1), ",", 1) + 1, Cells(1, 1), ":", 1) + 1, Len(Cells(1, 1)) - InStr(InStr(6, Cells(1, 1), ",", 1) + 1, Cells(1, 1), ":", 1) - 1), -1 * stepEk)
                    Else
                        endRow2 = Application.Ceiling(mid$(Cells(1, 1).Value, InStr(InStr(6, Cells(1, 1), ",", 1) + 1, Cells(1, 1), ":", 1) + 1, Len(Cells(1, 1)) - InStr(InStr(6, Cells(1, 1), ",", 1) + 1, Cells(1, 1), ":", 1) - 1), stepEk)
                    End If                    
                    If StrComp(LCase(Cells(10, 3 * i + 1).Value), "pe", 1) = 0 Then
                        For j = 0 To numDataT - 1
                            If iniRow1 <= Cells(12 + numDataT + 8 + j, 3 * i + 2).Value Then
                                p = j + 1
                                Exit For
                            End If
                        Next
                        
                        For j = 0 To numDataT - 1
                            If endRow1 <= Cells(12 + numDataT + 8 + j, 3 * i + 2).Value Then
                                q = j + 1
                                Exit For
                            End If
                        Next
                        
                        If p >= 1 And q > p Then
                            Set rng = Range(Cells(11 + p - 1, (3 + (i * 3))), Cells(11 + q - 1, (3 + (i * 3))))
                            Cells(9, 3 * i + 2).Value = Application.WorksheetFunction.Average(rng)
                        End If
                        
                        For j = 0 To numDataT - 1
                            If iniRow2 >= Cells(11 + (numDataT * 2) + 8 - j, 3 * i + 2).Value Then
                                q = j + 1
                                Exit For
                            End If
                        Next
                        
                        For j = 0 To numDataT - 1
                            If endRow2 >= Cells(11 + (numDataT * 2) + 8 - j, 3 * i + 2).Value Then
                                p = j + 1
                                Exit For
                            End If
                        Next
                        
                        If p >= 1 And q > p Then
                            Set dataData = Range(Cells(10 + numDataT - q + 1, (3 + (i * 3))), Cells(10 + numDataT - p + 1, (3 + (i * 3))))
                            Cells(9, 3 * i + 3).Value = 1 / Abs(Application.WorksheetFunction.Average(dataData) - Cells(9, 3 * i + 2).Value)
                        End If
                    Else
                        For j = 0 To numDataT - 1
                            If iniRow1 <= Cells(11 + (numDataT * 2) + 8 - j, 3 * i + 2).Value Then
                                p = j + 1
                                Exit For
                            End If
                        Next
                        
                        For j = 0 To numDataT - 1
                            If endRow1 <= Cells(11 + (numDataT * 2) + 8 - j, 3 * i + 2).Value Then
                                q = j + 1
                                Exit For
                            End If
                        Next
                        
                        If p >= 1 And q > p Then
                            Set rng = Range(Cells(10 + numDataT - q + 1, (3 + (i * 3))), Cells(10 + numDataT - p + 1, (3 + (i * 3))))
                            Cells(9, 3 * i + 2).Value = Application.WorksheetFunction.Average(rng)
                        End If
                        
                        For j = 0 To numDataT - 1
                            If iniRow2 >= Cells(12 + numDataT + 8 + j, 3 * i + 2).Value Then
                                q = j + 1
                                Exit For
                            End If
                        Next
                        
                        For j = 0 To numDataT - 1
                            If endRow2 >= Cells(12 + numDataT + 8 + j, 3 * i + 2).Value Then
                                p = j + 1
                                Exit For
                            End If
                        Next
                        
                        If p >= 1 And q > p Then
                            Set dataData = Range(Cells(11 + p - 1, (3 + (i * 3))), Cells(11 + q - 1, (3 + (i * 3))))
                            Cells(9, 3 * i + 3).Value = 1 / Abs(Application.WorksheetFunction.Average(dataData) - Cells(9, 3 * i + 2).Value)
                        End If
                    End If
                End If
            ElseIf IsNumeric(mid$(Cells(1, 1).Value, 5, Len(Cells(1, 1)) - 4)) = True Then
                k = mid$(Cells(1, 1).Value, 5, Len(Cells(1, 1)) - 4)
                If k >= 0 And k < numDataT / 2 Then
                Else
                    k = 0
                End If
                
                If k = 0 Then       ' Auto0 makes all default
                    Cells(9, 3 * i + 2).Value = 0
                    Cells(9, 3 * i + 3).Value = 1
                ElseIf Cells(10 + k, (3 + (i * 3))).Value > Cells(11 + numDataT - k, (3 + (i * 3))).Value Then  ' PES mode
                    Cells(9, 3 * i + 2).Value = Cells(11 + numDataT - k, (3 + (i * 3))).Value
                    Cells(9, 3 * i + 3).Value = 1 / (Cells(10 + k, (3 + (i * 3))).Value - Cells(11 + numDataT - k, (3 + (i * 3))).Value)
                Else    ' XAS mode
                    Cells(9, 3 * i + 2).Value = Cells(10 + k, (3 + (i * 3))).Value
                    Cells(9, 3 * i + 3).Value = 1 / (Cells(11 + numDataT - k, (3 + (i * 3))).Value - Cells(10 + k, (3 + (i * 3))).Value)
                End If
            Else
                k = 0
                
                If StrComp(LCase(Cells(10, 3 * i + 1).Value), "pe", 1) = 0 Then 'XAS mode
                    Cells(9, 3 * i + 2).Value = Cells(11 + k, (3 + (i * 3))).Value
                    Cells(9, 3 * i + 3).Value = 1 / (Cells(10 + numDataT - k, (3 + (i * 3))).Value - Cells(11 + k, (3 + (i * 3))).Value)
                Else    ' PES mode
                    Cells(9, 3 * i + 2).Value = Cells(10 + numDataT - k, (3 + (i * 3))).Value
                    Cells(9, 3 * i + 3).Value = 1 / (Cells(11 + k, (3 + (i * 3))).Value - Cells(10 + numDataT - k, (3 + (i * 3))).Value)
                End If
            End If
        Else ' point calibration in "auto" at start and end points
            k = 0
            If Cells(11 + k, (3 + (i * 3))).Value > Cells(10 + numDataT - k, (3 + (i * 3))).Value Then  ' PES mode
                Cells(9, 3 * i + 2).Value = Cells(10 + numDataT - k, (3 + (i * 3))).Value
                Cells(9, 3 * i + 3).Value = 1 / (Cells(11 + k, (3 + (i * 3))).Value - Cells(10 + numDataT - k, (3 + (i * 3))).Value)
            Else    ' XAS mode
                Cells(9, 3 * i + 2).Value = Cells(11 + k, (3 + (i * 3))).Value
                Cells(9, 3 * i + 3).Value = 1 / (Cells(10 + numDataT - k, (3 + (i * 3))).Value - Cells(11 + k, (3 + (i * 3))).Value)
            End If
        End If
    Next
    
    Cells(40, para + 11).Value = Cells(1, 1).Value
    Cells(1, 1).Value = "Grating"
    If ncomp > 0 Then
        strErr = "skip"
    Else
        off = 0
        multi = 1
    End If
End Sub

Sub ExportCmp()
    Dim numDataT As Integer
    
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
    
    If LCase(Cells(1, 1).Value) = "exp" Or strList = "Is" Then
        If strList = "Is" Then
            Cells(1, 1).Value = "Grating"
            ncomp = 0
        Else
            Cells(1, 1).Value = "Goto Exp_sheet"
        End If
        
        For i = 0 To ncomp
            Set rng = Range(Cells(11, (2 + (i * 3))), Cells(11, (2 + (i * 3))).End(xlDown))
            numDataT = Application.CountA(rng)
            sheetGraph.Range(Cells(11 + numDataT + 8, (2 + (i * 3))), Cells(11 + (numDataT * 2) + 8, (3 + (i * 3)))).Copy
            sheetAna.Cells(1, 1 + (i * 2)).PasteSpecial Paste:=xlValues
        Next
        
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    Application.CutCopyMode = False
    If strList = "Is" Then
    Else
        strErr = "skip"
    End If
End Sub

Sub ExportChk()
    If ExistSheet(strSheetAnaName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetAnaName).Delete
        Application.DisplayAlerts = True
    End If
        
    Worksheets.Add().Name = strSheetAnaName
    Set sheetAna = Worksheets(strSheetAnaName)
    Set sheetCheck = Worksheets(strSheetCheckName)
    wb = ActiveWorkbook.Name
    sheetCheck.Activate
    
    If LCase(Cells(1, 1).Value) = "exp" And InStr(1, sh, "Check_") > 0 Then
        Cells(1, 1).Value = "Goto Eck_sheet"
        Set rng = Range(Cells(1, 1), Cells(1, 1).End(xlDown))
        iRow = Application.CountA(rng)
        Set rng = [1:1]
        iCol = Application.CountA(rng)

        For i = 0 To iCol - 3
            ' 0 for CPS, 1 for Ip, 2 for Ie, 3 for CPS/Ip
            sheetCheck.Range(Cells(1 + (8 + iRow) * 3, i + 2), Cells(iRow - 1 + (8 + iRow) * 3, i + 2)).Copy
            sheetAna.Cells(1, 2 + (i * 2)).PasteSpecial Paste:=xlValues
            sheetCheck.Range(Cells(2, 1), Cells(iRow - 1, 1)).Copy
            sheetAna.Cells(2, 1 + (i * 2)).PasteSpecial Paste:=xlValues
            sheetAna.Cells(1, 1 + (i * 2)).Value = "KE/eV"
            sheetAna.Cells(1, 2 + (i * 2)).Value = mid$(sh, 7, Len(sh) - 6) & "_" & mid$(sheetAna.Cells(1, 2 + (i * 2)).Value, Len(sheetAna.Cells(1, 2 + (i * 2)).Value), 1)
        Next
        
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    ElseIf LCase(Cells(1, 1).Value) = "exp" And InStr(1, sh, "Photo_") > 0 Then
        Cells(1, 1).Value = "Goto Eck_sheet"
        If InStr(1, sh, "_Is") > 0 Then
            p = 3
        ElseIf InStr(1, sh, "_Ip") > 0 Then
            p = 1
        Else
            p = 5
        End If
        Set rng = Range(Cells(11, 1), Cells(11, 1).End(xlDown))
        iRow = Application.CountA(rng)
        Set rng = Range(Cells(11, 1), Cells(11, 1).End(xlToRight))
        iCol = Application.CountA(rng)

        For i = 0 To iCol - 3
            ' 0 for Ie, 1 for Ip, 2 for Is, 3 for Is/Ip, 4 for If, 5 for If/Ip
            sheetCheck.Range(Cells(10 + (3 + iRow) * p, i + 2), Cells(iRow + 10 + (3 + iRow) * p, i + 2)).Copy
            sheetAna.Cells(1, 2 + (i * 2)).PasteSpecial Paste:=xlValues
            sheetCheck.Range(Cells(11, 1), Cells(iRow + 10, 1)).Copy
            sheetAna.Cells(2, 1 + (i * 2)).PasteSpecial Paste:=xlValues
            sheetAna.Cells(1, 1 + (i * 2)).Value = "PE/eV"
            sheetAna.Cells(1, 2 + (i * 2)).Value = mid$(sh, 7, Len(sh) - 6) & "_" & mid$(sheetAna.Cells(1, 2 + (i * 2)).Value, Len(sheetAna.Cells(1, 2 + (i * 2)).Value), 1)
        Next
        
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    ElseIf LCase(Cells(1, 1).Value) = "exr" Then
        Cells(1, 1).Value = "Goto Eck_sheet"
        Set rng = Range(Cells(1, 1), Cells(1, 1).End(xlDown))
        iRow = Application.CountA(rng)
        Set rng = [1:1]
        iCol = Application.CountA(rng)

        For i = 0 To iCol - 3
            ' 0 for CPS, 1 for Ip, 2 for Ie, 3 for CPS/Ip
            sheetCheck.Range(Cells(1 + (8 + iRow) * 0, i + 2), Cells(iRow - 1 + (8 + iRow) * 0, i + 2)).Copy
            sheetAna.Cells(1, 2 + (i * 2)).PasteSpecial Paste:=xlValues
            sheetCheck.Range(Cells(2, 1), Cells(iRow - 1, 1)).Copy
            sheetAna.Cells(2, 1 + (i * 2)).PasteSpecial Paste:=xlValues
            sheetAna.Cells(1, 1 + (i * 2)).Value = "ME/eV"
        Next
        
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    Application.CutCopyMode = False
    strErr = "skip"
End Sub

Sub Convert2Txt()
    Dim numDataT As Integer
    Dim numDataF As Integer
    Dim ElemT As String
    
    Set rng = [1:1]
    iCol = Application.CountA(rng)
    strCpa = ActiveWorkbook.Path
    strSheetAnaName = ActiveSheet.Name
    Set sheetAna = Worksheets(strSheetAnaName)
    ElemT = vbNullString
    numDataF = FreeFile
    ' http://www.homeandlearn.org/write_to_a_text_file.html
    For i = 0 To (iCol / 2) - 1
        If iCol <= 3 Then
            If strList = "Ip" Then
                strLabel = strSheetAnaName
            ElseIf strList = "Is" Then
                strLabel = strSheetDataName
            Else
                strLabel = strSheetDataName
            End If
            iCol = 2
        Else
            strLabel = sheetAna.Cells(1, 2 + (i * 2)).Value
        End If
        strTest = strCpa & "\" & strLabel & ".txt"
        Set rng = sheetAna.Range(Cells(1, 2 + (i * 2)), Cells(1, (2 + (i * 2))).End(xlDown))
        numDataT = Application.CountA(rng)
        
        Open strTest For Output As #numDataF
        For j = 1 To numDataT
            For k = 1 To 2
                If k = 2 Then
                    ElemT = ElemT + Trim(sheetAna.Cells(j, k + (i * 2)).Value)
                Else
                    ElemT = Trim(sheetAna.Cells(j, k + (i * 2)).Value) + vbTab
                End If
            Next k
            Print #numDataF, ElemT
            ElemT = vbNullString
        Next j
        Close #numDataF
        numDataF = numDataF + 1
    Next i
    
    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub

    Application.CutCopyMode = False
    If strList = "Is" Or strList = "Ip" Then
    Else
        strErr = "skip"
    End If
End Sub

Sub FitRatioAnalysis()
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
    ChDrive mid$(ActiveWorkbook.Path, 1, 1)
    ChDir ActiveWorkbook.Path
    OpenFileName = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Please select a file", MultiSelect:=True)
    
    If IsArray(OpenFileName) Then
        If UBound(OpenFileName) > para / 3 Then
            TimeCheck = MsgBox("Stop a comparison because you select too many files: " & UBound(OpenFileName) & " over the total limit: " & para / 3, vbExclamation)
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
        
        strAna = "FitRatioAnalysis"
        
        sheetAna.Activate
        spacer = sheetAna.Cells(2, para + 1).Value
        g = sheetAna.Cells(3, para + 1).Value         ' # of Fit peaks
        fileNum = sheetAna.Cells(4, para + 1).Value   ' # of Fit files
        sheetAna.Cells(5, para + 1).Value = UBound(OpenFileName)  ' # of Ana files
        sheetAna.Cells(5, para).Value = "# ana files"
        scanNum = UBound(OpenFileName)
        sheetAna.Cells(1, 1) = vbNullString
        D = sheetAna.Range(Cells(1, 1), Cells(para * 3 - 1, para * 3 - 1)) ' No check in matching among the peak names.
        sheetFit.Activate
        sheetFit.Range(Cells(1, 1), Cells(para * 3 - 1, para * 3 - 1)) = D
        A = sheetFit.Range(Cells(4, para / 2), Cells(3 + fileNum, para - 1))    ' store the BGs
        For i = 1 To fileNum
            A(i, 1) = D(3 + i, g + 6) & D(3 + i, g + 7) & D(3 + i, g + 8)
        Next
        sheetFit.Range(Cells(1, 5 + g), Cells((spacer + fileNum) * 5 + 3, 10 + g * 2)).ClearContents
        sheetFit.Cells(1, 4 + g).Value = ActiveWorkbook.Name
        sheetFit.Cells(2, 4 + g).Value = sheetAna.Name
        sheetFit.Cells(1, 1).Value = "Multiple-element ratio analysis"
        D = sheetFit.Range(Cells(1, 1), Cells(para * 3 - 1, para * 3 - 1))
        
        i = 0
        q = 0
        j = 0
        Call EachComp       ' Copy fitting parameters in each Fit sheet
        
        sheetFit.Activate
        D(1, 4) = "File"
        D(2, 4) = "Sheet"
        D(3, g + 6) = "Background"      ' G is # of peaks in the main sheet. Peaks over this # do not appear.
        D(2, g + 8 + scanNum) = "Difference"   ' scanNum represents number of BGs
        D(3 + (spacer + fileNum - 1), g + 6) = "Total peak area"
        D(2 + (spacer + fileNum - 1), g + 9) = "T.I.Area ratio"
        D(3 + (spacer + fileNum - 1) * 2, g + 6) = "Summation"             ' you can choose
        D(2 + (spacer + fileNum - 1) * 2, g + 9) = "S.I.Area ratio"            ' normalized by summation
        D(3 + (spacer + fileNum - 1) * 2, 2 * g + 9) = "Total ratio"
        D(3 + (spacer + fileNum - 1) * 3, g + 6) = "Summation"               ' you can choose
        D(2 + (spacer + fileNum - 1) * 3, g + 9) = "N.I.Area ratio"            ' normalized by summation
        D(3 + (spacer + fileNum - 1) * 3, 2 * g + 9) = "Total ratio"
        D(3 + (spacer + fileNum - 1) * 4, g + 6) = "Average"
        
        For i = 0 To 4      ' i represents # of parameters to be summarized
            Range(Cells(3 - i + (spacer + fileNum) * i, 5), Cells(3 - i + (spacer + fileNum) * i, 4 + g)).Interior.ColorIndex = 38
            Cells(3 + (spacer + fileNum - 1) * i, 1).Interior.ColorIndex = 3
            Range(Cells(3 + (spacer + fileNum - 1) * i, 2), Cells(3 + (spacer + fileNum - 1) * i, 3)).Interior.ColorIndex = 4
            Cells(3 + (spacer + fileNum - 1) * i, 4).Interior.ColorIndex = 5
            If i = 0 Then
                Range(Cells(3 + (spacer + fileNum - 1) * i, g + 6), Cells(3 + (spacer + fileNum - 1) * i, g + 6 + scanNum)).Interior.ColorIndex = 6
            Else
                Range(Cells(3 + (spacer + fileNum - 1) * i, g + 6), Cells(3 + (spacer + fileNum - 1) * i, g + 7)).Interior.ColorIndex = 6
            End If
            Cells(3 + (spacer + fileNum - 1) * i, 4).Font.ColorIndex = 2
            For k = 0 To fileNum - 1
                D(4 + k + (spacer + fileNum - 1) * i, 4) = g
            Next
            For k = 0 To g - 1
                D(3 + (spacer + fileNum - 1) * 2, g + 9 + k) = D(3 + (spacer + fileNum - 1) * 2, 5 + k)
                D(3 + (spacer + fileNum - 1) * 3, g + 9 + k) = D(3 + (spacer + fileNum - 1) * 3, 5 + k)
            Next
        Next
        
        Cells(1, 4).Interior.ColorIndex = 9
        Cells(2, 4).Interior.ColorIndex = 10
        For i = 0 To 1
            Cells(1 + i, 4).Font.ColorIndex = 2
        Next

        Range(Cells(2 + (spacer + fileNum - 1) * 0, g + 8 + scanNum), Cells(2 + (spacer + fileNum - 1) * 0, g + 9 + scanNum)).Interior.ColorIndex = 8  ' Difference
        For i = 1 To 4
            Range(Cells(2 + (spacer + fileNum - 1) * i, g + 9), Cells(2 + (spacer + fileNum - 1) * i, g + 10)).Interior.ColorIndex = 8   ' Area ratio
        Next
        
        Cells(3 + (spacer + fileNum - 1) * 2, 2 * g + 9).Interior.ColorIndex = 26   ' Total ratio in S. Area ratio
        Cells(3 + (spacer + fileNum - 1) * 3, 2 * g + 9).Interior.ColorIndex = 26   ' Total ratio in N. Area ratio
        Range(Cells(3 + (spacer + fileNum - 1) * 2, g + 9), Cells(3 + (spacer + fileNum - 1) * 2, 2 * g + 8)).Interior.ColorIndex = 38  ' Peak names in S. Area ratio
        Range(Cells(3 + (spacer + fileNum - 1) * 3, g + 9), Cells(3 + (spacer + fileNum - 1) * 3, 2 * g + 8)).Interior.ColorIndex = 38  ' Peak names in N. Area ratio
        sheetFit.Range(Cells(1, 1), Cells(para - 1, para - 1)) = D
        sheetFit.Range(Cells(4, g + 6), Cells(3 + fileNum, 2 * g + 6)) = A ' back BG
        
        For i = 0 To fileNum - 1
            Cells(4 + i + 1 * (spacer + fileNum - 1), g + 6).FormulaR1C1 = "=Sum(RC5:RC" & (g + 4) & ")"                     ' Total P.Area
            Cells(4 + i + 2 * (spacer + fileNum - 1), g + 6).FormulaR1C1 = "=Sum(RC5:RC" & (g + 4) & ")"                     ' Total S.Area
            Cells(4 + i + 3 * (spacer + fileNum - 1), g + 6).FormulaR1C1 = "=Sum(RC5:RC" & (g + 4) & ")"                     ' Total N.Area
            Cells(4 + i + 4 * (spacer + fileNum - 1), g + 6).FormulaR1C1 = "=Average(RC5:RC" & (g + 4) & ")"                 ' Avg FHHM
            For p = 0 To g - 2
                Cells(4 + i, g + 8 + scanNum + p).FormulaR1C1 = "=(RC" & (6 + p) & " - RC" & (5 + p) & ")"                             ' Difference
                Cells(4 + i + 1 * (spacer + fileNum - 1), g + 9 + p).FormulaR1C1 = "=(RC" & (5 + p) & " / RC" & (6 + p) & ")"   ' P.Area ratio
            Next
            
            For p = 0 To g - 1
                Cells(4 + i + 2 * (spacer + fileNum - 1), g + 9 + p).FormulaR1C1 = "=(100 * RC" & (5 + p) & "/RC" & (g + 6) & ")"  ' S.Area ratio
            Next
            Cells(4 + i + 2 * (spacer + fileNum - 1), 2 * g + 9).FormulaR1C1 = "=Sum(RC[" & (-g) & "]:RC[-1])"               ' Total S.Area ratio
            
            For p = 0 To g - 1
                Cells(4 + i + 3 * (spacer + fileNum - 1), g + 9 + p).FormulaR1C1 = "=(100 * RC" & (5 + p) & "/RC" & (g + 6) & ")"  ' N.Area ratio
            Next
            Cells(4 + i + 3 * (spacer + fileNum - 1), 2 * g + 9).FormulaR1C1 = "=Sum(RC[" & (-g) & "]:RC[-1])"               ' Total N.Area ratio
        Next
        
        For i = 0 To 4
            If i > 0 Then
                For k = 0 To g - 1
                    Cells(3 + (spacer + fileNum - 1) * i, k + 5).FormulaR1C1 = "=R3C" & (k + 5) & ""
                Next
            End If
            
            Set dataBGraph = Range(Cells(4 + (spacer + fileNum - 1) * i, 5), Cells(4 + (spacer + fileNum - 1) * i, 5).Offset(fileNum - 1, g - 1))
            
            Charts.Add
            ActiveChart.ChartType = xlLineMarkers
            ActiveChart.SetSourceData Source:=dataBGraph, PlotBy:=xlColumns
            ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetFitName

            For k = 1 To g
                ActiveChart.SeriesCollection(k).Name = "='" & ActiveSheet.Name & "'!R3C" & (4 + k) & ""  ' Cells(3, 4 + k).Value
                ActiveChart.SeriesCollection(k).AxisGroup = 1
            Next
            
            If Cells(4 + (spacer + fileNum - 1) * i, 4).Value > 1 And i = 0 Then    ' difference
                For k = 1 To g - 1
                    Set dataKGraph = Range(Cells(4 + (spacer + fileNum - 1) * i, 2 * g + 7 + k - 1), Cells(4 + (spacer + fileNum - 1) * i + fileNum - 1, 2 * g + 7 + k - 1))
                    ActiveChart.SeriesCollection.NewSeries
                    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
                        .ChartType = xlColumnClustered
                        .Values = dataKGraph
                        Cells((3 + (spacer + fileNum - 1) * i), g + 7 + k + scanNum).FormulaR1C1 = "=R3C" & (5 + k) & " & ""-"" & R3C" & (4 + k) & ""
                        Cells((3 + (spacer + fileNum - 1) * i), g + 7 + k + scanNum).Interior.ColorIndex = 38
                        .Name = "='" & ActiveSheet.Name & "'!R3C" & (g + 7 + k + scanNum) & ""                'Cells(3, 5 + k).Value + "-" + Cells(3, 4 + k).Value
                        .AxisGroup = 2
                    End With
                Next
            ElseIf Cells(4 + (spacer + fileNum - 1) * i, 4).Value > 1 And i = 1 Then
                For k = 1 To g - 1
                    Set dataKGraph = Range(Cells(4 + (spacer + fileNum - 1) * i, g + 9 + k - 1), Cells(4 + (spacer + fileNum - 1) * i + fileNum - 1, g + 9 + k - 1))
                    ActiveChart.SeriesCollection.NewSeries
                    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
                        .ChartType = xlColumnClustered
                        .Values = dataKGraph
                        Cells((3 + (spacer + fileNum - 1) * i), g + 8 + k).FormulaR1C1 = "=R3C" & (4 + k) & " & ""/"" & R3C" & (5 + k) & ""
                        Cells((3 + (spacer + fileNum - 1) * i), g + 8 + k).Interior.ColorIndex = 38
                        .Name = "='" & ActiveSheet.Name & "'!R" & (3 + (spacer + fileNum - 1) * i) & "C" & (g + 8 + k) & ""               'Cells(3, 4 + k).Value + "/" + Cells(3, 5 + k).Value
                        .AxisGroup = 2
                    End With
                Next
            ElseIf Cells(4 + (spacer + fileNum - 1) * i, 4).Value > 0 And i >= 2 And i <= 3 Then
                For k = 1 To g
                    Set dataKGraph = Range(Cells(4 + (spacer + fileNum - 1) * i, g + 9 + k - 1), Cells(4 + (spacer + fileNum - 1) * i + fileNum - 1, g + 9 + k - 1))
                    ActiveChart.SeriesCollection.NewSeries
                    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
                        .ChartType = xlAreaStacked100
                        Cells((3 + (spacer + fileNum - 1) * i), g + 8 + k).FormulaR1C1 = "= ""Rto_"" & R3C" & (4 + k) & ""
                        .Name = "='" & ActiveSheet.Name & "'!R" & (3 + (spacer + fileNum - 1) * i) & "C" & (g + 8 + k) & ""     'Cells(3, 4 + k).Value
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
                If i = 0 Then
                    .AxisTitle.Text = "Binding energy (eV)"
                ElseIf i = 1 Then
                    .AxisTitle.Text = "T.I. Area"
                ElseIf i = 2 Then
                    .AxisTitle.Text = "S.I. Area"
                ElseIf i = 3 Then
                    .AxisTitle.Text = "N.I. Area"
                ElseIf i = 4 Then
                    .AxisTitle.Text = "FWHM (eV)"
                End If
                .AxisTitle.Font.Size = 12
                .AxisTitle.Font.Bold = False
            End With
            
            If i < 3 And g > 1 Then
                With ActiveChart.Axes(xlValue, xlSecondary)
                    .HasTitle = True
                    If i = 0 Then
                        .AxisTitle.Text = "Difference (eV)"
                    ElseIf i = 1 Then
                        .AxisTitle.Text = "Ratio (peak-to-peak)"
                    ElseIf i = 2 Then
                        .AxisTitle.Text = "Ratio (%)"
                    End If
                    .AxisTitle.Font.Size = 12
                    .AxisTitle.Font.Bold = False
                End With
            End If
        
            With ActiveSheet.ChartObjects(1 + i)
                .Top = 20 + (500 / (windowSize * 2)) * i
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
        TimeCheck = MsgBox("Stop a comparison; no file selected.", vbExclamation)
    End If
    
SkipFitRatioAnalysis:
    Call GetOut
End Sub

Sub FitAnalysis()
    strSheetAnaName = "Ana_" + strSheetDataName
    strSheetFitName = "Fit_" + strSheetDataName
    strSheetGraphName = "Graph_" + strSheetDataName
    
    If ExistSheet(strSheetAnaName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetAnaName).Delete
        Application.DisplayAlerts = True
    End If
        
    Worksheets.Add().Name = strSheetAnaName
    Set sheetAna = Worksheets(strSheetAnaName)
    Set sheetFit = Worksheets(strSheetFitName)
    Set sheetGraph = Worksheets(strSheetGraphName)

    ChDrive mid$(ActiveWorkbook.Path, 1, 1)
    ChDir ActiveWorkbook.Path
    OpenFileName = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Please select a file", MultiSelect:=True)
    
    startTime = Timer
    
    If IsArray(OpenFileName) Then
        If UBound(OpenFileName) > para / 3 Then
            TimeCheck = MsgBox("Stop a comparison because you select too many files: " & UBound(OpenFileName) & " over the total limit: " & para / 3, vbExclamation)
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf UBound(OpenFileName) > 1 Then
        End If
        
        strAna = "FitAnalysis"
        
        Cells(1, para).Value = "Parameters"
        Cells(2, para).Value = "Spacer"
        Cells(3, para).Value = "# peaks"
        Cells(4, para).Value = "# fit files"
        Cells(2, para + 1).Value = spacer
        Cells(3, para + 1).Value = g
        fileNum = UBound(OpenFileName)
        Cells(4, para + 1).Value = fileNum + 1
        
        D = sheetAna.Range(Cells(1, 1), Cells((4 + spacer * 4) + 5 * fileNum, 9 + 2 * g)) ' No check in matching among the peak names.
        D(3, g + 6) = "Background"      ' G is # of peaks in the main sheet. Peaks over this # do not appear.
        D(2, g + 9) = "Difference"
        D(2, 1) = "BE"
        D(2 + (spacer + fileNum), 1) = "T.I.Area"
        D(3 + (spacer + fileNum), g + 6) = "Total peak area"
        numData = sheetFit.Cells(5, 101).Value
        D(2 + (spacer + fileNum), g + 9) = "T.I.Area ratio"
        D(2 + (spacer + fileNum) * 2, 1) = "S.I.Area"
        D(3 + (spacer + fileNum) * 2, g + 6) = "Summation"               ' you can choose
        D(2 + (spacer + fileNum) * 2, g + 9) = "S.I.Area ratio"            ' normalized by summation
        D(3 + (spacer + fileNum) * 2, 2 * g + 9) = "Total ratio"
        D(2 + (spacer + fileNum) * 3, 1) = "N.I.Area"
        D(3 + (spacer + fileNum) * 3, g + 6) = "Summation"               ' you can choose
        D(2 + (spacer + fileNum) * 3, g + 9) = "N.I.Area ratio"            ' normalized by summation
        D(3 + (spacer + fileNum) * 3, 2 * g + 9) = "Total ratio"
        D(2 + (spacer + fileNum) * 4, 1) = "FWHM"
        D(3 + (spacer + fileNum) * 4, g + 6) = "Average"
        
        For iCol = 0 To g - 1
            D(3, iCol + 5) = C(1, iCol + 1)                                 ' Peak #1
            D(4, iCol + 5) = C(2, iCol + 1)                                 ' BE
            D(3 + (spacer + fileNum), iCol + 5) = C(1, iCol + 1)         ' Peak #2
            D(3 + (spacer + fileNum) * 2, iCol + 5) = C(1, iCol + 1)     ' Peak #3
            D(3 + (spacer + fileNum) * 3, iCol + 5) = C(1, iCol + 1)     ' Peak #4
            D(3 + (spacer + fileNum) * 2, iCol + 9 + g) = C(1, iCol + 1) ' Peak #3 for ratio
            D(3 + (spacer + fileNum) * 3, iCol + 9 + g) = C(1, iCol + 1) ' Peak #4 for ratio
            
            If C(10 + sftfit2, iCol + 1) > 0 Then
                D(4 + (spacer + fileNum), iCol + 5) = C(10 + sftfit2, iCol + 1)      ' P.Area
                D(4 + (spacer + fileNum) * 2, iCol + 5) = C(11 + sftfit2, iCol + 1)  ' S.Area
                D(4 + (spacer + fileNum) * 3, iCol + 5) = C(12 + sftfit2, iCol + 1)  ' N.Area
            Else
                D(4 + (spacer + fileNum), iCol + 5) = 0      ' P.Area
                D(4 + (spacer + fileNum) * 2, iCol + 5) = 0  ' S.Area
                D(4 + (spacer + fileNum) * 3, iCol + 5) = 0  ' N.Area
            End If

            D(3 + (spacer + fileNum) * 4, iCol + 5) = C(1, iCol + 1)     ' Peak #5
            D(4 + (spacer + fileNum) * 4, iCol + 5) = C(4, iCol + 1)     ' FWHM
        Next

        For i = 0 To 4      ' i represents # of parameters to be summarized
            D(3 + (spacer + fileNum) * i, 1) = "File"
            D(3 + (spacer + fileNum) * i, 2) = "Sheet"
            D(3 + (spacer + fileNum) * i, 4) = "# peaks"
            D(4 + (spacer + fileNum) * i, 4) = sheetFit.Cells(8 + sftfit2, 2).Value
            D(4 + (spacer + fileNum) * i, 1) = wb                  ' File name
            D(4 + (spacer + fileNum) * i, 2) = strSheetFitName     ' Sheet name
            Range(Cells(3 + (spacer + fileNum) * i, 5), Cells(3 + (spacer + fileNum) * i, 4 + g)).Interior.ColorIndex = 38
            Cells(3 + (spacer + fileNum) * i, 1).Interior.ColorIndex = 3
            Range(Cells(3 + (spacer + fileNum) * i, 2), Cells(3 + (spacer + fileNum) * i, 3)).Interior.ColorIndex = 4
            Cells(3 + (spacer + fileNum) * i, 4).Interior.ColorIndex = 33
            Range(Cells(3 + (spacer + fileNum) * i, g + 6), Cells(3 + (spacer + fileNum) * i, g + 7)).Interior.ColorIndex = 6
            Range(Cells(2 + (spacer + fileNum) * i, g + 9), Cells(2 + (spacer + fileNum) * i, g + 10)).Interior.ColorIndex = 8
        Next

        Cells(3 + (spacer + fileNum) * 2, 2 * g + 9).Interior.ColorIndex = 26
        Cells(3 + (spacer + fileNum) * 3, 2 * g + 9).Interior.ColorIndex = 26
        Range(Cells(3 + (spacer + fileNum) * 2, g + 9), Cells(3 + (spacer + fileNum) * 2, 2 * g + 8)).Interior.ColorIndex = 38
        Range(Cells(3 + (spacer + fileNum) * 3, g + 9), Cells(3 + (spacer + fileNum) * 3, 2 * g + 8)).Interior.ColorIndex = 38
        For i = 0 To 2
            D(4, g + 6 + i) = b(1, 1 + i)                                   ' BG
        Next

        i = 0
        q = 0
        j = 0
        Call EachComp       ' Copy fitting parameters in each Fit sheet
        
        sheetAna.Activate
        sheetAna.Range(Cells(1, 1), Cells((4 + spacer * 4) + 5 * fileNum, 9 + 2 * g)) = D
        
        For i = 0 To fileNum - cae
            Cells(4 + i + spacer + fileNum, g + 6).FormulaR1C1 = "=Sum(RC5:RC" & (g + 4) & ")"                       ' Total P.Area
            Cells(4 + i + 2 * (spacer + fileNum), g + 6).FormulaR1C1 = "=Sum(RC5:RC" & (g + 4) & ")"                     ' Total S.Area
            Cells(4 + i + 3 * (spacer + fileNum), g + 6).FormulaR1C1 = "=Sum(RC5:RC" & (g + 4) & ")"                     ' Total N.Area
            Cells(4 + i + 4 * (spacer + fileNum), g + 6).FormulaR1C1 = "=Average(RC5:RC" & (g + 4) & ")"                 ' Avg FHHM
            For p = 0 To g - 2
                Cells(4 + i, g + 9 + p).FormulaR1C1 = "=(RC" & (6 + p) & " - RC" & (5 + p) & ")"                            ' Difference
                Cells(4 + i + spacer + fileNum, g + 9 + p).FormulaR1C1 = "=(RC" & (5 + p) & " / RC" & (6 + p) & ")"    ' P.Area ratio
            Next
            
            For p = 0 To g - 1
                Cells(4 + i + 2 * (spacer + fileNum), g + 9 + p).FormulaR1C1 = "=(100 * RC" & (5 + p) & "/RC" & (g + 6) & ")"  ' S.Area ratio
            Next
            Cells(4 + i + 2 * (spacer + fileNum), 2 * g + 9).FormulaR1C1 = "=Sum(RC[" & (-g) & "]:RC[-1])"               ' Total S.Area ratio
            
            For p = 0 To g - 1
                Cells(4 + i + 3 * (spacer + fileNum), g + 9 + p).FormulaR1C1 = "=(100 * RC" & (5 + p) & "/RC" & (g + 6) & ")"  ' N.Area ratio
            Next
            Cells(4 + i + 3 * (spacer + fileNum), 2 * g + 9).FormulaR1C1 = "=Sum(RC[" & (-g) & "]:RC[-1])"               ' Total N.Area ratio
        Next
        
        For i = 0 To 4
            If i > 0 Then
                For k = 0 To g - 1
                    Cells(3 + (spacer + fileNum) * i, k + 5).FormulaR1C1 = "=R3C" & (k + 5) & ""
                Next
            End If
            
            Set dataBGraph = Range(Cells(4 + (spacer + fileNum) * i, 5), Cells(4 + (spacer + fileNum) * i, 5).Offset(fileNum, Cells(4 + (spacer + fileNum) * i, 4) - 1))
            
            Charts.Add
            ActiveChart.ChartType = xlLineMarkers
            ActiveChart.SetSourceData Source:=dataBGraph, PlotBy:=xlColumns
            ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetAnaName
            
            For k = 1 To g
                If IsEmpty(Cells(3, 4 + k).Value) = True Then
                Else
                    ActiveChart.SeriesCollection(k).Name = "='" & ActiveSheet.Name & "'!R3C" & (4 + k) & ""  ' Cells(3, 4 + k).Value
                    ActiveChart.SeriesCollection(k).AxisGroup = 1
                End If
            Next
            
            If Cells(4 + (spacer + fileNum) * i, 4).Value > 1 And i < 2 Then
                For k = 1 To g - 1
                    Set dataKGraph = Range(Cells(4 + (spacer + fileNum) * i, g + 9 + k - 1), Cells(4 + (spacer + fileNum) * i + fileNum, g + 9 + k - 1))
                    ActiveChart.SeriesCollection.NewSeries
                    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
                        
                        .ChartType = xlColumnClustered
                        .Values = dataKGraph
                        If i = 0 Then
                            Cells(3, g + 8 + k).FormulaR1C1 = "=R3C" & (5 + k) & " & ""-"" & R3C" & (4 + k) & ""
                            Cells(3, g + 8 + k).Interior.ColorIndex = 38
                            .Name = "='" & ActiveSheet.Name & "'!R3C" & (g + 8 + k) & ""             'Cells(3, 5 + k).Value + "-" + Cells(3, 4 + k).Value
                        ElseIf i = 1 Then
                            Cells((3 + (spacer + fileNum) * i), g + 8 + k).FormulaR1C1 = "=R3C" & (4 + k) & " & ""/"" & R3C" & (5 + k) & ""
                            Cells((3 + (spacer + fileNum) * i), g + 8 + k).Interior.ColorIndex = 38
                            .Name = "='" & ActiveSheet.Name & "'!R" & (3 + (spacer + fileNum) * i) & "C" & (g + 8 + k) & ""            'Cells(3, 4 + k).Value + "/" + Cells(3, 5 + k).Value
                        End If
                        
                        .AxisGroup = 2
                        'SourceRangeColor2 = .Border.Color
                    End With
                Next
            ElseIf Cells(4 + (spacer + fileNum) * i, 4).Value > 0 And i >= 2 And i <= 3 Then
                For k = 1 To g
                    Set dataKGraph = Range(Cells(4 + (spacer + fileNum) * i, g + 9 + k - 1), Cells(4 + (spacer + fileNum) * i + fileNum, g + 9 + k - 1))
                    ActiveChart.SeriesCollection.NewSeries
                    With ActiveChart.SeriesCollection(ActiveChart.SeriesCollection.Count)
                        .ChartType = xlAreaStacked100
                        Cells((3 + (spacer + fileNum) * i), g + 8 + k).FormulaR1C1 = "= ""Rto_"" & R3C" & (4 + k) & ""
                        .Name = "='" & ActiveSheet.Name & "'!R" & (3 + (spacer + fileNum) * i) & "C" & (g + 8 + k) & ""   'Cells(3, 4 + k).Value
                        .Values = dataKGraph
                        .AxisGroup = 2
                        If i = 2 Then
                        ElseIf i = 3 Then
                        End If
                        'SourceRangeColor2 = .Border.Color
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
                If i = 0 Then
                    .AxisTitle.Text = "Binding energy (eV)"
                ElseIf i = 1 Then
                    .AxisTitle.Text = "T.I. Area"
                ElseIf i = 2 Then
                    .AxisTitle.Text = "S.I. Area"
                ElseIf i = 3 Then
                    .AxisTitle.Text = "N.I. Area"
                ElseIf i = 4 Then
                    .AxisTitle.Text = "FWHM (eV)"
                End If
                .AxisTitle.Font.Size = 12
                .AxisTitle.Font.Bold = False
            End With
            
            If i < 3 And g > 1 Then
                With ActiveChart.Axes(xlValue, xlSecondary)
                    .HasTitle = True
                    If i = 0 Then
                        .AxisTitle.Text = "Difference (eV)"
                    ElseIf i = 1 Then
                        .AxisTitle.Text = "Ratio (peak-to-peak)"
                    ElseIf i = 2 Then
                        .AxisTitle.Text = "Ratio (%)"
                    End If
                    .AxisTitle.Font.Size = 12
                    .AxisTitle.Font.Bold = False
                End With
            End If
        
            With ActiveSheet.ChartObjects(1 + i)
                .Top = 20 + (500 / (windowSize * 2)) * i
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
        strSheetAnaName = "Cmp_" + strSheetDataName
        
        If ExistSheet(strSheetAnaName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetAnaName).Delete
            Application.DisplayAlerts = True
        End If
            
        Worksheets.Add().Name = strSheetAnaName
        Set sheetAna = Worksheets(strSheetAnaName)
        
        i = 0
        q = 0
        sheetFit.Activate
        numData = Cells(5, 101).Value
        imax = numData + 10
        tmp = sheetFit.Range(Cells(20 + sftfit, 1), Cells(20 + sftfit + numData, 1))
        en = sheetFit.Range(Cells(20 + sftfit, 4), Cells(20 + sftfit + numData, 4))
        sheetAna.Activate
        sheetAna.Range(Cells(10, 1), Cells(10 + numData, 1)) = tmp
        sheetAna.Range(Cells(10, 3), Cells(10 + numData, 3)) = en
        sheetGraph.Activate
        dblMin = Cells(41, para + 10).Value
        dblMax = Cells(42, para + 10).Value
        multi = Cells(9, 3).Value
        
        If IsEmpty(Cells(51, para + 10)) = False Then
            'sheetGraph.Range(Cells(40, para + 9), Cells((Cells(51, para + 10).End(xlDown).Row), para + 30)).Copy sheetAna.Cells(40, para + 9)
            If Cells(42, para + 12) >= (Cells(43, para + 12) + Cells(42, para + 12)) Then
                sheetGraph.Range(Cells(40, para + 9), Cells((50 + Cells(42, para + 12).Value), para + 30)).Copy Destination:=sheetAna.Cells(40, para + 9)
            Else
                sheetGraph.Range(Cells(40, para + 9), Cells((50 + Cells(43, para + 12).Value + Cells(42, para + 12).Value), para + 30)).Copy Destination:=sheetAna.Cells(40, para + 9)
            End If
            sheetAna.Cells(41, para + 10).Value = dblMin * multi
            sheetAna.Cells(42, para + 10).Value = dblMax * multi
            sheetAna.Cells(45, para + 10).Value = fileNum
        End If
        
        sheetAna.Activate
        Cells(1, 2).Value = wb
        Cells(9, 1).Value = "Offset/multp"
        Cells(9, 2).Value = 0
        Cells(9, 3).Value = 1
        
        If StrComp(mid$(Cells(10, 1), 1, 2), "BE", 1) = 0 Then
            str1 = "Be"
            str2 = "Sh"
            str3 = "In"
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
            str1 = "Pe"
            str2 = "Sh"
            str3 = "Ab"
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
            str1 = "Po"
            str2 = "Sh"
            str3 = "Ab"
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
        Cells(9, 1).Interior.Color = RGB(139, 195, 74)
        Range(Cells(9, 2), Cells(9, 3)).Interior.Color = RGB(197, 225, 165)
        Set dataBGraph = Range(Cells(10 + (imax), 2), Cells((2 * imax) - 1, 3))
        
        Charts.Add
        ActiveChart.ChartType = xlXYScatterLinesNoMarkers 'xlXYScatterSmoothNoMarkers
        ActiveChart.SetSourceData Source:=dataBGraph, PlotBy:=xlColumns
        ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetAnaName
        ActiveChart.SeriesCollection(1).Name = "='" & ActiveSheet.Name & "'!R1C2"
        ActiveChart.ChartTitle.Delete
        SourceRangeColor1 = ActiveChart.SeriesCollection(1).Border.Color
        
        With ActiveChart.Axes(xlCategory, xlPrimary)
            If StrComp(str1, "Pe", 1) = 0 Then
                .MinimumScale = startEb
                .MaximumScale = endEb
                strLabel = "Photon energy (eV)"
            ElseIf StrComp(str1, "Po", 1) = 0 Then
                .MinimumScale = startEb
                .MaximumScale = endEb
                strLabel = "Position (a.u.)"
            Else
                .MinimumScale = endEb
                .MaximumScale = startEb
                .ReversePlotOrder = True
                .Crosses = xlMaximum
                strLabel = "Binding energy (eV)"
            End If
            .HasTitle = True
            .AxisTitle.Text = strLabel
        End With
        
        With ActiveChart.Axes(xlCategory, xlPrimary)
            .MinorTickMark = xlOutside
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .HasMajorGridlines = True
            '.MajorUnit = numMajorUnit
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With ActiveChart.Axes(xlValue)
            .HasTitle = True
            .AxisTitle.Text = "Intensity (arb. units)"
            '.MinimumScale = dblMin
            '.MaximumScale = dblMax
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
                '.Width = 100
                '.Height = 100
                .Top = (50 / windowSize)
                With .Format.Fill
                    .Visible = msoTrue
                    .ForeColor.RGB = RGB(255, 255, 255)
                    .ForeColor.TintAndShade = 0.1
                End With
            End With
            With .Chart
                '.PlotArea.Height = ((500 - 40) / windowSize)
                .PlotArea.Width = (((550 * windowRatio) - 40) / windowSize)
                .ChartArea.Border.LineStyle = 0
                '.ChartArea.Interior.ColorIndex = xlNone    'transparent plot
            End With
        End With

        Range(Cells(10, 1), Cells(10, 2)).Interior.Color = SourceRangeColor1
        Range(Cells(10, 2), Cells(10, 2)).Interior.Color = SourceRangeColor1
        Range(Cells(9 + (imax), 2), Cells(9 + (imax), 2)).Interior.Color = SourceRangeColor1
        strTest = mid$(Cells(1, 2).Value, 1, Len(Cells(1, 2).Value) - 5)
        Cells(8 + (imax), 2).Value = Cells(1, 2).Value
        Cells(9 + (imax), 1).Value = str1 + strTest
        Cells(9 + (imax), 2).Value = str2 + strTest
        Cells(9 + (imax), 3).Value = str3 + strTest
        
        strAna = "FitComp"
        Set sheetGraph = Worksheets(strSheetAnaName)
        Call PlotElem
        Call PlotChem
        
        i = 0
        q = 0
        j = 0
        ncomp = 0
        Call EachComp       ' Copy BG-substracted data in each Fit sheets.
        sheetAna.Activate
    Else
        TimeCheck = MsgBox("Stop a comparison; no file selected.", vbExclamation)
        
    End If
    
    endTime = Timer
    Call GetOut
End Sub

Sub TargetDataAnalysis()
    strTest = Cells(2, 1).Value
    strTest = mid$(strTest, 1, 5)
    
    If StrComp(strTest, "CLAM2", 1) = 0 Then
        If StrComp(strCpa, "comp", 1) = 0 Then
            Call GetCompare
        ElseIf StrComp(strAna, "ana", 1) = 0 Then
            Call FitCurve
        ElseIf StrComp(strAna, "chem", 1) = 0 Then
            Call PlotChem
        Else
            Call FormatCLAM2
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call PlotCLAM2
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call ElemXPS
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call ElemAES
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call PlotElem
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call FitCurve
        End If
    ElseIf StrComp(strTest, "Energ", 1) = 0 And Len(Cells(1, 1).Value) = 0 Then            ' Scan grating data anaysis with fixed gap
        If Len(Cells(7, 6).Value) = 0 Then GoTo SkipEngBL   ' XAS_simple: XAS mode without gap scan.
        If IsEmpty(Cells(12, 1).Value) = True Then
            Call GetOut  ' No data included.
        ElseIf StrComp(strCpa, "comp", 1) = 0 Then
            Call GetCompare
        ElseIf StrComp(strAna, "ana", 1) = 0 Then
            Call FitCurve
        Else
            Call EngBL      ' Grating scan at fixed gap
            Call descriptHidden1
            Call GetOut
        End If
        If StrComp(TimeCheck, "yes", 1) = 0 Then TimeCheck = vbNullString
    ElseIf StrComp(strTest, "Photo", 1) = 0 Then            ' Photoabsorption data analysis
SkipEngBL:
        If IsEmpty(Cells(12, 1).Value) = True Then
            Call GetOut
        ElseIf StrComp(strCpa, "comp", 1) = 0 Then
            Call GetCompare
        ElseIf StrComp(strAna, "ana", 1) = 0 Then
            Call FitCurve
        ElseIf StrComp(strAna, "chem", 1) = 0 Then
            Call PlotChem
        Else
            Call PhotoBL            ' XAS with and without gap scan
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call descriptHidden1
            Call ElemXPS
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call ElemAES
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call PlotElem
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
            Call FitCurve
        End If
    ElseIf Not StrComp(strTest, "CLAM2", 1) = 0 Then        ' Other data
        strTest = Cells(1, 8).Value
        strTest = mid$(strTest, 1, 3)
        
        If StrComp(strTest, "Acq", 1) = 0 Or StrComp(mid$(Cells(2, 3).Value, 1, 3), "Ele", 1) = 0 Then
            Call ThermoAvgBL                ' Avantage software exported data analysis: Alpha110
        End If
        strTest = Cells(1, 1).Value
        
        If InStr(strTest, "E/eV") > 0 Then          ' Manually imported data analsysis
            Do
                If InStr(strTest, "'") > 0 Then     ' remove "'" generated in Igor produced text
                    q = InStr(strTest, "'")
                    strTest = Left(strTest, q - 1) + mid(strTest, q + 1)
                Else
                    Cells(1, 1).Value = strTest
                    Exit Do
                End If
            Loop
            
            If InStr(Cells(1, 3).Value, "E/eV") > 0 Then
                Call Convert2Txt
                TimeCheck = MsgBox("Data were exported in the " & i & " text files.", vbExclamation)
            End If
            
            If StrComp(strCpa, "comp", 1) = 0 Then
                Call GetCompare
            ElseIf StrComp(strAna, "ana", 1) = 0 Then
                Call FitCurve
            ElseIf StrComp(strAna, "chem", 1) = 0 Then
                Call PlotChem
            Else
                Call KeBL            ' KE, BE, PE, GE, AE, ME/eV data setup
                If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
                If StrComp(strTest, "GE/eV", 1) = 0 Then        ' Grating scan with fixed gap
                    Call EngBL
                    Call descriptHidden1
                    Call GetOut
                Else
                    Call PlotCLAM2
                    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
                    Call ElemXPS
                    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
                    Call ElemAES
                    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
                    Call PlotElem
                    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
                    Call FitCurve
                End If
            End If
        Else
            If StrComp(mid$(Cells(1, 4).Value, 1, 6), "Elapse", 1) = 0 Then Call ScanDivide ' for Off
            If StrComp(mid$(Cells(1, 1).Value, 1, 3), "phi", 1) = 0 Then Call ExportPHI ' for MultiPak exported ascii csv
            If StrComp(mid$(Cells(1, 1).Value, 1, 7), "Dataset", 1) = 0 Then Call ExportKratos ' for Kratos exported ascii text
            If StrComp(mid$(Cells(1, 1).Value, 1, 9), "transpose", 1) = 0 Then Call TransposeSheet ' for MultiPak exported curve fit results
            If StrComp(TimeCheck, "yes", 1) = 0 Then TimeCheck = "yes2"
            Call GetOut
        End If
    End If
End Sub

Sub FormatCLAM2()
    If graphexist = 0 Then
        pe = mid$(Cells(8, 1).Value, 19, (Len(Cells(8, 1).Value) - 18 - 2))
        highpe(0) = pe
        wf = 4
        char = 0
        off = 0
        multi = 1
        ncomp = 0
    End If
    
    startEk = mid$(Cells(10, 1).Value, 21, (InStr(1, Cells(10, 1).Value, "-") - 21))
    endEk = mid$(Cells(10, 1).Value, (InStr(1, Cells(10, 1).Value, "-") + 1), (Len(Cells(10, 1).Value) - InStr(1, Cells(10, 1).Value, "-") - 3))
    stepEk = mid$(Cells(11, 1).Value, 22, (Len(Cells(11, 1).Value) - 21 - 3))
    strscanNum = mid$(Cells(15, 1).Value, 17, (Len(Cells(15, 1).Value) - 16 - 1))
    cae = mid$(Cells(9, 1).Value, (InStr(1, Cells(9, 1).Value, "=") + 1), (Len(Cells(9, 1))))
    g = mid$(Cells(12, 1).Value, 2, 4)
    numData = ((endEk - startEk) / stepEk) + 1  ' numData for CLAM2 DAQ
    
    If IsNumeric(strscanNum) = True Then
        scanNum = strscanNum
    ElseIf IsNumeric(strscanNum) = False Then
        scanNum = 1
        strscanNum = Cells(20 + ((scanNum - 1) * (3 + numData)), 2).Value

        Do While IsNumeric(strscanNum) = True
            scanNum = scanNum + 1
            strscanNum = Cells(20 + ((scanNum - 1) * (3 + numData)), 2).Value
        Loop

        scanNum = scanNum - 1
    Else
        scanNum = 1
    End If

    scanNumR = scanNum

    Do
        iniRow = 22 + ((scanNum - 1) * (4 + ((endEk - startEk) / stepEk)))
        endRow = iniRow + ((endEk - startEk) / stepEk)
        numData = ((endEk - startEk) / stepEk) + 1  ' numData for CLAM2 DAQ
    
        If Len(Cells(endRow, 1)) = 0 And scanNum > 1 Then
            scanNum = scanNum - 1
        ElseIf Len(Cells(endRow, 1)) = 0 And scanNum = 1 Then
            endEk = Cells(iniRow, 1).End(xlDown)
            endRow = iniRow + ((endEk - startEk) / stepEk)
            numData = ((endEk - startEk) / stepEk) + 1
            If endEk = 0 Then   ' if endRow is negative, data is only single cell.
                strErr = "skip"
                Exit Sub
            End If
            Exit Do
        Else
            Exit Do
        End If
    Loop
    
    Set dataData = Union(Range(Cells(iniRow, 1), Cells(endRow, 1)), Range(Cells(iniRow, 7), Cells(endRow, 7)))
    Set dataKeData = Range(Cells(iniRow, 1), Cells(endRow, 1))
    Set dataIntData = dataKeData.Offset(, 6)
    numscancheck = 0        ' previously used by p
    
    If Len(NoCheck) > 2 And StrComp(strTest, "CLAM2", 0) = 0 And scanNumR > 0 Then
        Call numMajorUnitsCheck
        strscanNum = strscanNumR
        Call ScanRangeCheck
        Call ScanCheck
    
        If Not scanNumR = scanNum Then
            If numscancheck <= 0 And scanNum > 1 Then
                totalDataPoints = scanNum * numData + (Cells(22 + ((scanNumR - 1) * (1 + numData)), 1).End(xlDown).Row) - (22 + ((scanNumR - 1) * (1 + numData))) + 1
            Else
                totalDataPoints = scanNum * numData
            End If
        Else
            totalDataPoints = scanNum * numData
        End If
    
        If StrComp(NoCheck, "Obb", 0) = 0 Then Call ObbCheck
    End If
End Sub

Sub SheetCheckGenerator()
    If ExistSheet(strSheetCheckName) Or NoCheck = "ON" Then Exit Sub
    
    Worksheets.Add().Name = strSheetCheckName
    Set sheetCheck = Worksheets(strSheetCheckName)
    Cells(1, 1).Value = "X"
    Cells(1, 2).Value = "Y"
    Cells(1, 3).Value = "Norm"
    Range(Cells(2, 1), Cells(1 + numData, 1)) = C
    Range(Cells(2, 2), Cells(1 + numData, 2)) = A
    Range(Cells(2, 3), Cells(1 + numData, 3)) = D
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
        .AxisTitle.Text = strLabel
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

Sub PlotCLAM2()
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
    Set dataIntGraph = dataKeGraph.Offset(, 2)
    dataKeGraph.Value = dataKeData.Value
    C = dataKeData
    U = dataIntData
    A = dataIntGraph
    
    If StrComp(strTest, "AE/eV", 1) = 0 Then
        b = Range(Cells(1, para + 1), Cells(1, para + 5))
        For i = 1 To numData
            startR = i - 1
            If (i - 1) < 1 Then startR = 1
            endR = i + wf
            If (i + wf) > numData Then endR = numData
            b(1, 1) = 0
            b(1, 2) = 0
            b(1, 3) = 0
            b(1, 4) = 0
            b(1, 5) = 0
            
            For j = startR To endR
                b(1, 1) = b(1, 1) + C(j, 1) * C(j, 1)
                b(1, 2) = b(1, 2) + C(j, 1)
                b(1, 3) = b(1, 3) + 1
                b(1, 4) = b(1, 4) + C(j, 1) * U(j, 1)
                b(1, 5) = b(1, 5) + U(j, 1)
            Next
            
            A(i, 1) = (b(1, 3) * b(1, 4) - b(1, 2) * b(1, 5)) / (b(1, 1) * b(1, 3) - b(1, 2) * b(1, 2))
            Range(Cells(11, 2), Cells((numData + 10), 2)) = U
        Next
    ElseIf InStr(strTest, "E/eV") > 0 Then
        If StrComp(Cells(1, 3).Value, "Ip", 1) = 0 Or StrComp(Cells(1, 3).Value, "Ie", 1) = 0 Then
            D = dataKeData.Offset(, 2)
        Else
            D = dataKeData.Offset(, para + 30)      ' Empty Ip
        End If
        
        For i = 1 To numData
        
            If IsEmpty(D(i, 1)) Then
                D(i, 1) = 1
            Else
                If IsNumeric(D(i, 1)) = False Then
                    D(i, 1) = 1
                Else
                    If D(i, 1) <= 0 Then
                        D(i, 1) = 1
                    End If
                End If
            End If
            
            A(i, 1) = (U(i, 1) / D(i, 1))
        Next
    Else
        If numscancheck <= 0 Then
            For i = 1 To numData
                If IsNumeric(U(i, 1)) = False Then Exit For
                A(i, 1) = U(i, 1) * 0.000000000001   ' if no scanrangecheck, dataIntData from the original data sheet, Ip normalized by pico-amps
            Next
        Else
            For i = 1 To numData
                If IsNumeric(U(i, 1)) = False Then Exit For
                A(i, 1) = U(i, 1) * 1                ' if scanrangecheck done, dataIntData from the check sheet (Avg. CPS/Ip; Sub ScanCheck), Ip is already normalized by pic-amps
            Next
        End If
    End If

    Range(Cells(11, 3), Cells((numData + 10), 3)) = A

    Call descriptGraph
    Call scalecheck
    
    If strTest = "ME/eV" Then Call SheetCheckGenerator      ' Check Sheet for "ME/eV"
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
    
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers 'xlXYScatterSmoothNoMarkers
    ActiveChart.SetSourceData Source:=dataBGraph, PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetGraphName
    ActiveChart.SeriesCollection(1).Name = ActiveWorkbook.Name  '"BE graph"
    ActiveChart.ChartTitle.Delete
    
    With ActiveChart.Axes(xlCategory, xlPrimary)
        If StrComp(str1, "Pe", 1) = 0 Or StrComp(str3, "De", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then
            .MinimumScale = startEb
            .MaximumScale = endEb
        Else
            .MinimumScale = endEb
            .MaximumScale = startEb
            .ReversePlotOrder = True
            .Crosses = xlMaximum
        End If
        .HasTitle = True
        .AxisTitle.Text = strLabel
    End With
    
    SourceRangeColor1 = ActiveChart.SeriesCollection(1).Border.Color
    
    With ActiveSheet.ChartObjects(1)
        .Top = 20
    End With

    If StrComp(str1, "Pe", 1) = 0 Or StrComp(str1, "Be", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then GoTo SkipGraph2
    
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
    For Each myChartOBJ In ActiveSheet.ChartObjects
        With myChartOBJ
            .Left = 200
            .Width = (550 * windowRatio) / windowSize
            .Height = 500 / windowSize
        End With
        With myChartOBJ.Chart.Axes(xlCategory, xlPrimary)
            .MinorTickMark = xlOutside
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .HasMajorGridlines = True
            .MajorUnit = numMajorUnit
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With myChartOBJ.Chart.Axes(xlValue)
            If StrComp(str3, "De", 1) = 0 Then
                .HasTitle = True
                .AxisTitle.Text = "Intensity (arb. units)"
                .Crosses = xlMinimum
            Else
                .HasTitle = True
                If InStr(strTest, "E/eV") > 0 Then
                    If sheetData.Cells(1, 2).Value = "AlKa" Then
                        .AxisTitle.Text = "K counts per sec."
                    Else
                        .AxisTitle.Text = "Intensity (arb. units)"
                    End If
                Else
                    .AxisTitle.Text = "Intensity normalized by Ip (arb. units)"
                End If
            End If
            .MinimumScale = dblMin
            .MaximumScale = dblMax
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With myChartOBJ.Chart.Legend
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
        With myChartOBJ.Chart
            .PlotArea.Width = (((550 * windowRatio) - 40) / windowSize)
            .ChartArea.Border.LineStyle = 0
        End With
    Next
    
    If StrComp(str3, "De", 1) = 0 Then
        ActiveSheet.ChartObjects(2).Activate
        With ActiveChart.Axes(xlValue)
            .MinimumScale = chkMin
            .MaximumScale = chkMax
        End With
    End If
    
    Range(Cells(10, 2), Cells(10, 2)).Interior.Color = SourceRangeColor1
    If StrComp(str1, "Pe", 1) = 0 Or StrComp(str1, "Be", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then
        Range(Cells(10, 1), Cells(10, 1)).Interior.Color = SourceRangeColor1
    End If
    Range(Cells(9 + (imax), 2), Cells(9 + (imax), 2)).Interior.Color = SourceRangeColor1
    strTest = mid$(strSheetGraphName, InStr(strSheetGraphName, "_") + 1, Len(strSheetGraphName) - 6)
    Cells(8 + (imax), 2).Value = strTest + ".xlsx"
    Cells(9 + (imax), 1).Value = str1 + strTest
    Cells(9 + (imax), 2).Value = str2 + strTest
    Cells(9 + (imax), 3).Value = str3 + strTest
    
    If ExistSheet("Sort_" & strSheetDataName) Then
        Application.DisplayAlerts = False
        Worksheets("Sort_" & strSheetDataName).Delete
        Application.DisplayAlerts = True
    End If
End Sub

Sub ElemXPS()
    Dim xpsoffset As Integer
    xpsoffset = 0
CheckElemAgain:
    finTime = Timer

    If StrComp(testMacro, "debug", 1) = 0 Then
        ElemD = ElemX
    Else
        ElemD = Application.InputBox(Title:="Input atomic elements", Prompt:="Example:C,O,Co,etc ... without space!", Default:=ElemD, Type:=2)
    End If
    
    If ElemD <> "False" Then
        If ElemD = "" Then
            Call FitCurve
            'Call GetOut
            Exit Sub
        End If
    Else
        Call GetOut
        Exit Sub
    End If
    
    startTime = Timer
    
    i = 0
    j = 0
    k = 0
    q = 0
    
    If ExistSheet(strSheetXPSFactors) = False Then
        Worksheets.Add().Name = strSheetXPSFactors
        Set sheetXPSFactors = Worksheets(strSheetXPSFactors)
        sheetXPSFactors.Activate
    End If
    
    Fname = direc + "UD.xlsx"
    xpsoffset = 2
    
    If Not WorkbookOpen("UD.xlsx") Then
        graphexist = 0
        Workbooks.Open Fname
        Workbooks("UD.xlsx").Activate
        If Err.Number > 0 Then
            MsgBox "Error in " & Target, vbOKOnly, "Error code: " & Err.Number
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf StrComp(ActiveWorkbook.Name, "UD.xlsx", 1) <> 0 Then
            MsgBox "Error in " & Target
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
        A = Range(Cells(1, 1), Cells(1, 1).Offset(iRow, 5))  '
        If mid$(Cells(1, 4).Value, 1, 1) = "R" Then
            asf = "RSF"  ' Relative Sensitivity factors
        Else
            asf = "ASF"  ' Absolute Sensitivity factors
        End If
        
        If graphexist = 0 Then
            Workbooks("UD.xlsx").Close False
        End If
        sheetXPSFactors.Activate
    Else
        If graphexist = 0 Then
            Workbooks("UD.xlsx").Close False
        End If
        Call GetOut
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    
    Range(Cells(1, 1), Cells(1, 1).Offset((iRow - 1), 5)) = A
    Range(Cells(1, 9), Cells(1, 9).End(xlDown).Offset(, 8)).ClearContents
    Set rng = [A:A]
    imax = Application.CountA(rng)
    
    If imax < 2 Then
        numXPSFactors = 0
        strErrX = "skip"
        Exit Sub
    End If
    
    C = Range(Cells(2, 1), Cells(imax, 6))
    A = Range(Cells(1, 10), Cells(imax, 19))

    k = 0
    tmp = Split(ElemD, ",")
    For i = 0 To UBound(tmp)
        Elem = tmp(i)
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
            ElemD = Replace(ElemD, tmp(i), Elem)
        End If
        k = 0
    Next

    tmp = Split(ElemD, ",")
    
    k = 0
    For i = 0 To UBound(tmp)
        Elem = tmp(i)
        j = 1 + k
        For q = 1 To (imax - 1)
            If C(q, 1) = Elem Then
                A(j, 1) = C(q, 1)   ' Elem
                A(j, 2) = C(q, 2)   ' orbit
                A(j, 3) = C(q, 3)   ' BE
                A(j, 7) = C(q, 6 - xpsoffset)   ' RSF
                j = j + 1
            ElseIf LCase(Elem) = "all" Then
                A(j, 1) = C(q, 1)   ' Elem
                A(j, 2) = C(q, 2)   ' orbit
                A(j, 3) = C(q, 3)   ' BE
                A(j, 7) = C(q, 6 - xpsoffset)   ' RSF
                j = j + 1
            End If
        Next
        If j = 1 + k Then
            If Elem = vbNullString Then
            Else
                TimeCheck = MsgBox(Elem + " : No such an element in database!", vbExclamation, "Input error")
                If StrComp(testMacro, "debug", 1) = 0 Then  ' debugAll code needs this
                    Call GetOut
                    strErrX = "skip"
                    Exit Sub
                Else
                    GoTo CheckElemAgain
                End If
            End If
        End If
        k = j - 1
    Next
    
    numXPSFactors = k
End Sub

Sub ElemAES()
    Dim aesoffset As Integer
    
    If numXPSFactors = 0 Then GoTo SkipXPSnumZero
    
    Worksheets.Add().Name = strSheetPICFactors
    Set sheetPICFactors = Worksheets(strSheetPICFactors)
    sheetPICFactors.Activate
    maxXPSFactor = 0
    b = Range(Cells(1, 1), Cells(1, 1).Offset(numXPSFactors, 6))
    
    For i = 1 To numXPSFactors
        strTest = A(i, 1) + Left(A(i, 2), 2)
        b(i, 1) = strTest
        If Dir(direc + "webCross\") = vbNullString Then
            q = 0
            GoTo SkipElem
        End If
        Fname = direc + "webCross\" + LCase(strTest) + ".txt"
        
        If Len(Dir(Fname)) = 0 Then
            TimeCheck = MsgBox("File Not Found in " + Fname + "!", vbExclamation, "Database error")
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        End If
        
        If Fname = False Then Exit Sub
        fileNum = FreeFile(0)
        Open Fname For Input As #fileNum
        iRow = 1
        Line Input #fileNum, Record
        q = 0
        
        Do Until EOF(fileNum)
            C = Split(Record, vbTab)
            
            If str1 = "Pe" Then         ' XAS mode
                If A(i, 3) < 15 Then    ' if PE < 15 eV, ignore it.
                ElseIf A(i, 3) > 15 And C(0) >= A(i, 3) And q = 0 And A(i, 3) <> 1486.6 Then
                    b(i, 2) = C(0)      ' PE
                    b(i, 3) = C(1)      ' Cross section at PE
                    b(i, 6) = C(4)  ' asymmetric parameter
                    q = 1
                ElseIf A(i, 3) > 15 And C(0) >= A(i, 3) And q = 0 And A(i, 3) = 1486.6 Then
                    b(i, 2) = C(0)      ' PE
                    b(i, 3) = C(1)      ' Cross section at PE
                    b(i, 4) = C(0)      ' Al Ka PE
                    b(i, 5) = C(1)      ' Cross section at Al Ka
                    b(i, 6) = C(4)  ' asymmetric parameter
                    q = 1
                ElseIf C(0) = 1486.6 Then
                    b(i, 4) = C(0)      ' Al Ka PE
                    b(i, 5) = C(1)      ' Cross section at Al Ka
                End If
            Else
                If C(0) >= pe And q = 0 And C(0) <> 1486.6 Then
                    b(i, 2) = C(0)
                    b(i, 3) = C(1)
                    b(i, 6) = C(4)  ' asymmetric parameter
                    q = 1
                ElseIf C(0) >= pe And q = 0 And C(0) = 1486.6 Then
                    b(i, 2) = C(0)
                    b(i, 3) = C(1)
                    b(i, 4) = C(0)
                    b(i, 5) = C(1)
                    b(i, 6) = C(4)  ' asymmetric parameter
                    q = 1
                ElseIf C(0) = 1486.6 Then
                    b(i, 4) = C(0)
                    b(i, 5) = C(1)
                End If
            End If
            
            iRow = iRow + 1
            Line Input #fileNum, Record
        Loop
        Close #fileNum
SkipElem:
        If q = 0 Or StrComp(asf, "ASF", 1) = 0 Then
            b(i, 2) = 0
            b(i, 3) = 1        ' if no data in webcross, multiply this factor !
            b(i, 4) = 0
            b(i, 5) = 1
            b(i, 6) = 1
        End If
    Next
    
    Range(Cells(1, 1), Cells(1, 1).Offset(numXPSFactors, 6)) = b
    
    For i = 1 To numXPSFactors
        If A(i, 7) = "NaN" Then A(i, 7) = 0
        A(i, 2) = A(i, 1) + A(i, 2)
        A(i, 7) = A(i, 7) * b(i, 3) / b(i, 5)
        A(i, 10) = b(i, 6)
    Next
    
    For i = 1 To numXPSFactors
        If A(i, 7) >= maxXPSFactor Then maxXPSFactor = A(i, 7) Else maxXPSFactor = maxXPSFactor
    Next
    
    If Abs(startEb - endEb) > fitLimit Then
        maxXPSFactor = maxXPSFactor * 2
    Else
        maxXPSFactor = maxXPSFactor * 1.2
    End If
    
    For i = 1 To numXPSFactors
        A(i, 8) = dblMin + (A(i, 7) * ((dblMax - dblMin) / (maxXPSFactor)))
        If A(i, 7) = 0 Then
            A(i, 8) = vbNullString
        End If
    Next

    sheetGraph.Activate
    Range(Cells(51, para + 10), Cells((numXPSFactors + 50), para + 19)) = A
    
    If StrComp(Cells(2, 1).Value, "PE", 1) = 0 Then
        If UBound(highpe) > 0 Then      ' higher order or ghost effects
            For i = 1 To UBound(highpe)
                Range(Cells(51 + numXPSFactors * (i), para + 10), Cells((50 + numXPSFactors * (i + 1)), para + 19)) = A
                Cells(40 + i, para + 13).Value = "pe" & i
                Cells(40 + i, para + 14).Value = highpe(i)
            Next
            oriXPSFactors = numXPSFactors
            numXPSFactors = (UBound(highpe) + 1) * numXPSFactors
        End If
    End If

SkipXPSnumZero:
    Worksheets.Add().Name = strSheetAESFactors
    Set sheetAESFactors = Worksheets(strSheetAESFactors)
    sheetAESFactors.Activate
    aesoffset = 0
    
    Fname = direc + "UD.xlsx"
    strAES = "User Defined"
  
    If Not WorkbookOpen("UD.xlsx") Then
        graphexist = 0
        Workbooks.Open Fname
        Workbooks("UD.xlsx").Activate
        If Err.Number > 0 Then
            MsgBox "Error in " & Target, vbOKOnly, "Error code: " & Err.Number
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf StrComp(ActiveWorkbook.Name, "UD.xlsx", 1) <> 0 Then
            MsgBox "Error in " & Target
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
        D = Range(Cells(1, 1), Cells(1, 1).Offset(iRow, 3 + aesoffset))
        
        If graphexist = 0 Then
            Workbooks("UD.xlsx").Close False
        End If
        sheetAESFactors.Activate
    Else
        If graphexist = 0 Then
            Workbooks("UD.xlsx").Close False
        End If
        Call GetOut
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    
    Range(Cells(1, 1), Cells(1, 1).Offset((iRow - 1), 3 + aesoffset)) = D
    Range(Cells(1, 9), Cells(1, 9).End(xlDown).Offset(, 8)).ClearContents
    Set rng = [A:A]
    imax = Application.CountA(rng)
    
    If imax < 2 Then
        numAESFactors = 0
        strErrX = "skip"
        Exit Sub
    End If
    
    C = Range(Cells(2, 1), Cells(imax, 6))
    D = Range(Cells(1, 10), Cells(imax, 18))
    tmp = Split(ElemD, ",")
    k = 0
    For i = 0 To UBound(tmp)
        Elem = tmp(i)
        j = 1 + k
        For q = 1 To (imax - 1)
            If C(q, 1) = Elem Then
                D(j, 1) = C(q, 1)
                D(j, 2) = C(q, 2)
                D(j, 4) = C(q, 3)
                D(j, 7) = C(q, 4 + aesoffset)       ' AES RSF is data at 10keV.
                j = j + 1
            ElseIf LCase(Elem) = "all" Then
                D(j, 1) = C(q, 1)
                D(j, 2) = C(q, 2)
                D(j, 4) = C(q, 3)
                D(j, 7) = C(q, 4 + aesoffset)       ' AES RSF is data at 10keV.
                j = j + 1
            End If
        Next
        k = j - 1
    Next
    
    numAESFactors = k
    maxAESFactor = 0

    For i = 1 To k
        If D(i, 7) = "NaN" Then D(i, 7) = 0
        If D(i, 7) >= maxAESFactor Then maxAESFactor = D(i, 7) Else maxAESFactor = maxAESFactor
    Next
    
    If Abs(startEb - endEb) > fitLimit Then
        maxAESFactor = maxAESFactor * 4
    End If
    
    For i = 1 To numAESFactors
        D(i, 8) = dblMin + (D(i, 7) * ((dblMax - dblMin) / (maxAESFactor)))
        D(i, 2) = D(i, 1) + D(i, 2)
        D(i, 9) = (D(i, 7) * ((chkMin) / (maxAESFactor)))
    Next
    
    sheetGraph.Activate
    Range(Cells((numXPSFactors + 51), para + 10), Cells((numXPSFactors + numAESFactors + 50), para + 18)) = D
End Sub

Sub PlotElem()
    sheetGraph.Activate
    
    If strAna = "FitComp" Then
        maxXPSFactor = Cells(43, para + 10).Value
        maxAESFactor = Cells(44, para + 10).Value
        numChemFactors = Cells(42, para + 12).Value
        numXPSFactors = Cells(43, para + 12).Value
        numAESFactors = Cells(44, para + 12).Value
        If numXPSFactors = 0 And numAESFactors = 0 Then Exit Sub
        
        With ActiveSheet.ChartObjects(1).Chart
            For i = .SeriesCollection.Count To 1 Step -1
                If .SeriesCollection(i).Name = "XPS peaks in BE" Or .SeriesCollection(i).Name = "AES peaks in BE" Then
                    .SeriesCollection(i).Delete
                End If
            Next i
        End With
        
        If ActiveSheet.ChartObjects.Count > 1 Then
            With ActiveSheet.ChartObjects(2).Chart
                For i = .SeriesCollection.Count To 1 Step -1
                    If .SeriesCollection(i).Name = "XPS peaks in KE" Or .SeriesCollection(i).Name = "AES peaks in KE" Then
                        .SeriesCollection(i).Delete
                    End If
                Next i
            End With
        End If
    Else
        Call descriptHidden2
    End If
    
    Dim rngElemBeX As Range
    Dim rngElemBeA As Range
    Dim numFinal As Integer
    numFinal = numXPSFactors + numAESFactors + 50
    Set rngElemBeX = Range(Cells(51, para + 14), Cells((50 + numXPSFactors), para + 14))
    Set rngElemBeA = Range(Cells((numXPSFactors + 51), para + 14), Cells(numFinal, para + 14))

    If numXPSFactors + numAESFactors = 0 Then
        Exit Sub
    ElseIf numXPSFactors = 0 And numAESFactors > 0 Then
        Cells((51 + numXPSFactors), para + 15).FormulaR1C1 = "=RC[-2] - R3C2 - R4C2"        ' KE char from KE
        Cells((51 + numXPSFactors), para + 14).FormulaR1C1 = "=R2C2 - RC[-1]"      ' BE char from KE
        Cells((51 + numXPSFactors), para + 17).FormulaR1C1 = "=R9C3 * ((R41C" & (para + 10) & " + (RC[-1] * (R42C" & (para + 10) & " - R41C" & (para + 10) & ")/R44C" & (para + 10) & ")) - R9C2)"
        Cells((51 + numXPSFactors), para + 18).FormulaR1C1 = "= (RC[-2] * " & (chkMin) & "/R44C" & (para + 10) & ") * R9C3"     ' Sens automatic update
    ElseIf numXPSFactors > 0 And numAESFactors = 0 Then
        Cells(51, para + 15).FormulaR1C1 = "=R2C2 - R3C2 - R4C2 - RC[-3]"     ' KE char from BE
        Cells(51, para + 14).FormulaR1C1 = "=RC[-2]"      ' BE char from BE
        Cells(51, para + 17).FormulaR1C1 = "=R9C3 * ((R41C" & (para + 10) & " + (RC[-1] * (R42C" & (para + 10) & " - R41C" & (para + 10) & ")/R43C" & (para + 10) & ")) - R9C2)"
    Else
        Cells(51, para + 15).FormulaR1C1 = "=R2C2 - R3C2 - R4C2 - RC[-3]"     ' KE char from BE
        Cells((51 + numXPSFactors), para + 15).FormulaR1C1 = "=RC[-2] - R3C2 - R4C2"        ' KE char from KE
        Cells(51, para + 14).FormulaR1C1 = "=RC[-2]"      ' BE char from BE
        Cells((51 + numXPSFactors), para + 14).FormulaR1C1 = "=R2C2 - RC[-1]"      ' BE char from KE
        Cells(51, para + 17).FormulaR1C1 = "=R9C3 * ((R41C" & (para + 10) & " + (RC[-1] * (R42C" & (para + 10) & " - R41C" & (para + 10) & ")/R43C" & (para + 10) & ")) - R9C2)"
        Cells((51 + numXPSFactors), para + 17).FormulaR1C1 = "=R9C3 * ((R41C" & (para + 10) & " + (RC[-1] * (R42C" & (para + 10) & " - R41C" & (para + 10) & ")/R44C" & (para + 10) & ")) - R9C2)"
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
            For i = 1 To UBound(highpe)
                For q = 0 To oriXPSFactors - 1
                    Cells(51 + q + oriXPSFactors * i, para + 11) = Cells(51 + q + oriXPSFactors * i, para + 11).Value & "_" & Cells(40 + i, para + 13).Value
                Next
                
                Cells(51 + oriXPSFactors * i, para + 14).FormulaR1C1 = "=R2C2 - R" & (40 + i) & "C" & (para + 14) & " + RC[-2]"     ' BE higher order from BE
                Cells(51 + oriXPSFactors * i, para + 15).FormulaR1C1 = "=R" & (40 + i) & "C" & (para + 14) & " - R3C2 - R4C2 - RC[-3]"     ' KE char higher order from BE
                Cells(51 + oriXPSFactors * i, para + 17).FormulaR1C1 = "=R9C3 * (R41C" & (para + 10) & " + (RC[-1] * (R42C" & (para + 10) & " - R41C" & (para + 10) & ")/(R43C" & (para + 10) & " * " & (i + 1) & ")))"
                
                If (oriXPSFactors > 1) Then
                    Range(Cells(51 + oriXPSFactors * i, para + 14), Cells((50 + oriXPSFactors * (i + 1)), para + 14)).FillDown
                    Range(Cells(51 + oriXPSFactors * i, para + 15), Cells((50 + oriXPSFactors * (i + 1)), para + 15)).FillDown
                    Range(Cells(51 + oriXPSFactors * i, para + 17), Cells((50 + oriXPSFactors * (i + 1)), para + 17)).FillDown
                End If
            Next
        End If
    End If
    
    ActiveSheet.ChartObjects(1).Activate
    
    If StrComp(str3, "De", 1) = 0 Then
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
        i = 0
        Set pts = .Points
        For Each pt In pts
            i = i + 1
            With pt.DataLabel
                .Text = rngElemBeX.Offset(0, -3).Cells(i).Value
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
        i = 0
        Set pts = .Points
        For Each pt In pts
            i = i + 1
            With pt.DataLabel
                .Text = rngElemBeA.Offset(0, -3).Cells(i).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 12 / Sqr(windowSize)
            End With
        Next
        End With
    End If
    
    If ActiveChart.HasLegend = True Then
        With ActiveSheet.ChartObjects(1).Chart
            For i = .SeriesCollection.Count To 1 Step -1
                If .SeriesCollection(i).Name = "XPS peaks in BE" Or .SeriesCollection(i).Name = "AES peaks in BE" Then
                    .Legend.LegendEntries(i).Delete
                End If
            Next i
        End With
    End If
    
    If StrComp(str1, "Pe", 1) = 0 Or StrComp(str1, "Be", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then Exit Sub
    If ActiveSheet.ChartObjects.Count = 1 Then Exit Sub
        ActiveSheet.ChartObjects(2).Activate
        If StrComp(str3, "De", 1) = 0 Then GoTo AESmode2
        
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
        i = 0
        Set pts = .Points
        For Each pt In pts
            i = i + 1
            With pt.DataLabel
                .Text = rngElemBeX.Offset(0, -3).Cells(i).Value
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
        i = 0
        Set pts = .Points
        For Each pt In pts
            i = i + 1
            With pt.DataLabel
                .Text = rngElemBeA.Offset(0, -3).Cells(i).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 12 / Sqr(windowSize)
            End With
        Next
        End With
    End If
    
    If ActiveChart.HasLegend = True Then
        With ActiveSheet.ChartObjects(2).Chart
            For i = .SeriesCollection.Count To 1 Step -1
                If .SeriesCollection(i).Name = "XPS peaks in KE" Or .SeriesCollection(i).Name = "AES peaks in KE" Then
                    .Legend.LegendEntries(i).Delete
                End If
            Next i
        End With
    End If
End Sub

Sub PlotChem()
    Set sheetGraph = Worksheets("Graph_" + strSheetDataName)
    sheetGraph.Activate
    If LCase(Cells(10, 1).Value) = "pe" Then
        Cells(10, 3).Value = "Ab"   'str3
    Else
        Cells(10, 3).Value = "In"   'str3
    End If
    Call GetOut
    End
End Sub

Sub GetCompare()
    If StrComp(TimeCheck, "yes", 1) = 0 Then TimeCheck = vbNullString
    Worksheets(strSheetGraphName).Activate
    i = k
    j = 0
    If Cells(51, para + 9).Value = vbNullString Then
        p = 2   ' XPS and AES modes without any factors plots only data.
    ElseIf Cells(42, para + 12).Value > 0 Then
        p = 5   ' XPS mode with chemical shifts plots Data, XPS, AES, and Chem factors.
    ElseIf StrComp(Cells(2, 1).Value, "KE shifts", 1) = 0 Then
        p = 3   ' AES mode plots Data and AES factors.
    Else
        p = 4   ' XPS mode without chemical shifts plots Data, XPS, and AES factors.
    End If
    
    If StrComp(Cells(2, 1).Value, "PE shifts", 1) = 0 Then
        q = 1 'for XAS mode
        str1 = "Pe"
        str2 = "Sh"
        str3 = "Ab"
    ElseIf StrComp(Cells(2, 1).Value, "PE", 1) = 0 Then
        q = 2 'for XPS mode
        If StrComp(Cells(10, 1).Value, "Be", 1) = 0 Then
            str1 = "Be"
            str2 = "Sh"
        Else
            str1 = "Ke"
            str2 = "Be"
        End If
        str3 = "In"
    ElseIf StrComp(Cells(2, 1).Value, "KE shifts", 1) = 0 Then
        q = 3 ' for AES mode
        If StrComp(Cells(1, 1).Value, "AES elec.", 1) = 0 Then
            str1 = "Ke"
            str2 = "Ae"
            str3 = "De"
        End If
    ElseIf StrComp(Cells(2, 1).Value, "Shifts", 1) = 0 Then
        q = 4 'for DC mode
        str1 = "Po"
        str2 = "Sh"
        str3 = "Ab"
    Else
        Call GetOut
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
            
    ChDrive mid$(ActiveWorkbook.Path, 1, 1)
    ChDir ActiveWorkbook.Path
    OpenFileName = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", Title:="Please select a file", MultiSelect:=True)
    
    If IsArray(OpenFileName) Then
        If UBound(OpenFileName) + ncomp - (ncomp - k) > CInt(para / 3) Then
            TimeCheck = MsgBox("Stop a comparison because you select too many files: " & (UBound(OpenFileName) + ncomp - (ncomp - k)) & " over the total limit: " & CInt(para / 3), vbExclamation)
            Cells(1, 4 + (k * 3)).Value = vbNullString
            Call GetOut
            If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
        ElseIf UBound(OpenFileName) > 1 Then
        End If
        
        Application.Calculation = xlCalculationManual
        Call EachComp
        Application.Calculation = xlCalculationAutomatic
        
        Workbooks(wb).Sheets(strSheetGraphName).Activate
        If Not (i - k) = 0 Then Call offsetmultiple
        If ncomp > i Then
            Cells(45, para + 10).Value = ncomp
        Else
            Cells(45, para + 10).Value = i
        End If
    Else
        TimeCheck = MsgBox("Stop a comparison; no file selected.", vbExclamation)
        Cells(1, 4 + (k * 3)).Value = vbNullString
        Call GetOut
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    
    Cells(1, 4 + (k * 3)).Value = vbNullString
    
    If ExistSheet("samples") Then
        Results = "i" & i & "k" & k
        Call CombineLegend
    End If
    
    Call GetOut
End Sub

Sub GetOut()
    Call Delsheets
    
    If Not Cells(8, 101).Value = 0 Then
        If StrComp(TimeCheck, "yes", 1) = 0 Then TimeCheck = "yes1"
            'GoTo ResultFit
    Else
        If ExistSheet(strSheetFitName) And strAna = "FitRatioAnalysis" Then
            Worksheets(strSheetFitName).Activate
        ElseIf ExistSheet(strSheetGraphName) Then
            Worksheets(strSheetGraphName).Activate
        End If
    End If
    
    endTime = Timer

    Cells(1, 1).Select
    Application.ScreenUpdating = True
    
    If StrComp("Fit", mid$(ActiveSheet.Name, 1, 3)) = 0 And IsNumeric(TimeCheck) = False Then 'Cells(9 + sftfit2, 1).Value = "Solve LSM" Then
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
                    If Cells(18, 101).Value < 0.000001 Then
                        TimeCheck = MsgBox("Fitting does not work properly, because avaraged In data is less than 1E-6!")
                    ElseIf Cells(18, 101).Value > 1E+29 Then
                        TimeCheck = MsgBox("Fitting does not work properly, because avaraged In data is more than 1E+29!")
                    End If
                Else
                End If
            End If
        End If
    End If
    
    If StrComp(TimeCheck, "yes", 1) = 0 Then          ' graph processes.
        gamma = (finTime - iniTime) - (TimeC2 - TimeC1)
        lambda = endTime - startTime
        MsgBox "Progress time: " & Application.Text(gamma, "0.00") & "," & Application.Text(lambda, "0.00") & ".", vbInformation
    ElseIf StrComp(TimeCheck, "yes1", 1) = 0 Then     ' this is for fitting process.
        gamma = (endTime - iniTime) - (finTime - startTime)
        lambda = TimeC2 - TimeC1
        MsgBox "Progress time: " & Application.Text(gamma, "0.00") & "," & Application.Text(lambda, "0.00") & ".", vbInformation
    ElseIf StrComp(TimeCheck, "yes2", 1) = 0 Then     ' this is for the other process.
        gamma = finTime - iniTime
        MsgBox "Progress time: " & Application.Text(gamma, "0.00") & ".", vbInformation
    ElseIf StrComp(TimeCheck, "yes3", 1) = 0 Then     ' this is for the other process.
        gamma = endTime - startTime
        MsgBox "Progress time: " & Application.Text(gamma, "0.00") & ".", vbInformation
    End If
    
    testMacro = vbNullString
    
    Application.DisplayAlerts = False
    If Len(ActiveWorkbook.Path) < 2 Then
        Application.Dialogs(xlDialogSaveAs).Show
    Else
        On Error GoTo Error1
        ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path + "\" + wb, FileFormat:=51
    End If
    Application.DisplayAlerts = True
    strErr = "skip"
    Exit Sub
Error1:
    MsgBox Error(Err)
    wb = mid$(ActiveWorkbook.Name, 1, InStr(ActiveWorkbook.Name, ".") - 1) + "_bk.xlsx"
    ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path + "\" + wb, FileFormat:=51
    Err.Clear
    strErr = "skip"
    Resume Next
End Sub

Sub FitInitial()
    If StrComp(strAna, "ana", 1) = 0 Or StrComp(str1, "Pe", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then
        Worksheets(strSheetGraphName).Activate
        numData = Cells(41, para + 12).Value '((Cells(6, 2).Value - Cells(5, 2).Value) / Cells(7, 2).Value) + 1
        Gnum = Cells(45, para + 12).Value
        Set dataBGraph = Range(Cells(20 + numData, 2), Cells(20 + numData, 2).Offset(numData - 1, 1))
        Set dataKeGraph = Range(Cells(20, 1), Cells(20 + numData, 1).Offset(numData - 1, 0))
        Set dataIntGraph = dataKeGraph.Offset(, 2)
        Call scalecheck
        If StrComp(str1, "Pe", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then
            Cells(10, 3).Value = "Ab"
        Else
            Cells(10, 3).Value = "In"
        End If
        If ExistSheet(strSheetFitName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetFitName).Delete
            Application.DisplayAlerts = True
        End If
    ElseIf StrComp(str3, "De", 1) = 0 Then
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
    
    en = dataBGraph
    For i = 1 To numData
        en(i, 1) = Round(en(i, 1), 3)   ' This makes round en off to third decimal places.
    Next
    
    Range(Cells(21 + sftfit, 1), Cells((numData + 20 + sftfit), 2)).Value = en
    Set dataBGraph = Range(Cells(21 + sftfit, 1), Cells((numData + 20 + sftfit), 2))
    Set dataKGraph = Range(Cells(21 + sftfit, 1), Cells((numData + 20 + sftfit), 1))
    Set dataKeGraph = Range(Cells(11, 103), Cells(15, 104))
    
    If StrComp(str1, "Pe", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then
        Cells(11 + sftfit2, 2).Value = Cells(21 + sftfit, 1).Value
        Cells(12 + sftfit2, 2).Value = Cells(numData + 20 + sftfit, 1).Value
    Else
        Cells(11 + sftfit2, 2).Value = Cells(numData + 20 + sftfit, 1).Value
        Cells(12 + sftfit2, 2).Value = Cells(21 + sftfit, 1).Value
    End If
    
    ActiveWorkbook.Charts.Add Before:=Worksheets(Worksheets.Count)  ' it makes no additional series in plot
    
    If Abs(startEb - endEb) < fitLimit Then
        ActiveChart.ChartType = xlXYScatter
    Else
        ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    End If
    ActiveChart.SetSourceData Source:=dataBGraph, PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetFitName
    ActiveChart.SeriesCollection(1).Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C2"
    
    ActiveWorkbook.Charts.Add Before:=Worksheets(Worksheets.Count)
    If Abs(startEb - endEb) < fitLimit Then
        ActiveChart.ChartType = xlXYScatter
    Else
        ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    End If

    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetFitName
    ActiveChart.SeriesCollection.NewSeries
    
    With ActiveChart.SeriesCollection(1)
        '.ChartType = xlXYScatterLinesNoMarkers
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
    For i = k To 2 Step -1
        ActiveChart.SeriesCollection(i).Delete
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
            If str1 = "Pe" Then
                .AxisTitle.Text = "Photon energy (eV)"
                .MinimumScale = startEb
                .MaximumScale = endEb
                .ReversePlotOrder = False
                .Crosses = xlMinimum
            ElseIf str1 = "Po" Then
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
            .MajorUnit = numMajorUnit
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
    Dim obchk As String
    Dim obchk2 As String
    Dim dbltchk As Single
    Dim dblt As Integer
    
    dblt = 0
    dbltchk = 0
    sheetGraph.Activate
    numXPSFactors = Cells(43, para + 12).Value
    C = Range(Cells(51, para + 10), Cells((51 + numXPSFactors), para + 12)) ' peak name and BE
    D = Range(Cells(51, para + 16), Cells((51 + numXPSFactors), para + 19)) ' Amp and sensitivity
    sheetFit.Activate
    Set dataFit = Range(Cells(1, 5), Cells(15 + sftfit2 + 3, (numXPSFactors + 5)))
    A = dataFit
    
    For i = numXPSFactors To 1 Step -1
        If StrComp(str1, "Pe", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then
            If C(i, 3) > Cells(startR, 1).Value And C(i, 3) < Cells(endR, 1).Value Then
                j = j + 1
                A(1, j) = C(i, 2)
                
                If Len(C(i, 2)) - Len(C(i, 1)) > 2 Then
                    obchk = mid$(C(i, 2), Len(C(i, 1)) + 2, 2)
                    Debug.Print obchk
                    If StrComp(obchk, "p3", 1) = 0 Then
                        dbltchk = C(i, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "p3", 1) = 0 And StrComp(obchk, "p1", 1) = 0 Then
                        A(19, dblt) = "(2;"
                        A(19, j) = "1)"
                        A(20, dblt) = "["
                        If dbltchk <= C(i, 3) Then
                            dbltchk = C(i, 3) - dbltchk
                            A(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C(i, 3)
                            A(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf StrComp(obchk, "d5", 1) = 0 Then
                        dbltchk = C(i, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "d5", 1) = 0 And StrComp(obchk, "d3", 1) = 0 Then
                        A(19, dblt) = "(3;"
                        A(19, j) = "2)"
                        A(20, dblt) = "["
                        If dbltchk <= C(i, 3) Then
                            dbltchk = C(i, 3) - dbltchk
                            A(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C(i, 3)
                            A(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf StrComp(obchk, "f7", 1) = 0 Then
                        dbltchk = C(i, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "f7", 1) = 0 And StrComp(obchk, "f5", 1) = 0 Then
                        A(19, dblt) = "(4;"
                        A(19, j) = "3)"
                        A(20, dblt) = "["
                        If dbltchk <= C(i, 3) Then
                            dbltchk = C(i, 3) - dbltchk
                            A(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C(i, 3)
                            A(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf dbltchk <> 0 And obchk2 <> vbNullString Then
                        A(19, dblt) = vbNullString
                        A(20, dblt) = vbNullString
                        dbltchk = 0
                        obchk2 = vbNullString
                        dblt = 0
                    End If
                End If
                
                A(2, j) = C(i, 3)
                A(6, j) = D(i, 2)
                A(9 + sftfit2, j) = D(i, 1)
                A(7 + sftfit2, j) = D(i, 4) ' beta
            End If
        Else
            If C(i, 3) < Cells(startR, 1).Value And C(i, 3) > Cells(endR, 1).Value Then
                j = j + 1
                A(1, j) = C(i, 2)   ' peak name
                
                If Len(C(i, 2)) - Len(C(i, 1)) > 2 Then
                    obchk = mid$(C(i, 2), Len(C(i, 1)) + 2, 2)
                    Debug.Print obchk
                    If StrComp(obchk, "p3", 1) = 0 Then
                        dbltchk = C(i, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "p3", 1) = 0 And StrComp(obchk, "p1", 1) = 0 Then
                        A(19, dblt) = "(2;"
                        A(19, j) = "1)"
                        A(20, dblt) = "["
                        If dbltchk <= C(i, 3) Then
                            dbltchk = C(i, 3) - dbltchk
                            A(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C(i, 3)
                            A(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf StrComp(obchk, "d5", 1) = 0 Then
                        dbltchk = C(i, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "d5", 1) = 0 And StrComp(obchk, "d3", 1) = 0 Then
                        A(19, dblt) = "(3;"
                        A(19, j) = "2)"
                        A(20, dblt) = "["
                        If dbltchk <= C(i, 3) Then
                            dbltchk = C(i, 3) - dbltchk
                            A(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C(i, 3)
                            A(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf StrComp(obchk, "f7", 1) = 0 Then
                        dbltchk = C(i, 3)
                        obchk2 = obchk
                        dblt = j
                    ElseIf StrComp(obchk2, "f7", 1) = 0 And StrComp(obchk, "f5", 1) = 0 Then
                        A(19, dblt) = "(4;"
                        A(19, j) = "3)"
                        A(20, dblt) = "["
                        If dbltchk <= C(i, 3) Then
                            dbltchk = C(i, 3) - dbltchk
                            A(20, j) = dbltchk & "]"
                        Else
                            dbltchk = dbltchk - C(i, 3)
                            A(20, j) = "n" & dbltchk & "]"
                        End If
                        dbltchk = 0
                        obchk2 = vbNullString
                    ElseIf dbltchk <> 0 And obchk2 <> vbNullString Then
                        A(19, dblt) = vbNullString
                        A(20, dblt) = vbNullString
                        dbltchk = 0
                        obchk2 = vbNullString
                        dblt = 0
                    End If
                End If
                
                A(2, j) = C(i, 3)   ' BE
                A(6, j) = D(i, 2)   ' Amp.
                A(9 + sftfit2, j) = D(i, 1) ' sensitivity
                A(7 + sftfit2, j) = D(i, 4) ' beta
            End If
        End If
    Next
    
    Range(Cells(1, 5), Cells(15 + sftfit2 + 3, (numXPSFactors + 5))) = A
    Range(Cells(1, 4), Cells(15 + sftfit2, 4)).Interior.Color = RGB(77, 208, 225)  '33
    Range(Cells(15 + sftfit2 + 1, 4), Cells(15 + sftfit2 + 3, 4)).Interior.Color = RGB(176, 190, 197)
    Range(Cells(1, 5), Cells(15 + sftfit2, 5)).Interior.Color = RGB(178, 235, 242) '34
    Range(Cells(15 + sftfit2 + 1, 5), Cells(15 + sftfit2 + 3, 5)).Interior.Color = RGB(207, 216, 220)
    
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

Sub FitRange()
    Set rng = [A:A]
    
    strSheetGraphName = "Graph_" + strSheetDataName
    If StrComp(mid$(strTest, 8, 5), "range", 1) = 0 Then
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
    
    If IsEmpty(Cells(2, 103).Value) Then
        If Cells(15 + sftfit2, 2).Value > 1 Then
            Cells(2, 103).Value = 10       ' max FWHM1 limit
            Cells(3, 103).Value = 0.5       ' min FWHM1 limit
            Cells(4, 103).Value = 10       ' max FWHM2 limit
            Cells(5, 103).Value = 0.5       ' min FWHM2 limit
            Cells(6, 103).Value = 0.999       ' max shape limit
            Cells(7, 103).Value = 0.001       ' min shape limit
            If str1 = "Pe" Then             ' additional BE step
                Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (4)
                Cells(2, 103).Value = 5       ' max FWHM1 limit
            Else
                Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (20)
            End If
        Else      ' grating #1 or 0
            Cells(2, 103).Value = 2       ' max FWHM1 limit
            Cells(3, 103).Value = 0.1       ' min FWHM1 limit
            Cells(4, 103).Value = 2       ' max FWHM2 limit
            Cells(5, 103).Value = 0.1       ' min FWHM2 limit
            Cells(6, 103).Value = 0.999       ' max shape limit
            Cells(7, 103).Value = 0.001       ' min shape limit
            Cells(10, 101).Value = 20          ' average points for poly BG
            If str1 = "Pe" Then             ' additional BE step
                Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (20)
                Cells(2, 103).Value = 1       ' max FWHM1 limit
            Else
                Cells(8, 103).Value = Abs(Cells(7, 101).Value - Cells(6, 101).Value) / (100)
            End If
        End If
    ElseIf Cells(8, 101).Value > 0 Then      ' fit done
    End If
    
    pe = Cells(12, 101).Value
    wf = Cells(13, 101).Value
    char = Cells(14, 101).Value
    ns = Cells(10, 101).Value
    
    If StrComp(Worksheets(strSheetGraphName).Cells(10, 3).Value, "Ab", 1) = 0 And StrComp(Worksheets(strSheetGraphName).Cells(10, 1).Value, "Pe", 1) = 0 Then
        str1 = "Pe"
    ElseIf StrComp(Worksheets(strSheetGraphName).Cells(10, 1).Value, "Po", 1) = 0 Then
        str1 = "Po"
        Cells(10, 101).Value = 3        ' Average points for edges of BG
    End If
    
    If Abs(startEb - endEb) > fitLimit Then
        If StrComp(testMacro, "debug", 1) = 0 Then  ' debug mode skip fitting in the specific range.
            TimeCheck = 0
            Call GetOutFit
            strErrX = "skip"
            Exit Sub
        End If
        startTime = Timer
        ElemD = Application.InputBox(Title:="Specify the fitting range", Prompt:="Input BE energy: 320-350eV", Default:="320-350eV", Type:=2)
        finTime = Timer
        If ElemD = "False" Or Len(ElemD) = 0 Then
            TimeCheck = 0
            Call GetOutFit
            strErrX = "skip"
            Exit Sub
        Else
            tmp = Split(ElemD, "-")
            If IsNumeric(mid$(tmp(1), 1, Len(tmp(1)) - 2)) = True Then
                If mid$(tmp(1), 1, Len(tmp(1)) - 2) < startEb And mid$(tmp(1), 1, Len(tmp(1)) - 2) > endEb Then
                    startEb = mid$(tmp(1), 1, Len(tmp(1)) - 2)
                Else
                    'GoTo GetOutFit
                End If
            Else
                TimeCheck = MsgBox("BE range format is not appropriate!")
                Call GetOutFit
                strErrX = "skip"
                Exit Sub
            End If
            
            If IsNumeric(tmp(0)) = True Then
                If tmp(0) < startEb And tmp(0) > endEb Then
                    endEb = tmp(0)
                ElseIf tmp(0) > startEb Then
                    startEb = tmp(0)
                    endEb = mid$(tmp(1), 1, Len(tmp(1)) - 2)
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
            ElseIf Abs(startEb - endEb) <= 20 Then
                numMajorUnit = 1 * windowSize
            End If
    
            If StrComp(str1, "Pe", 1) = 0 Or StrComp(str3, "De", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then
                startEb = Application.Floor(startEb, numMajorUnit)
            ElseIf startEb > 0 Then
                startEb = Application.Ceiling(startEb, numMajorUnit)
            Else
                startEb = Application.Floor(startEb, (-1 * numMajorUnit))
            End If

            If StrComp(str1, "Pe", 1) = 0 Or StrComp(str3, "De", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then
                endEb = Application.Ceiling(endEb, numMajorUnit)
            ElseIf endEb > 0 Then
                endEb = Application.Floor(endEb, numMajorUnit)
            Else
                endEb = Application.Ceiling(endEb, (-1 * numMajorUnit))
            End If
            
            Cells(6, 101).Value = startEb
            Cells(7, 101).Value = endEb
            Cells(11, 100).Value = "majorUnit"
            Cells(11, 101).Value = numMajorUnit
            If Abs(startEb - endEb) / Abs(Cells(22 + sftfit, 1).Value - Cells(21 + sftfit, 1).Value) < 30 Then
                Cells(10, 101).Value = 3     ' Average # points for Solver around startR and endR points
            ElseIf Abs(startEb - endEb) / Abs(Cells(22 + sftfit, 1).Value - Cells(21 + sftfit, 1).Value) < 60 Then
                Cells(10, 101).Value = 5
            Else
                Cells(10, 101).Value = 10
            End If
            
            If str1 = "Pe" Then             ' additional BE step
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
    en = Range(Cells(21 + sftfit, 1), Cells((numData + 20 + sftfit), 1))
    k = 0
    j = 0
    
    If StrComp(str1, "Pe", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then
        For i = 1 To numData - 1
            If Cells(11 + sftfit2, 2) >= en(i, 1) And Cells(11 + sftfit2, 2) < en((i + 1), 1) Then
                startR = i + 20 + sftfit
            End If
        Next
        For i = 2 To numData
            If Cells(12 + sftfit2, 2) <= en(i, 1) And Cells(12 + sftfit2, 2) > en((i - 1), 1) Then
                endR = i + 20 + sftfit
            End If
        Next
    Else
        For i = 1 To numData - 1
            If Cells(12 + sftfit2, 2) <= en(i, 1) And Cells(12 + sftfit2, 2) > en((i + 1), 1) Then
                startR = i + 20 + sftfit
            End If
        Next
        For i = 2 To numData
            If Cells(11 + sftfit2, 2) >= en(i, 1) And Cells(11 + sftfit2, 2) < en((i - 1), 1) Then
                endR = i + 20 + sftfit
            End If
        Next
    End If
    
    If startR < 1 Or CStr(startR) = vbNullString Then
        startR = 21 + sftfit
        Cells(12 + sftfit2, 2).Value = Cells(21 + sftfit, 1).Value
    End If
    
    If endR > numData + 20 + sftfit Or endR < startR Or CStr(endR) = vbNullString Then
        endR = numData + 20 + sftfit
        Cells(11 + sftfit2, 2).Value = Cells(numData + 20 + sftfit, 1).Value
    End If
    
    numDataN = endR - startR + 1
    C = Range(Cells(startR, 2), Cells(endR, 2))
    A = Range(Cells(startR, 3), Cells(endR, 3))
    A(numDataN, 1) = C(numDataN, 1)
    A((numDataN - 1), 1) = C(numDataN, 1)
    
    If IsEmpty(Cells(1, 101).Value) = False Then    ' range > fitLimit eV
        Cells(2, 101).Value = Application.Min(C)
        Cells(3, 101).Value = Application.Max(C)
        dblMin = Cells(2, 101).Value - ((Cells(3, 101).Value - Cells(2, 101).Value) / 100)
        dblMax = Cells(3, 101).Value + ((Cells(3, 101).Value - Cells(2, 101).Value) / 10)
        i = 0
        For Each myChartOBJ In ActiveSheet.ChartObjects
            i = i + 1
            If i = 1 Then
                With myChartOBJ.Chart.Axes(xlCategory, xlPrimary)
                    .MinimumScale = Cells(7, 101).Value
                    .MaximumScale = Cells(6, 101).Value
                    .MajorUnit = Cells(11, 101).Value
                End With
                With myChartOBJ.Chart.Axes(xlValue)
                    .MinimumScale = dblMin
                    .MaximumScale = dblMax
                End With
            ElseIf i = 2 Then
                With myChartOBJ.Chart.Axes(xlCategory, xlPrimary)
                    .MinimumScale = Cells(7, 101).Value
                    .MaximumScale = Cells(6, 101).Value
                    .MajorUnit = Cells(11, 101).Value
                End With
                Exit For
            End If
        Next
    End If
    
    strTest = LCase(mid$(Cells(1, 1).Value, 1, 1))
    strLabel = LCase(mid$(Cells(1, 2).Value, 1, 1))
    
    If strTest = LCase(mid$(Cells(16, 100).Value, 1, 1)) And strLabel = LCase(mid$(Cells(16, 101).Value, 1, 1)) Then
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

Sub FitCurve()
    Application.Calculation = xlCalculationManual
    
    If StrComp(mid$(strTest, 1, 6), "Do fit", 1) = 0 Then
    Else
        Call FitInitial
        Exit Sub
    End If

    If IsEmpty(Cells(19, 101).Value) Then
        MsgBox "VBA code version analyzed in the sheet is too old, regenerate the fit sheet from graph sheet again.", vbInformation
        Call GetOut
        Exit Sub
    End If

    Call FitRange
    If strErrX = "skip" Then Exit Sub
    
    Call SolverSetup           ' SolverSetup2 for accurate
    
    If StrComp(strTest, "t", 1) <> 0 And StrComp(strLabel, "t", 1) <> 0 And StrComp(strTest, "e", 1) <> 0 Then
        Range("DG31").CurrentRegion.ClearContents
        Range("DE31").CurrentRegion.ClearContents
    End If
    
    If Cells(11 + sftfit2, 2).Value < 0 And Cells(12 + sftfit2, 2).Value > 0 And (Cells(12 + sftfit2, 2).Value - Cells(11 + sftfit2, 2).Value) < 6 Then
        Call FitEF
        Call GetOutFit
        Exit Sub
    ElseIf StrComp(strTest, "p", 1) = 0 Then
        Call PolynominalBG
    ElseIf StrComp(strTest, "a", 1) = 0 Then
        Call TangentArcBG
    ElseIf StrComp(strTest, "t", 1) = 0 Then
        Call TougaardBG
    ElseIf StrComp(strTest, "v", 1) = 0 Then
        Call VictoreenBG
    ElseIf StrComp(strTest, "e", 1) = 0 Then
        Call FitEF
        Call GetOutFit
        Exit Sub
    Else
        Call ShirleyBG      ' Solver mode; ShirleyBG2 for iteration mode.
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
    For i = k To 2 Step -1
        ActiveChart.SeriesCollection(i).Delete
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
    ActiveSheet.Calculate
AsymIteration:
    A = Cells(9 + sftfit2, 2).Value
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
            'Call descriptInitialFit
            'sheetGraph.Activate
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
    
    Call SolverSetup           ' SolverSetup2 for accurate

    If StrComp(Cells(1, 1).Value, "Polynominal", 1) = 0 Then
        SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(7 + sftfit2 - 2, (4 + j)))
        ' Error here : No Solver reference in VBE - Tools - References - Solver checked.

        For k = 2 To 5
            If Cells(k, 2).Font.Bold = "True" Then
                SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
            End If
        Next
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
        SolverOk SetCell:=Cells(9 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(7 + sftfit2 - 2, (4 + j)))  ' active
        For k = 2 To 11
            If Cells(k, i).Font.Bold = "True" Then
                SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
            End If
        Next
    End If
    
    If StrComp(str1, "Pe", 1) = 0 Or StrComp(str1, "Po", 1) = 0 Then
        SolverAdd CellRef:=Range(Cells(2, 5), Cells(2, (4 + j))), Relation:=3, FormulaText:=Cells(startR, 1).Value
        SolverAdd CellRef:=Range(Cells(2, 5), Cells(2, (4 + j))), Relation:=1, FormulaText:=Cells(endR, 1).Value
    Else
        SolverAdd CellRef:=Range(Cells(2, 5), Cells(2, (4 + j))), Relation:=1, FormulaText:=Cells(startR, 1).Value
        SolverAdd CellRef:=Range(Cells(2, 5), Cells(2, (4 + j))), Relation:=3, FormulaText:=Cells(endR, 1).Value
    End If
    
    SolverAdd CellRef:=Range(Cells(4, 5), Cells(4, (4 + j))), Relation:=1, FormulaText:=Cells(2, 103).Value    ' width1 max
    SolverAdd CellRef:=Range(Cells(4, 5), Cells(4, (4 + j))), Relation:=3, FormulaText:=Cells(3, 103).Value    ' width1 min
    SolverAdd CellRef:=Range(Cells(6, 5), Cells(6, (4 + j))), Relation:=3, FormulaText:=0             ' minimum amplitude (1E-6/dblMin)
    SolverAdd CellRef:=Range(Cells(3, 5), Cells(3, (4 + j))), Relation:=2, FormulaText:=0
        
    For i = 1 To j
        If Cells(7, (4 + i)).Value = 0 Or Cells(7, (4 + i)).Value = "Gauss" Then ' G
            SolverAdd CellRef:=Cells(7, (4 + i)), Relation:=2, FormulaText:=0
            SolverAdd CellRef:=Range(Cells(8, (4 + i)), Cells(10, (4 + i))), Relation:=2, FormulaText:=0
            SolverAdd CellRef:=Cells(5, (4 + i)), Relation:=2, FormulaText:=0        ' width2
        ElseIf Cells(7, (4 + i)).Value = 1 Or Cells(7, (4 + i)).Value = "Lorentz" Then
            SolverAdd CellRef:=Cells(7, (4 + i)), Relation:=2, FormulaText:=1
            If Cells(7, (4 + i)).Font.Italic = "True" Then  'Doniach-Sunjic (DS)
                SolverAdd CellRef:=Cells(8, (4 + i)), Relation:=1, FormulaText:=1
                SolverAdd CellRef:=Cells(8, (4 + i)), Relation:=3, FormulaText:=0
                SolverAdd CellRef:=Cells(5, (4 + i)), Relation:=2, FormulaText:=0        ' width2
                SolverAdd CellRef:=Cells(10, (4 + i)), Relation:=2, FormulaText:=0
            Else    ' L
                SolverAdd CellRef:=Range(Cells(8, (4 + i)), Cells(10, (4 + i))), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(5, (4 + i)), Relation:=2, FormulaText:=0        ' width2
            End If
        Else
            If mid(Cells(11, (4 + i)).Value, 1, 1) = "E" Then
                SolverAdd CellRef:=Range(Cells(9, (4 + i)), Cells(10, (4 + i))), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(5, (4 + i)), Relation:=2, FormulaText:=0        ' width2
            ElseIf mid(Cells(11, (4 + i)).Value, 1, 1) = "T" Then       ' MultiPak Asymmetric GL with exp tail
                SolverAdd CellRef:=Cells(10, (4 + i)), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(5, (4 + i)), Relation:=2, FormulaText:=0        ' width2
                SolverAdd CellRef:=Cells(8, (4 + i)), Relation:=1, FormulaText:=3         ' max a : Tail scale
                SolverAdd CellRef:=Cells(8, (4 + i)), Relation:=3, FormulaText:=0         ' min a : Tail scale
                SolverAdd CellRef:=Cells(9, (4 + i)), Relation:=1, FormulaText:=Abs(Cells(6, 101).Value - Cells(7, 101).Value)          ' max b : Tail length
                SolverAdd CellRef:=Cells(9, (4 + i)), Relation:=3, FormulaText:=1         ' min b : Tail length
            ElseIf mid(Cells(11, (4 + i)).Value, 1, 1) = "D" Then
                SolverAdd CellRef:=Cells(10, (4 + i)), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(8, (4 + i)), Relation:=1, FormulaText:=1         ' max a
                SolverAdd CellRef:=Cells(8, (4 + i)), Relation:=3, FormulaText:=0         ' min a
                SolverAdd CellRef:=Cells(9, (4 + i)), Relation:=1, FormulaText:=1         ' max b
                SolverAdd CellRef:=Cells(9, (4 + i)), Relation:=3, FormulaText:=0         ' min b
            ElseIf mid(Cells(11, (4 + i)).Value, 1, 1) = "U" Then
                SolverAdd CellRef:=Cells(10, (4 + i)), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(8, (4 + i)), Relation:=1, FormulaText:=1         ' max a
                SolverAdd CellRef:=Cells(8, (4 + i)), Relation:=3, FormulaText:=0         ' min a
                SolverAdd CellRef:=Cells(9, (4 + i)), Relation:=1, FormulaText:=1         ' max b
                SolverAdd CellRef:=Cells(9, (4 + i)), Relation:=3, FormulaText:=0         ' min b
            Else
                SolverAdd CellRef:=Range(Cells(8, (4 + i)), Cells(10, (4 + i))), Relation:=2, FormulaText:=0
                SolverAdd CellRef:=Cells(5, (4 + i)), Relation:=1, FormulaText:=Cells(4, 103).Value        ' width2 max
                SolverAdd CellRef:=Cells(5, (4 + i)), Relation:=3, FormulaText:=Cells(5, 103).Value         ' width2 min
            End If

            SolverAdd CellRef:=Cells(7, (4 + i)), Relation:=1, FormulaText:=Cells(6, 103).Value         ' max shape
            SolverAdd CellRef:=Cells(7, (4 + i)), Relation:=3, FormulaText:=Cells(7, 103).Value         ' min shape
        End If
    Next

    For i = 5 To (4 + j)
        For k = 1 To 9
            If Cells((k + 1), i).Font.Bold = "True" Then
                SolverAdd CellRef:=Cells((k + 1), i), Relation:=2, FormulaText:=Cells((k + 1), i)
            End If
        Next
    Next
    
    strErr = "Amp. ratio format error: (i; j; k) and i,j,k > 0"
    iRow = 1
    For i = 5 To (4 + j)
        If Not Cells(14 + sftfit2, i) = vbNullString Then
            If iRow = 1 And mid$(Cells(14 + sftfit2, i), 1, 1) = "(" And mid$(Cells(14 + sftfit2, i), Len(Cells(14 + sftfit2, i)), 1) = ";" Then
                If IsNumeric(mid$(Cells(14 + sftfit2, i), 2, Len(Cells(14 + sftfit2, i)) - 2)) = True Then
                    ReDim ratio(1)
                    ratio(1) = mid$(Cells(14 + sftfit2, i), 2, Len(Cells(14 + sftfit2, i)) - 2)
                    a3 = ratio(1)
                    iRow = iRow + 1
                Else
                    TimeCheck = MsgBox(strErr, vbCritical)
                    Call GetOutFit
                    Exit Sub
                End If
            ElseIf iRow > 1 And mid$(Cells(14 + sftfit2, i), 1, 1) = "(" Then
                TimeCheck = MsgBox(strErr, vbCritical)
                Call GetOutFit
                Exit Sub
            ElseIf iRow > 1 And mid$(Cells(14 + sftfit2, i), Len(Cells(14 + sftfit2, i)), 1) = ";" Then
                If IsNumeric(mid$(Cells(14 + sftfit2, i), 1, Len(Cells(14 + sftfit2, i)) - 1)) = True Then
                    ReDim Preserve ratio(iRow)
                    ratio(iRow) = mid$(Cells(14 + sftfit2, i), 1, InStr(1, Cells(14 + sftfit2, i), ";") - 1)
                    iRow = iRow + 1
                Else
                    TimeCheck = MsgBox(strErr, vbCritical)
                    Call GetOutFit
                    Exit Sub
                End If
            ElseIf iRow > 1 And mid$(Cells(14 + sftfit2, i), Len(Cells(14 + sftfit2, i)), 1) = ")" Then
                If IsNumeric(mid$(Cells(14 + sftfit2, i), 1, Len(Cells(14 + sftfit2, i)) - 1)) = True Then
                    ReDim Preserve ratio(iRow)
                    ratio(iRow) = mid$(Cells(14 + sftfit2, i), 1, Len(Cells(14 + sftfit2, i)) - 1)
                Else
                    TimeCheck = MsgBox(strErr, vbCritical)
                    Call GetOutFit
                    Exit Sub
                End If
                For iCol = iRow - 1 To 0 Step -1        ' max amplitude ratio to be reference, not in the first bracket!
                    If IsNumeric(ratio(iRow - iCol)) = True Then
                        If ratio(iRow - iCol) >= a3 Then
                            a3 = ratio(iRow - iCol)
                            k = iRow - iCol
                        Else
                            k = 1           ' Added in ver. 7.19
                        End If
                    End If
                Next
                For iCol = iRow - 1 To 0 Step -1
                    If IsNumeric(ratio(iRow - iCol)) = True And ratio(iRow - iCol) > 0 Then
                        If iRow - iCol = k Then
                           SolverAdd CellRef:=Cells(6, i - iRow + k), Relation:=1, FormulaText:=Cells(3, 101).Value - Cells(2, 101).Value
                           Cells(15 + sftfit2, i - iCol + 110).Value = ratio(iRow - iCol) / ratio(k)
                        Else
                           SolverAdd CellRef:=Cells(6, i - iCol), Relation:=2, FormulaText:=Cells(6, i - iRow + k) * ratio(iRow - iCol) / ratio(k)
                           Cells(15 + sftfit2, i - iCol + 110).Value = ratio(iRow - iCol) / ratio(k)
                        End If
                    ElseIf ratio(iRow - iCol) = "NaN" Then
                        SolverAdd CellRef:=Cells(6, i - iCol), Relation:=1, FormulaText:=Cells(3, 101).Value - Cells(2, 101).Value
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
            SolverAdd CellRef:=Cells(6, i), Relation:=1, FormulaText:=Cells(3, 101).Value - Cells(2, 101).Value
        Else
            SolverAdd CellRef:=Cells(6, i), Relation:=1, FormulaText:=Cells(3, 101).Value - Cells(2, 101).Value
        End If
    Next
    
    strErr = "BE diff format error: [ i; nj; k] and i,j,k > 0" & vbCrLf & "n represents negative sign."
    iRow = 0
    For i = 5 To (4 + j)
        If Not Cells(15 + sftfit2, i) = vbNullString Then
            If iRow = 0 And StrComp(Cells(15 + sftfit2, i), "[", 1) = 0 Then
                ReDim bediff(1)
                iRow = iRow + 1
            ElseIf iRow > 0 And StrComp(Cells(15 + sftfit2, i), "[", 1) = 0 Then
                TimeCheck = MsgBox(strErr, vbCritical)
                Call GetOutFit
                Exit Sub
            ElseIf iRow > 0 And mid$(Cells(15 + sftfit2, i), Len(Cells(15 + sftfit2, i)), 1) = ";" Then
                If IsNumeric(mid$(Cells(15 + sftfit2, i), 1, Len(Cells(15 + sftfit2, i)) - 1)) = True Then
                    ReDim Preserve bediff(iRow)
                    bediff(iRow) = mid$(Cells(15 + sftfit2, i), 1, Len(Cells(15 + sftfit2, i)) - 1)
                    iRow = iRow + 1
                ElseIf mid$(Cells(15 + sftfit2, i), 1, 1) = "n" Then
                    If IsNumeric(mid$(Cells(15 + sftfit2, i), 2, Len(Cells(15 + sftfit2, i)) - 2)) = True Then
                        ReDim Preserve bediff(iRow)
                        bediff(iRow) = -1 * mid$(Cells(15 + sftfit2, i), 2, Len(Cells(15 + sftfit2, i)) - 2)
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
            ElseIf iRow > 0 And mid$(Cells(15 + sftfit2, i), Len(Cells(15 + sftfit2, i)), 1) = "]" Then
                If IsNumeric(mid$(Cells(15 + sftfit2, i), 1, Len(Cells(15 + sftfit2, i)) - 1)) = True Then
                    ReDim Preserve bediff(iRow)
                    bediff(iRow) = mid$(Cells(15 + sftfit2, i), 1, Len(Cells(15 + sftfit2, i)) - 1)
                    iRow = iRow
                ElseIf mid$(Cells(15 + sftfit2, i), 1, 1) = "n" Then
                    If IsNumeric(mid$(Cells(15 + sftfit2, i), 2, Len(Cells(15 + sftfit2, i)) - 2)) = True Then
                        ReDim Preserve bediff(iRow)
                        bediff(iRow) = -1 * mid$(Cells(15 + sftfit2, i), 2, Len(Cells(15 + sftfit2, i)) - 2)
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
                        SolverAdd CellRef:=Cells(2, i - iCol), Relation:=2, FormulaText:=Cells(2, i - iRow) + bediff(iRow - iCol)
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
    For i = 5 To (4 + j)
        If Not Cells(14 + sftfit2, i) = vbNullString Then
            If iRow = 1 And mid$(Cells(14 + sftfit2, i), 1, 1) = "(" Then
                iRow = iRow + 1
            ElseIf iRow > 1 And mid$(Cells(14 + sftfit2, i), Len(Cells(14 + sftfit2, i)), 1) = ";" Then
                iRow = iRow + 1
            ElseIf iRow > 1 And mid$(Cells(14 + sftfit2, i), Len(Cells(14 + sftfit2, i)), 1) = ")" Then
                For iCol = iRow - 1 To 0 Step -1
                    If IsNumeric(Cells(15 + sftfit2, i - iCol + 110)) = True And Cells(6, i - iCol) > 0 And Cells(6, i - iRow + 1) > 0 Then
                        Cells(16 + sftfit2, i - iCol + 110).Value = Cells(6, i - iCol) / Cells(6, i - iRow + 1)
                        If Cells(15 + sftfit2, i - iCol + 110) > 0 Then
                            If Abs((Cells(16 + sftfit2, i - iCol + 110) - Cells(15 + sftfit2, i - iCol + 110)) / Cells(15 + sftfit2, i - iCol + 110)) > a2 And fileNum < Cells(17, 101).Value Then
                                GoTo Resolve
                            ElseIf fileNum >= Cells(17, 101).Value And Abs((Cells(16 + sftfit2, i - iCol + 110) - Cells(15 + sftfit2, i - iCol + 110)) / Cells(15 + sftfit2, i - iCol + 110)) > a2 Then
                                a0 = a0 + Abs((Cells(16 + sftfit2, i - iCol + 110) - Cells(15 + sftfit2, i - iCol + 110)) / Cells(15 + sftfit2, i - iCol + 110))
                                a1 = a2
                                GoTo ExitIter
                            Else
                                a0 = a0 + Abs((Cells(16 + sftfit2, i - iCol + 110) - Cells(15 + sftfit2, i - iCol + 110)) / Cells(15 + sftfit2, i - iCol + 110))
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
    For i = 5 To (4 + j)
        If Not Cells(15 + sftfit2, i) = vbNullString Then
            If iRow = 0 And StrComp(Cells(15 + sftfit2, i), "[", 1) = 0 Then
                iRow = iRow + 1
            ElseIf iRow > 0 And mid$(Cells(15 + sftfit2, i), Len(Cells(15 + sftfit2, i)), 1) = ";" Then
                iRow = iRow + 1
            ElseIf iRow > 0 And mid$(Cells(15 + sftfit2, i), Len(Cells(15 + sftfit2, i)), 1) = "]" Then
                For iCol = iRow - 1 To 0 Step -1
                    If IsNumeric(bediff(iRow - iCol)) = True Then
                        If Abs((bediff(iRow - iCol) - (Cells(2, i - iCol) - Cells(2, i - iRow))) / bediff(iRow - iCol)) > a2 And fileNum < Cells(17, 101).Value Then
                            Debug.Print Cells(2, i - iCol), Cells(2, i - iRow), Abs(Cells(2, i - iCol) - Cells(2, i - iRow))
                            GoTo Resolve
                        ElseIf fileNum >= Cells(17, 101).Value Then
                            a1 = a1 + Abs((bediff(iRow - iCol) - (Cells(2, i - iCol) - Cells(2, i - iRow))) / bediff(iRow - iCol))
                            GoTo ExitIter
                        Else
                            a1 = a1 + Abs((bediff(iRow - iCol) - (Cells(2, i - iCol) - Cells(2, i - iRow))) / bediff(iRow - iCol))
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
    'Call ShowResults        ' Show Solver results
    'Cells(16, 4) = fileNum 'iteration numbers
    Range(Cells(15 + sftfit2, 4 + 110), Cells(16 + sftfit2, 4 + j + 110)).ClearContents
    
    If Cells(7, (4 + i)).Value < 0.999 And Cells(7, (4 + i)).Value > 0.001 Then
        For i = 1 To j
            If Cells(7, (4 + i)).Font.Italic = "True" Then
                For k = 1 To numData
                    If Cells((startR - 1 + k), 1).Value < Cells(2, (4 + i)).Value Then
                        Cells((startR - 1 + k), (4 + i)).FormulaR1C1 = "=R6C * ((R7C) *((((R5C)/2)^2)/((RC[" & (-3 - i) & "]-R2C)^2 + ((R5C)/2)^2)) + (1- R7C) *(EXP(-(1/2)*((RC[" & (-3 - i) & "]-R2C)/(R5C/2.35))^2)))"
                    Else
                        Cells((startR - 1 + k), (4 + i)).FormulaR1C1 = "=R6C * ((R7C) *((((R4C)/2)^2)/((RC[" & (-3 - i) & "]-R2C)^2 + ((R4C)/2)^2)) + (1- R7C) *(EXP(-(1/2)*((RC[" & (-3 - i) & "]-R2C)/(R4C/2.35))^2)))"
                    End If
                Next
                p = p + 1
            End If
        Next
        If p > 0 Then
            imax = imax + 1
            'Debug.Print Abs(Cells(9, 2).Value - A)
            If Abs(Cells(9 + sftfit2, 2).Value - A) > 1 Then GoTo AsymIteration     ' Tolerance = 1
        End If
    End If
    
    For i = 1 To j
        If Cells(7, (4 + i)).Value > 0 And Cells(7, (4 + i)).Value < 1 Then
            If Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleSingle Then   ' exponential asymmetric blend based Voigt
                Cells(13 + sftfit2, (4 + i)).Value = vbNullString ' R5C to be exp decay parameter.
            ElseIf Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleDouble Then   ' Ulrik Gelius mode asymmetric blend based Voigt (GL sum)
                Cells(13 + sftfit2, (4 + i)).Value = vbNullString
            ElseIf Cells(11, (4 + i)).Value = "GL" Then
                Cells(13 + sftfit2, (4 + i)).Value = vbNullString
            Else
                Cells(13 + sftfit2, (4 + i)).Value = Cells(5, (4 + i)).Value / Cells(4, (4 + i)).Value
            End If
        ElseIf Cells(7, (4 + i)).Value = 0 Then
            Cells(7, (4 + i)).Value = "Gauss"
            Cells(5, (4 + i)).Value = vbNullString
            Cells(13 + sftfit2, (4 + i)).Value = vbNullString
        ElseIf Cells(7, (4 + i)).Value = 1 Then
            Cells(7, (4 + i)).Value = "Lorentz"
            Cells(13 + sftfit2, (4 + i)).Value = vbNullString
            If Cells(7, (4 + i)).Font.Italic = "True" Then  'Doniach-Sunjic (DS)
                ' FWHM2 to be alpha; asymmetric parameter
            Else
                Cells(5, (4 + i)).Value = vbNullString
            End If
        End If
        
        Cells(3, (4 + i)).FormulaR1C1 = "=(R12C101 - R13C101 - R14C101 - R2C)" ' KE
    Next
    
    Cells(8, 101).Value = Cells(8, 101).Value + 1     ' means already fit once
    
    If startR > 21 + sftfit Then
        Range(Cells(23 + sftfit + numData, 5), Cells(2 + numData + startR - 1, 55)).ClearContents
    End If
    
    If endR < numData + 20 + sftfit Then
        Range(Cells(2 + numData + endR + 1, 5), Cells(2 * numData + 22 + sftfit, 55)).ClearContents
    End If
    
    Call descriptInitialFit
    
    If StrComp(str1, "Pe", 1) = 0 Then
        Cells(2, 4).Value = "PE"
        Range(Cells(3, 5), Cells(3, 55)).ClearContents
        Range(Cells(12 + sftfit2, 5), Cells(12 + sftfit2, 55)).ClearContents
        Range(Cells(18 + sftfit2, 5), Cells(18 + sftfit2, 55)).ClearContents
    ElseIf StrComp(str1, "Po", 1) = 0 Then
        Cells(2, 4).Value = "Po"
        Range(Cells(3, 5), Cells(3, 55)).ClearContents
        Range(Cells(12 + sftfit2, 5), Cells(12 + sftfit2, 55)).ClearContents
        Range(Cells(18 + sftfit2, 5), Cells(18 + sftfit2, 55)).ClearContents
    End If
    
    Call GetOutFit
End Sub

Sub FitEF()
    If startR > 21 + sftfit Then
        If IsEmpty(Cells(startR - 1, 3)) = False Then
            Range(Cells(21 + sftfit, 3), Cells(startR - 1, 4)).ClearContents
            Cells(8, 101).Value = 0
        ElseIf IsEmpty(Cells(startR, 3)) = True Then
            Cells(8, 101).Value = 0
        End If
    End If
    
    If endR < numData + 20 + sftfit Then
        If IsEmpty(Cells(endR + 1, 3)) = False Then
            Range(Cells(endR + 1, 3), Cells(numData + 20 + sftfit, 4)).ClearContents
            Cells(8, 101).Value = 0
        ElseIf IsEmpty(Cells(endR, 3)) = True Then
            Cells(8, 101).Value = 0
        End If
    End If
    
    If Cells(8, 101).Value > 0 Then GoTo SkipInitialEF2
    Range(Cells(1, 3), Cells(15 + sftfit2, 55)).ClearContents
    Range(Cells(20 + sftfit, 3), Cells((2 * numData + 22 + sftfit), 55)).ClearContents
    Range(Cells(1, 3), Cells(15 + sftfit2, 55)).Interior.ColorIndex = xlNone
    
    Call descriptEFfit1
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

    For k = 2 To 8
        If Cells(k, 5).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 5), Relation:=2, FormulaText:=Cells(k, 5)
        End If
    Next

    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
    
SkipInitialEF2:
    ' start second solver
    
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
    ' end second solver
    
    Cells(8, 101).Value = Cells(8, 101).Value + 1     ' means already fit once
    
    Call descriptEFfit2
    
    If Cells(8, 101).Value > 1 Then Exit Sub
    ActiveSheet.ChartObjects(1).Activate
    ActiveChart.SeriesCollection.NewSeries  '7.45
    With ActiveChart.SeriesCollection(2)
        .ChartType = xlXYScatterLinesNoMarkers
        .XValues = rng
        .Values = rng.Offset(, 2)
        .Border.ColorIndex = 33
        .Format.Line.Weight = 3
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
        strTest = "e"
    End If
    
    If StrComp(strTest, "p", 1) = 0 Then
        Range(Cells(6, 1), Cells(7 + sftfit2 - 2, 2)).ClearContents
        Range(Cells(6, 1), Cells(7 + sftfit2 - 2, 2)).Interior.ColorIndex = xlNone
        Cells(5, 1).Value = "a3"
    ElseIf StrComp(strTest, "a", 1) = 0 Then
        Range(Cells(8, 1), Cells(7 + sftfit2 - 2, 2)).ClearContents
        Range(Cells(8, 1), Cells(7 + sftfit2 - 2, 2)).Interior.ColorIndex = xlNone
        Cells(6, 1).Value = "Slope"
        Cells(7, 1).Value = "ratio L:A"
    ElseIf StrComp(strTest, "t", 1) = 0 Then
        Cells(5, 1).Value = "Norm"
        Cells(5, 2).Font.Bold = "False"
        Range(Cells(7, 1), Cells(7 + sftfit2 - 2, 2)).ClearContents
        Range(Cells(7, 1), Cells(7 + sftfit2 - 2, 2)).Interior.ColorIndex = xlNone
    ElseIf StrComp(strTest, "v", 1) = 0 Then
        If Cells(9, 2).Value = vbNullString Then
            Cells(9, 1).Value = "No edge"
        ElseIf Cells(9, 2).Value < Cells(12 + sftfit2, 2).Value And Cells(9, 2).Value > Cells(11 + sftfit2, 2).Value Then
            Cells(9, 1).Value = "Pre-edge"
        Else
            Cells(9, 1).Value = "Both ends"
        End If
        Range(Cells(7, 1), Cells(8, 2)).ClearContents
        Range(Cells(7, 1), Cells(8, 2)).Interior.ColorIndex = xlNone
        Range(Cells(10, 1), Cells(7 + sftfit2 - 2, 2)).ClearContents
        Range(Cells(10, 1), Cells(7 + sftfit2 - 2, 2)).Interior.ColorIndex = xlNone
    ElseIf StrComp(strTest, "e", 1) = 0 Then
        Cells(8, 1).Value = "Norm"
        Cells(6, 1).Value = "Poly2nd"
        Cells(7, 1).Value = "Poly3rd"
        Cells(5, 1).Value = "Slope BG"
        Range(Cells(9, 4), Cells(19 + sftfit2, 5)).ClearContents
        Range(Cells(9, 4), Cells(19 + sftfit2, 5)).Interior.ColorIndex = xlNone
        Range(Cells(9, 1), Cells(7 + sftfit2 - 2, 5)).ClearContents
        Range(Cells(9, 1), Cells(7 + sftfit2 - 2, 5)).Interior.ColorIndex = xlNone
    ElseIf StrComp(strTest, "s", 1) = 0 Then
        ' Shirley
        Cells(5, 2).Value = fileNum
        Cells(5, 1).Value = "Iteration"
        Cells(5, 2).Font.Bold = "False"
        Range(Cells(6, 1), Cells(7 + sftfit2 - 2, 2)).ClearContents
        Range(Cells(6, 1), Cells(7 + sftfit2 - 2, 2)).Interior.ColorIndex = xlNone
    Else      ' Solver mode; ShirleyBG2 for iteration mode.
        Cells(5, 2).Value = fileNum
        Cells(5, 1).Value = "Iteration"
        Cells(5, 2).Font.Bold = "False"
    End If
    
    For i = 1 To j
        If mid(Cells(11, (4 + i)).Value, 1, 1) = "E" Then
            Range(Cells(9, (4 + i)), Cells(10, (4 + i))) = vbNullString
            Cells(5, (4 + i)) = vbNullString
            'Debug.Print "E"
        ElseIf mid(Cells(11, (4 + i)).Value, 1, 1) = "T" Then
            Cells(10, (4 + i)) = vbNullString
            Cells(5, (4 + i)) = vbNullString
            'Debug.Print "E"
        ElseIf Cells(7, (4 + i)).Value = 0 Or Cells(7, (4 + i)).Value = "Gauss" Or Cells(11, (4 + i)).Value = "GL" Then ' G
            Range(Cells(8, (4 + i)), Cells(10, (4 + i))) = vbNullString
            Cells(5, (4 + i)) = vbNullString
        ElseIf Cells(7, (4 + i)).Value = 1 Or Cells(7, (4 + i)).Value = "Lorentz" Then
            If Cells(7, (4 + i)).Font.Italic = "True" Then  'Doniach-Sunjic (DS)
               Cells(5, (4 + i)) = vbNullString
               Cells(10, (4 + i)) = vbNullString
            Else    ' L
                Range(Cells(8, (4 + i)), Cells(10, (4 + i))) = vbNullString
                Cells(5, (4 + i)) = vbNullString
            End If
        Else
            If Cells(1, 1).Value = "EF" Then
                Cells(10, (4 + i)) = vbNullString
            Else
                Range(Cells(8, (4 + i)), Cells(10, (4 + i))) = vbNullString
            End If
        End If
    Next
    
    Cells(7 + sftfit2, 1).Value = "Peak fit"
    Cells(7 + sftfit2, 2).Value = vbNullString
    Range(Cells(2, 3), Cells(7 + sftfit2, 3)).ClearContents
    Range(Cells(2, 3), Cells(7 + sftfit2, 3)).Interior.ColorIndex = xlNone
    Cells(1, 1).Select
    
    'Call ShowResults        ' Show Solver results
    Application.Calculation = xlCalculationAutomatic
    Call GetOut
End Sub

Function ExistSheet(sheetName) As Boolean
    Dim r, cnt As Integer
    
    cnt = Sheets.Count
    ExistSheet = False
    For r = 1 To cnt
        If Sheets(r).Name = sheetName Then
            ExistSheet = True
            Exit For
        End If
    Next
End Function

Sub ThermoAvgBL()
    If mid$(Cells(2, 3).Value, 1, 3) = "Ele" Then
        strList = ActiveSheet.Name + ":"
        imax = 1
        startR = 5
        Cells(2, 3).Value = " Elements="
    Else
        NumSheets = Sheets.Count
        strList = ""
        imax = 0
        startR = 0
    End If
    
    For ns = NumSheets To 1 Step -1
        Sheets(ns).Activate
        If mid$(Cells(1, 8).Value, 1, 3) = "Acq" Then
            Cells(1, 8).Value = " Acquisition Parameters :"
            strList = strList + ActiveSheet.Name + ":"
            imax = imax + 1
        End If
    Next
    
    If imax = 0 Then Exit Sub
    sh = ActiveWorkbook.Name
    wbp = ActiveWorkbook.Path
    Application.DisplayAlerts = False
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=wbp + "\sum.xlsx", FileFormat:=51
    wbc = ActiveWorkbook.Name
    Application.DisplayAlerts = True
    Workbooks(sh).Activate
SkipAvg:
    Gnum = 1
    NumSheets = imax
    
    For ns = NumSheets To 1 Step -1
        imax = InStr(Gnum, strList, ":", 1)
        strSheetAvgName = mid$(strList, Gnum, InStr(imax, strList, ":", 1) - Gnum)
        Gnum = imax + 1
        Set sheetAvg = Worksheets(strSheetAvgName)
        sheetAvg.Activate
        Set rng = [A:A]
        numDataN = Application.CountA(rng) - 4
        iRow = Cells(8 - startR, 4).Value           ' # of spectra
        strscanNum = Cells(17 - startR, 1).Value    ' BE or KE
        Elem = mid$(Cells(7, 9).Value, 1, 3)        ' CRR or CAE
        k = 18 - startR
        imax = 1
        
        If iRow = 0 And IsEmpty(Cells(9 - startR, 1).Value) = True Then
            iRow = 1
            numDataN = numDataN + 1
            strscanNum = Cells(16 - startR, 1).Value
            Elem = mid$(Cells(7, 9).Value, 1, 3)
            k = 17 - startR
        ElseIf iRow = 0 And IsEmpty(Cells(9 - startR, 1).Value) = False Then       ' new Avantage 4.87
            iRow = 1
            strscanNum = Cells(15 - startR, 1).Value
            Elem = mid$(Cells(7, 9).Value, 1, 3)
            k = 16 - startR
        ElseIf iRow > 0 And IsEmpty(Cells(10 - startR, 1).Value) = False Then       ' new Avantage 4.87
            numDataN = Application.CountA(rng) - 5
            strscanNum = Cells(16 - startR, 1).Value
            k = 17 - startR
        End If
    
        For iCol = iRow To 1 Step -1
            sheetAvg.Activate
            Set dataIntGraph = Range(Cells((k + 1), iCol + 2), Cells((k + 1), iCol + 2).End(xlDown))
            j = numDataN
            numData = numDataN
            q = 0
            C = dataIntGraph
            For p = 1 To j
                If IsNumeric(C(p, 1)) = False Then
                    numData = numData - 1
                    q = 0
                ElseIf IsNumeric(C(p, 1)) = True And q = 0 Then
                    iniRow = k + p
                    q = 1
                End If
            Next
    
            b = Range(Cells(iniRow, 1), Cells(iniRow + numData - 1, 1))
            D = Range(Cells(iniRow, iCol + 2), Cells(iniRow + numData - 1, iCol + 2))
    
            strTest = Cells(k - 2, iCol + 2).Value        ' new Avantage 4.87
            
            If Len(strTest) = 0 Then
                strTest = "_" + strSheetAvgName
            End If
            strSheetDataName = mid$(strTest, 1, 25)
    
            If ExistSheet(strSheetDataName) Then
                Application.DisplayAlerts = False
                Worksheets(strSheetDataName).Delete
                Application.DisplayAlerts = True
            End If
    
            Worksheets.Add().Name = strSheetDataName
            Set sheetData = Worksheets(strSheetDataName)
            sheetData.Activate
    
            Range(Cells(2, 1), Cells(numData + 1, 1)) = b
            Range(Cells(2, 2), Cells(numData + 1, 2)) = D
            
            If mid$(strscanNum, 1, 3) = "Bin" Then
                Cells(1, 1).Value = "BE/eV"
            ElseIf Elem = "CRR" Then
                Cells(1, 1).Value = "AE/eV"
            Else
                Cells(1, 1).Value = "KE/eV"
            End If
            Cells(1, 2).Value = strTest
            
            If StrComp(mid$(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") + 1, 4), "xlsx", 1) = 0 Then
                wb = strTest + ".xlsx"
            Else
                wb = strTest + ".xls"
            End If
                
            A = Range(Cells(1, 1), Cells(numData + 1, 2))
            Workbooks(wbc).Activate
            Range(Cells(1, ns * 2 - 1), Cells(numData + 1, ns * 2)) = A
            Cells(1, ns * 2 - 1).Value = "KE" + mid$(Cells(1, ns * 2).Value, 7, 7)
            Cells(1, ns * 2).Value = mid$(Cells(1, ns * 2).Value, 2, 12)
            Workbooks(sh).Activate
            sheetData.Activate
            
            If iCol > 1 Or ns > 1 Then
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveCopyAs Filename:=ActiveWorkbook.Path + "\" + wb
                Application.DisplayAlerts = True
                If ExistSheet(strSheetDataName) Then
                    Application.DisplayAlerts = False
                    Worksheets(strSheetDataName).Delete
                    Application.DisplayAlerts = True
                End If
            Else
                wb = strTest + ".xlsx"
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path + "\" + wb, FileFormat:=51
                Application.DisplayAlerts = True
                strSheetGraphName = "Graph_" + mid$(strTest, 1, 25)
                strSheetCheckName = "Check_" + mid$(strTest, 1, 25)
                strSheetFitName = "Fit_" + mid$(strTest, 1, 25)
            End If
        Next
    Next
    
    Application.DisplayAlerts = False
    Workbooks(wbc).Save
    Workbooks(wbc).Close
    Application.DisplayAlerts = True
End Sub

Sub ScanCheck()
    j = 3   ' 0 for CPS, 1 for Ie, 2 for Ip, 3 for CPS/Ip
    If ExistSheet(strSheetCheckName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetCheckName).Delete
        Application.DisplayAlerts = True
    End If
    
    Worksheets.Add().Name = strSheetCheckName
    Set sheetCheck = Worksheets(strSheetCheckName)
    iniRow = 22
    endRow = iniRow + numData - 1
    sheetData.Activate
    C = Range(Cells(iniRow, 1), Cells(endRow, 1))
    sheetCheck.Activate
    
    For k = 0 To j
        Range(Cells(2 + (k * (10 + numData)), 1), Cells((k * (10 + numData)) + numData + 1, 1)) = C
        Cells((k * (10 + numData)) + numData + 2, 1).Value = "Avg"
    Next
    
    sheetData.Activate
    A = Range(Cells(2, 10), Cells(numData + 2 + (j * (10 + numData)), (scanNumR + 10)))
    
    For i = 1 To scanNumR
        iniRow = 22 + ((i - 1) * (3 + numData))
        endRow = iniRow + numData - 1
        C = Range(Cells(iniRow, 2), Cells(endRow, (2 + j)))
        
        For k = 1 To numData
            For q = 0 To j
                If q = 2 Then
                    chkMax = 1000000000000#     ' Ip to be pico-amps
                ElseIf q = 3 Then
                    chkMax = 0.000000000001     ' CPS/Ip to be normalized by pico-amps
                Else
                    chkMax = 1
                End If
                
                If IsNumeric(C(k, 1 + q)) = False Then Exit For
                
                A(k + (q * (10 + numData)), i) = C(k, 1 + q) * chkMax
                A(numData + 1 + (q * (10 + numData)), i) = A(k + (q * (10 + numData)), i) + A(numData + 1 + (q * (10 + numData)), i)
                A(k + (q * (10 + numData)), scanNumR + 1) = A(k + (q * (10 + numData)), i) + A(k + (q * (10 + numData)), scanNumR + 1)
            Next
        Next
    Next
    
    For i = 1 To scanNumR
        For q = 0 To j
            A(numData + 1 + (q * (10 + numData)), i) = A(numData + 1 + (q * (10 + numData)), i) / numData
        Next
    Next
    
    For q = 0 To j
        For k = 1 To numData
            A(k + (q * (10 + numData)), scanNumR + 1) = A(k + (q * (10 + numData)), scanNumR + 1) / scanNumR
        Next
    Next
    
    sheetCheck.Activate
    Range(Cells(2, 2), Cells(numData + 2 + (j * (10 + numData)), (scanNumR + 2))) = A
    A = Range(Cells(2, 1), Cells(numData + 2 + (j * (10 + numData)), (scanNumR + 1)))
    
    For k = 0 To j
        Set dataCheck = Range(Cells(2 + (k * (10 + numData)), 1), Cells(numData + (k * (10 + numData)) + 1, 1))
        Cells(1 + k * (numData + 10), scanNumR + 2).Value = "Avg"
        
        With ActiveSheet.ChartObjects.Add(20, 200, 1100, 500).Chart
            .ChartType = xlXYScatterLinesNoMarkers
            For i = 1 To scanNumR
                With .SeriesCollection.NewSeries
                    .XValues = dataCheck
                    .Values = dataCheck.Offset(, i)
                    .Format.Line.Weight = 1
                End With
            Next
            .ChartType = xlXYScatterLinesNoMarkers
        End With
        
        With ActiveSheet.ChartObjects(1 + k)
            .Top = 20 + (k * (500 / windowSize))
            With .Chart.Axes(xlValue)
                 .HasTitle = True
                If k = 1 Then
                    .AxisTitle.Text = "Ie (mA)"
                ElseIf k = 2 Then
                    .AxisTitle.Text = "Ip (pA)"
                ElseIf k = 3 Then
                    .AxisTitle.Text = "CPS/Ip"
                Else
                    .AxisTitle.Text = "CPS"
                End If
            End With

        End With
    Next
    
    For Each myChartOBJ In ActiveSheet.ChartObjects
        With myChartOBJ
            .Left = 200
            .Width = (550 * windowRatio) / windowSize
            .Height = 500 / windowSize
            '.Chart.Legend.Delete
        End With
        With myChartOBJ.Chart.Axes(xlCategory, xlPrimary)
            .MinorTickMark = xlOutside
            .MinimumScale = startEk
            .MaximumScale = endEk
            .HasTitle = True
            .AxisTitle.Text = "Kinetic energy (eV)"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .HasMajorGridlines = True
            .MajorUnit = numMajorUnit   '2
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With myChartOBJ.Chart.Axes(xlValue)
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
    
    For i = 1 To scanNumR
        For q = 0 To j
            Cells(1 + (q * (10 + numData)), (1 + i)).Interior.Color = ActiveSheet.ChartObjects(1).Chart.SeriesCollection(i).Border.Color
            Cells(1 + (q * (10 + numData)), (1 + i)).Font.ColorIndex = 2
            Cells(1 + (q * (10 + numData)), (1 + i)).Value = "Scan#" + CStr(i)
            ActiveSheet.ChartObjects(1 + q).Chart.SeriesCollection(i).Name = "='" & ActiveSheet.Name & "'!R" & (1 + (q * (10 + numData))) & "C" & (1 + i) & ""
        Next
    Next

    Cells(1, 1).Value = "KE : CPS"
    Cells(1 + (1 * (10 + numData)), 1).Value = "Ie"
    Cells(1 + (2 * (10 + numData)), 1).Value = "Ip"
    Cells(1 + (3 * (10 + numData)), 1).Value = "CPS/Ip"
    Cells(1, 1).Select
SkipSelection3:
    If numscancheck > 0 Then
    
        U = Range(Cells(2, 1), Cells(numData + 2 + (j * (10 + numData)), (scanNumR + 2)))
        For i = 1 To numData
            For q = 0 To j
                U(i + ((numData + 10) * q), scanNumR + 2) = 0
            Next
        Next
        
        iCol = 0
        For k = 1 To numscancheck
            If ratio(k) <= scanNumR Then
                For i = 1 To numData
                    For q = 0 To j
                        U(i + ((numData + 10) * q), scanNumR + 2) = U(i + ((numData + 10) * q), 1 + ratio(k)) + U(i + ((numData + 10) * q), scanNumR + 2)
                    Next
                Next
            Else
                iCol = iCol + 1
            End If
        Next
        
        If numscancheck - iCol <= 0 Then
            numscancheck = 0
            GoTo SkipSelection3
        End If
        
        For i = 1 To numData
            For q = 0 To j
                U(i + ((numData + 10) * q), scanNumR + 2) = U(i + ((numData + 10) * q), scanNumR + 2) / (numscancheck - iCol)
            Next
        Next
        
        scanNum = (numscancheck - iCol)
        Range(Cells(2, 1), Cells(numData + 2 + (j * (10 + numData)), scanNumR + 2)) = U
        Set dataData = Union(Range(Cells(2 + (j * (10 + numData)), 1), Cells(numData + 1 + (j * (10 + numData)), 1)), Range(Cells(2 + (j * (10 + numData)), scanNumR + 2), Cells(numData + 1 + (j * (10 + numData)), scanNumR + 2)))
        Set dataKeData = Range(Cells(2 + (j * (10 + numData)), 1), Cells(numData + 1 + (j * (10 + numData)), 1))
        Set dataIntData = dataKeData.Offset(, scanNumR + 1)
        
        If scanNumR > scanNum Then
            TimeC1 = Timer
            MsgBox "# of scanned data: " & scanNum & "/" & scanNumR & " in " & strTest & " will be analyzed."
            TimeC2 = Timer
        End If
    End If
    
    sheetData.Activate
End Sub

Sub ObbCheck()
    iRow = 5          ' # of graph: 1 to CPS, 2 to Ie, 3 to CPS/Ip, 4 to Flux, 5 to fluence, 6 to Ip/100mA.
    sheetCheck.Activate
    A = Range(Cells(2, 1), Cells(numData + 1 + (j * (10 + numData)), (scanNumR + 1)))
    strSheetCheckName2 = "Time_" + strSheetDataName        ' check series scans
    
    If ExistSheet(strSheetCheckName2) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetCheckName2).Delete
        Application.DisplayAlerts = True
    End If
    
    Worksheets.Add().Name = strSheetCheckName2
    Set sheetCheck2 = Worksheets(strSheetCheckName2)
    Cells(1, 1).Value = "KE"
    Cells(1, 2).Value = "#scan"
    Cells(1, 3).Value = "Time"
    Cells(1, 4).Value = "Elapse"
    Cells(1, 5).Value = "CPS"
    Cells(1, 6).Value = "Ie"       ' unit in mA
    Cells(1, 7).Value = "Ip"       ' unit in pA
    Cells(1, 8).Value = "CPS/Ip"          ' normalized by Ip in pA
    Cells(1, 9).Value = "# ph./s"       ' unit in Giga counts
    Cells(1, 10).Value = "Fluence"      ' unit in Giga counts
    D = Range(Cells(2, 1), Cells(totalDataPoints + 1, 11))
    sheetData.Activate
    checkTime = TimeValue(mid$(Cells(4, 1), 10, 12))                                            ' Time
    a1 = mid$(Cells(14, 1), 11, 3)                                                              ' Sampling
    checkDate = DateValue(mid$(Cells(3, 1), InStr(1, Cells(3, 1), ",") + 1, Len(Cells(3, 1))))  ' Date
    sheetCheck2.Activate
    checkDate = checkDate & " " & checkTime   ' Date & Time
    
    If numscancheck > 0 Then
        chkMax = ratio(1)
    Else
        chkMax = 1
    End If
    
    For q = 0 To j
        imax = 1
        k = 1
        chkMin = 0
        
        For i = 1 To totalDataPoints
            D(i, q + 5) = A(k + (q * (10 + numData)), imax + 1)
            D(i, 1) = A(k, 1)
            D(i, 2) = imax + chkMax - 1
            D(i, 4) = a1 * 1.0199 * (i + (imax - 1))         '  1.0199 is the estimated time for one-point sampling.
            k = k + 1
            If i = imax * numData Then
                imax = imax + 1                              ' one additional sampling in between previous and present scans.
                k = 1
                'chkMin = 1.689 * (imax - 1)
            End If
        Next
    Next
    
    For i = 1 To totalDataPoints
        D(i, 3) = DateAdd("s", CInt(D(i, 4)), checkDate)
        D(i, 9) = D(i, 7) * (1E-21) / ((1.6E-19) * qe * (1 - trans)) ' # of ph./s (Giga counts) is evaluated from QE based on Ip and photodiode measurements.
        'D(i, 11) = D(i, 7) * 100 / D(i, 6)                           ' This is Ip (pA) normalized at 100 mA of Ie.
        If i = 1 Then
            D(i, 10) = D(i, 9)
        ElseIf i > 1 Then
            D(i, 10) = D(i - 1, 10) + D(i, 9) * (D(i, 4) - D(i - 1, 4)) ' This fluence (Giga counts) is integrated over the elapsed time.
        End If
    Next
    
    Range(Cells(2, 1), Cells(totalDataPoints + 1, 11)) = D
    Set dataCheck = Range(Cells(2, 3), Cells(totalDataPoints + 1, 3))
    Set dataIntCheck = Union(Range(Cells(2, 3), Cells(totalDataPoints + 1, 3)), Range(Cells(2, 5), Cells(totalDataPoints + 1, 8)))
    TimeC1 = Timer
    
    With ActiveSheet.ChartObjects.Add(20, 200, 1100, 500).Chart
        .ChartType = xlXYScatterLinesNoMarkers
        With .SeriesCollection.NewSeries
            .AxisGroup = 1
            .Name = Cells(1, 5).Value
            .XValues = dataCheck
            .Values = dataCheck.Offset(, 2)         ' CPS
            .Border.ColorIndex = 3
            .Format.Line.Weight = 1
        End With
        .ChartType = xlXYScatterLinesNoMarkers      ' this is necessary for slow PC.
    End With
    
    With ActiveSheet.ChartObjects.Add(20, 200, 1100, 500).Chart
        .ChartType = xlXYScatterLinesNoMarkers
        With .SeriesCollection.NewSeries
            .AxisGroup = 1
            .Name = Cells(1, 6).Value
            .XValues = dataCheck
            .Values = dataCheck.Offset(, 3)         ' Ie
            .Border.ColorIndex = 4
            .Format.Line.Weight = 1
        End With
        .ChartType = xlXYScatterLinesNoMarkers
    End With

    With ActiveSheet.ChartObjects.Add(20, 200, 1100, 500).Chart
        .ChartType = xlXYScatterLinesNoMarkers
        With .SeriesCollection.NewSeries
            .AxisGroup = 1
            .Name = Cells(1, 8).Value
            .XValues = dataCheck
            .Values = dataCheck.Offset(, 5)         ' CPS/Ip
            .Border.ColorIndex = 7
            .Format.Line.Weight = 1
        End With
        .ChartType = xlXYScatterLinesNoMarkers
    End With
    
    If iRow > 3 And iRow < 7 Then
        For i = 4 To iRow
            With ActiveSheet.ChartObjects.Add(20, 200, 1100, 500).Chart
                .ChartType = xlXYScatterLinesNoMarkers
                With .SeriesCollection.NewSeries
                    .AxisGroup = 1
                    .Name = Cells(1, 5 + i).Value
                    .XValues = dataCheck
                    .Values = dataCheck.Offset(, i + 2)       ' Flux,Fluence,Ip/100mA
                    .Border.ColorIndex = i + 5
                    .Format.Line.Weight = 1
                End With
                .ChartType = xlXYScatterLinesNoMarkers
            End With
        Next
    Else
        iRow = 3
    End If
    
    For k = 0 To iRow - 1
        With ActiveSheet.ChartObjects(k + 1)
            .Top = 20 + (k * (500 / windowSize))
            With .Chart.Axes(xlValue)
                .HasTitle = True
                If k = 1 Then
                    .AxisTitle.Text = "Ie (mA)"
                ElseIf k = 2 Then
                    .AxisTitle.Text = "CPS/Ip (arb. units)"
                ElseIf k = 3 Then
                    .AxisTitle.Text = "Flux (E+9 ph)"
                ElseIf k = 4 Then
                    .AxisTitle.Text = "Fluence (E+9 ph)"
                ElseIf k = 5 Then
                    .AxisTitle.Text = "Ip/100mA (pA)"
                Else
                    .AxisTitle.Text = "CPS"
                End If
            End With
        End With
    Next
    
    For Each myChartOBJ In ActiveSheet.ChartObjects
        With myChartOBJ
            .Left = 200
            .Width = (550 * windowRatio) / windowSize
            .Height = 500 / windowSize
        End With
        With myChartOBJ.Chart.SeriesCollection.NewSeries
            .AxisGroup = 2
            .Name = Cells(1, 7).Value       'Ip
            .XValues = dataCheck
            .Values = dataCheck.Offset(, 4)
            .Border.ColorIndex = 5
            .Format.Line.Weight = 1
        End With
        With myChartOBJ.Chart.Axes(xlCategory)
            With .TickLabels
                .NumberFormatLocal = "hh:mm:ss"
            End With
            .MinimumScale = Cells(2, 3)
            .MaximumScale = Cells(totalDataPoints + 1, 3)
            .HasTitle = True
            .AxisTitle.Text = "Time"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .HasMajorGridlines = True
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With myChartOBJ.Chart.Axes(xlValue)
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With myChartOBJ.Chart.Axes(xlValue, xlSecondary)
            .HasTitle = True
            .AxisTitle.Text = "Ip (pA)"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
        End With
        With myChartOBJ.Chart.Legend
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
        With myChartOBJ.Chart
            .PlotArea.Width = (((550 * windowRatio) - 100) / windowSize)
            .ChartArea.Border.LineStyle = 0
        End With
    Next

    Cells(1, 5).Interior.ColorIndex = 3              ' CPS
    Cells(1, 6).Interior.ColorIndex = 4              ' Ie
    Cells(1, 7).Interior.ColorIndex = 5              ' Ip
    Cells(1, 8).Interior.ColorIndex = 7              ' CPS/Ip
    Cells(1, 9).Interior.ColorIndex = 9              ' Flux
    Cells(1, 10).Interior.ColorIndex = 10              ' Fluence
    Range(Cells(1, 5), Cells(1, 11)).Font.ColorIndex = 2
    Range(Cells(2, 3), Cells(totalDataPoints + 1, 3)).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
    TimeC2 = Timer
    sheetData.Activate
End Sub

Sub ScanDivide()
    iniTime = Timer
    j = 5               ' number of data to be analyzed: 1 to CPS, 2 to Ie, 3 to CPS/Ip, 4 to Flux, 5 to fluence
    iCol = 12           ' Column number to start to fill the scandivide data.
    iRow = 3            ' numnuer of graph plotted in ObbCheck.
    Range(Cells(1, iCol), Cells(totalDataPoints + 1, iCol + 1 + j)).ClearContents
CheckWf:
    wf = Cells(1, 2).End(xlDown).Value
    wf = Application.InputBox(Title:="Calc. averaged Ie and Ip", Prompt:="Input the # of scans", Default:=wf, Type:=1)
    If wf <= 0 Or Len(wf) = 0 Then
        wf = Cells(1, 2).End(xlDown).Value
    ElseIf wf > totalDataPoints Then
        wf = totalDataPoints
    End If
    
    Set rng = [A:A]
    totalDataPoints = Application.CountA(rng) - 1
    
    If wf <= 0 Or wf > totalDataPoints Then GoTo CheckWf
    k = Application.Floor(totalDataPoints / wf, 1)
    D = Range(Cells(2, 1), Cells(totalDataPoints + 1, iCol + 1 + j))
    
    For q = 0 To j - 1
        p = 1
        For i = 1 To totalDataPoints
            D(p, iCol + 1 + q) = D(i, 5 + q) + D(p, iCol + 1 + q)
            D(p, iCol) = p
            If p > wf Then
                D(p, iCol + 1 + q) = 0
                D(p, iCol) = 0
            ElseIf i = k * p Then
                If q = 4 And p = 1 Then D(p, iCol + 2 + q) = D(p, iCol + 1 + q)
                If q = 4 And p > 1 Then D(p, iCol + 2 + q) = D(p, iCol + 1 + q) + D(p - 1, iCol + 2 + q)
                D(p, iCol + 1 + q) = D(p, iCol + 1 + q) / k
                p = p + 1
            End If
        Next
    Next
    
    Range(Cells(2, 1), Cells(totalDataPoints + 1, iCol + 1 + j)) = D
    Cells(1, iCol).Value = "# norm"
    Cells(1, iCol + 1).Value = "Avg CPS"
    Cells(1, iCol + 2).Value = "Avg Ie"
    Cells(1, iCol + 3).Value = "Avg Ip"
    Cells(1, iCol + 4).Value = "Avg CPS/Ip"
    Cells(1, iCol + 5).Value = "Avg # ph./s"
    Cells(1, iCol + 6).Value = "Fluence"        ' This fluence is calculated in the summation of # of ph./s.
    
    If Cells(1, iCol).End(xlDown).Value = 0 Then
        For i = 0 To j
            Cells(1, iCol + i).End(xlDown).Value = vbNullString
        Next
    End If
    
    For i = 1 To iRow
        ActiveSheet.ChartObjects(i).Left = 2000 / windowSize
    Next
    
    Cells(1, iCol + 1).Interior.ColorIndex = 3
    Cells(1, iCol + 3).Interior.ColorIndex = 5
    Cells(1, iCol + 2).Interior.ColorIndex = 4
    Cells(1, iCol + 4).Interior.ColorIndex = 7
    Cells(1, iCol + 5).Interior.ColorIndex = 9
    Cells(1, iCol + 6).Interior.ColorIndex = 10
    Range(Cells(1, iCol + 1), Cells(1, iCol + 6)).Font.ColorIndex = 2
    Range(Cells(1, iCol + 1), Cells(1, iCol + 6)).Font.Bold = True
    finTime = Timer
End Sub

Sub EngBL()
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

    If StrComp(strTest, "GE/eV", 1) = 0 Then
        C = dataKeData                                      ' PE
        A = dataKeData.Offset(, 1)                          ' Ip
        If StrComp(Cells(1, 3).Value, "Ie", 1) = 0 Then
            D = dataKeData.Offset(, 2)                      ' Ie
        Else
            D = dataKeData.Offset(, para + 30)              ' empty Ip
        End If
        tmp = A
        startEb = Cells(2, 1).Value
        endEb = Cells(numData + 1, 1).Value
        stepEk = Cells(3, 1).Value - Cells(2, 1).Value
        g = 0
        strscanNum = 1
        maxXPSFactor = 1
    Else
        startEb = Cells(12, 1).Value
        endEb = Cells(12, 1).End(xlDown).Value
        stepEk = Abs(Cells(13, 1).Value - Cells(12, 1).Value)
        numData = ((endEb - startEb) / stepEk) + 1
        g = mid$(Cells(5, 2).Value, 1, 4)
        strscanNum = Cells(10, 2).Value
        C = Range(Cells(12, 1), Cells(12, 1).Offset(numData - 1, 0))    ' PE
        A = Range(Cells(12, 2), Cells(12, 2).Offset(numData - 1, 0))    ' Ip
        D = Range(Cells(12, 3), Cells(12, 3).Offset(numData - 1, 0))    ' Ie
        tmp = A
        maxXPSFactor = 1000000000000#
    End If
    
    If IsNumeric(strscanNum) = False Then
        MsgBox "Only first scanned data will be plotted."
        scanNum = 1
    ElseIf IsNumeric(strscanNum) = True Then
        scanNum = strscanNum
        If scanNum > 1 Then
            MsgBox "Only first scanned data will be plotted."
        End If
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
    strLabel = "Photon energy (eV)"
    str1 = "Pe"
    str2 = "Sh"
    str3 = "Ab"
    [C2:C2].Value = "eV"
    [C5:C7].Value = "eV"
    [C8:C8].Value = "times"
    [A2:A2].Interior.ColorIndex = 3
    [B2:C2].Interior.ColorIndex = 38
    [A5:A8].Interior.ColorIndex = 41
    [B5:C8].Interior.ColorIndex = 37
    [A9:A9].Interior.ColorIndex = 43
    [B9:C9].Interior.ColorIndex = 35
    
    For i = 1 To numData
        A(i, 1) = A(i, 1) * maxXPSFactor  ' pA unit
        
        If IsNumeric(D(i, 1)) = True Then
            If D(i, 1) > 0 Then
            Else
                D(i, 1) = 100
            End If
        Else
            D(i, 1) = 100
        End If
        
        tmp(i, 1) = (A(i, 1) / D(i, 1)) * 100 ' normalized to 100mA
    Next
    
    Range(Cells(11, 1), Cells(10 + numData, 1)) = C
    Range(Cells(11, 3), Cells(10 + numData, 3)) = tmp
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
        .AxisTitle.Text = strLabel
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
    Cells(9 + (imax), 1).Value = str1 + strTest
    Cells(9 + (imax), 2).Value = str2 + strTest
    Cells(9 + (imax), 3).Value = str3 + strTest

    Call SheetCheckGenerator
End Sub

Function OrderOfMagnitude(ByVal jousuu As Integer) As String
    ' order of normalization factor inputs by jousuu, then outputs unit for axis label.
    Dim tanni As String
    
    If jousuu = 3 Then
        tanni = "m"
    ElseIf jousuu = 6 Then
        tanni = "u"
    ElseIf jousuu = 9 Then
        tanni = "n"
    ElseIf jousuu = 12 Then
        tanni = "p"
    ElseIf jousuu = 15 Then
        tanni = "f"
    ElseIf jousuu = 18 Then
        tanni = "a"
    ElseIf jousuu = -3 Then
        tanni = "k"
    ElseIf jousuu = -6 Then
        tanni = "M"
    ElseIf jousuu = -9 Then
        tanni = "G"
    ElseIf jousuu = -12 Then
        tanni = "T"
    ElseIf jousuu = -15 Then
        tanni = "P"
    ElseIf jousuu = -18 Then
        tanni = "E"
    Else
        tanni = vbNullString
    End If
    
    OrderOfMagnitude = tanni
End Function

Sub PhotoBL()
    If ExistSheet(strSheetGraphName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetGraphName).Delete
        Application.DisplayAlerts = True
    End If
    
    startEb = Cells(12, 1).Value
    endEb = Cells(12, 1).End(xlDown).Value
    stepEk = Abs(Cells(13, 1).Value - Cells(12, 1).Value)
    numData = ((endEb - startEb) / stepEk) + 1
    
    If numData <= 0 Then
        strErr = "skip"
        Exit Sub
    End If
    
    g = mid$(Cells(5, 2).Value, 1, 4)
    strAES = vbNullString
    strList = vbNullString
LoopIf:
    If strAES = "EY" Then
        strList = "EY"  ' TEY mode fixed gap for second loop
        
        If ExistSheet(strSheetDataName + "_Ip") Then
            Application.DisplayAlerts = False
            Worksheets(strSheetDataName + "_Ip").Delete
            Application.DisplayAlerts = True
        End If
        If ExistSheet("Graph_" + strSheetDataName + "_Ip") Then
            Application.DisplayAlerts = False
            Worksheets("Graph_" + strSheetDataName + "_Ip").Delete
            Application.DisplayAlerts = True
        End If
SkipExtractIP:
        Set sheetData = Worksheets(strSheetDataName)
        iCol = 4
    ElseIf strList = "Ip" Then      ' In the FL mode, calc If TFY in the second loop
        strList = "FL"
        iCol = 5
        
        If ExistSheet(strSheetGraphName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetGraphName).Delete
            Application.DisplayAlerts = True
        End If
        If ExistSheet(strSheetCheckName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetCheckName).Delete
            Application.DisplayAlerts = True
        End If
        If ExistSheet(strSheetCheckName2) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetCheckName2).Delete
            Application.DisplayAlerts = True
        End If
        
        strSheetDataName = mid$(strSheetDataName, 1, Len(strSheetDataName) - 3)
        
        If ExistSheet(strSheetDataName + "_Ip") Then
            Application.DisplayAlerts = False
            Worksheets(strSheetDataName + "_Ip").Delete
            Application.DisplayAlerts = True
        End If
        If ExistSheet("Graph_" + strSheetDataName + "_Ip") Then
            Application.DisplayAlerts = False
            Worksheets("Graph_" + strSheetDataName + "_Ip").Delete
            Application.DisplayAlerts = True
        End If

        sheetData.Activate
        sheetData.Name = strSheetDataName
        
        Set sheetData = Worksheets(strSheetDataName)
        If Cells(6, 6).Value = "If" Then
            Range(Cells(7, 6), Cells(7, 7)).Value = Range(Cells(6, 6), Cells(6, 7)).Value
            Range(Cells(6, 6), Cells(6, 7)).ClearContents
        End If
        
    ElseIf strList = "Is" Then
ExtractIP:
        strList = "Ip"          ' In the FL mode, calc Is TEY and save copy with "_Ip" prior to calc If TFY
        iCol = 3
        
        If ExistSheet(strSheetGraphName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetGraphName).Delete
            Application.DisplayAlerts = True
        End If
        If ExistSheet(strSheetCheckName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetCheckName).Delete
            Application.DisplayAlerts = True
        End If
        If ExistSheet(strSheetCheckName2) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetCheckName2).Delete
            Application.DisplayAlerts = True
        End If
        
        Set sheetData = Worksheets(strSheetDataName)    ' strSheetDataName + "_Is" to be invisible
        
        If strAES = "EY" Then
            strSheetAvgName = strSheetDataName + "_Ip"
        Else
            strSheetAvgName = mid$(strSheetDataName, 1, Len(strSheetDataName) - 3) + "_Ip"
        End If

        If ExistSheet(strSheetAvgName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetAvgName).Delete
            Application.DisplayAlerts = True
        End If
        Worksheets.Add().Name = strSheetAvgName
        Set sheetAvg = Worksheets(strSheetAvgName)
        sheetAvg.Move after:=Sheets(Sheets.Count)
        sheetData.Activate
        
    ElseIf StrComp(Cells(7, 6), "If", 1) = 0 Then
        strList = "Is"          ' In the FL mode, calc Is TEY and save copy with "_Is" prior to calc If TFY
        iCol = 5
        
        If WorkbookOpen(strSheetDataName + "_Is") Then
            Application.DisplayAlerts = False
            Worksheets(strSheetDataName + "_Is").Close
            Application.DisplayAlerts = True
        End If
        
        If WorkbookOpen(strSheetDataName + "_Ip") Then
            Application.DisplayAlerts = False
            Worksheets(strSheetDataName + "_Ip").Close
            Application.DisplayAlerts = True
        End If
        
        strSheetDataName = strSheetDataName + "_Is"
        If ExistSheet(strSheetDataName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetDataName).Delete
            Application.DisplayAlerts = True
        End If

        ActiveSheet.Name = strSheetDataName
        Set sheetData = Worksheets(strSheetDataName)
        sheetData.Activate
        Range(Cells(6, 6), Cells(6, 7)).Value = Range(Cells(7, 6), Cells(7, 7)).Value
        Range(Cells(7, 6), Cells(7, 7)).ClearContents
        
    ElseIf StrComp(Cells(7, 6), "GAP", 1) = 0 Then
        strList = "VG"  ' Varied Gap XAS scanned mode
        iCol = 4
    Else
        strAES = "EY"  ' TEY mode fixed gap
        If StrComp(Cells(6, 6), "If", 1) = 0 Then
            strSheetDataName = mid$(strSheetDataName, 1, Len(strSheetDataName) - 3) + "_Is"
            GoTo SkipExtractIP
        Else
            GoTo ExtractIP
        End If
    End If
    Call ScanRangeCheck
SkipSelection2:
    sheetData.Activate
    k = 0
    Do
        strscanNum = Cells(k * (numData + 3) + 10, 2).Value
        
        If IsNumeric(strscanNum) = False Then Exit Do
        k = k + 1
    Loop
    
    scanNum = k
    strscanNum = scanNum
    scanNumR = scanNum
    C = Range(Cells((scanNum - 1) * (numData + 3) + 12, 1), Cells((scanNum - 1) * (numData + 3) + 11 + numData, 1))
    For k = 1 To numData
        If IsEmpty(C(k, 1)) = True Then
            scanNum = scanNum - 1
            If scanNum = 0 Then scanNum = 1
            GoTo SkipLastscan
        End If
    Next
SkipLastscan:
    C = Range(Cells(12, 1), Cells((scanNumR - 1) * (numData + 3) + 11 + numData, (iCol + 2)))
    A = Range(Cells(12, 1), Cells((11 + numData), 1))
    
    strSheetCheckName2 = "Photo_" + strSheetDataName        ' check multiple XAS scans
    
    If ExistSheet(strSheetCheckName2) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetCheckName2).Delete
        Application.DisplayAlerts = True
    End If
    
    Worksheets.Add().Name = strSheetCheckName2
    Set sheetCheck = Worksheets(strSheetCheckName2) ' Photo sheet
    U = Range(Cells(11, 1), Cells((numData + 3) * (iCol + 2), scanNumR + 2))
    
    For k = 1 To scanNumR
        For i = 1 To numData
            For q = 0 To iCol
                If strList = "FL" And (q = 4 Or q = 5) Then
                    U(i + ((numData + 3) * q), 1 + k) = C(i + ((numData + 3) * (k - 1)), q + 2) * -1
                Else
                    U(i + ((numData + 3) * q), 1 + k) = C(i + ((numData + 3) * (k - 1)), q + 2)
                End If
            Next
        Next
    Next

    If numscancheck <= 0 Then
        For k = 1 To scanNum
            For i = 1 To numData
                For q = 0 To iCol
                    U(i + ((numData + 3) * q), scanNumR + 2) = U(i + ((numData + 3) * q), 1 + k) + U(i + ((numData + 3) * q), scanNumR + 2)
                Next
            Next
        Next
        
        For i = 1 To numData
            For q = 0 To iCol
                U(i + ((numData + 3) * q), 1) = A(i, 1)
                If U(i + ((numData + 3) * q), scanNumR + 2) = "Inf" Or U(i + ((numData + 3) * q), scanNumR + 2) = "NaN" Then    ' if Ip = inf, set NaN in Is/Ip.
                    U(i + ((numData + 3) * q), scanNumR + 2) = vbNullString
                Else
                    U(i + ((numData + 3) * q), scanNumR + 2) = U(i + ((numData + 3) * q), scanNumR + 2) / scanNum
                End If
            Next
        Next
    ElseIf numscancheck > 0 Then
        iRow = 0
        For k = 1 To numscancheck
            If ratio(k) <= scanNumR Then
                For i = 1 To numData
                    For q = 0 To iCol
                        U(i + ((numData + 3) * q), scanNumR + 2) = U(i + ((numData + 3) * q), 1 + ratio(k)) + U(i + ((numData + 3) * q), scanNumR + 2)
                    Next
                Next
            Else
                iRow = iRow + 1
            End If
        Next
        
        If numscancheck - iRow <= 0 Then
            numscancheck = 0
            GoTo SkipSelection2
        End If
        
        For i = 1 To numData
            For q = 0 To iCol
                U(i + ((numData + 3) * q), 1) = A(i, 1)
                U(i + ((numData + 3) * q), scanNumR + 2) = U(i + ((numData + 3) * q), scanNumR + 2) / (numscancheck - iRow)
            Next
        Next
        scanNum = (numscancheck - iRow)
    End If
    
    Range(Cells(11, 1), Cells((numData + 3) * (iCol + 2), scanNumR + 2)) = U
    D = Range(Cells(11 + ((numData + 3) * 0), scanNumR + 2), Cells(11 + ((numData + 3) * 0), scanNumR + 2).Offset(numData - 1, 0))    ' Ie
    b = Range(Cells(11 + ((numData + 3) * 1), scanNumR + 2), Cells(11 + ((numData + 3) * 1), scanNumR + 2).Offset(numData - 1, 0))   ' Ip
    C = Range(Cells(11 + ((numData + 3) * 2), scanNumR + 2), Cells(11 + ((numData + 3) * 2), scanNumR + 2).Offset(numData - 1, 0))   ' Is
    tmp = Range(Cells(11 + ((numData + 3) * 3), scanNumR + 2), Cells(11 + ((numData + 3) * 3), scanNumR + 2).Offset(numData - 1, 0))    ' Is/Ip
    U = Range(Cells(11 + ((numData + 3) * 4), scanNumR + 2), Cells(11 + ((numData + 3) * 4), scanNumR + 2).Offset(numData - 1, 0))    ' Gap or If
    en = Range(Cells(11 + ((numData + 3) * 5), scanNumR + 2), Cells(11 + ((numData + 3) * 5), scanNumR + 2).Offset(numData - 1, 0))    ' If/Ip
        
    If strList = "FL" Then      ' Electrometer DC current mode
        If Application.WorksheetFunction.Average(en) < 0 Then
            strList = "-FL"     ' MCP counting mode
        End If
    ElseIf strList = "Ip" Then
    ElseIf strList = "Is" Then
        If Application.WorksheetFunction.Average(U) < 0 Then
            strList = "-Is"     ' MCP counting mode
        End If
        iCol = 4
    End If
    
    dblMax1 = Application.Max(tmp)
    dblMin1 = Application.Min(tmp)
    
    If StrComp(testMacro, "debug", 1) = 0 Then
    ElseIf strList = "Is" Or strList = "-Is" Or strList = "Ip" Then
    Else
        If strscanNum > scanNum Then
            MsgBox "# of scanned data: " & scanNum & "/" & strscanNum & " in " & strTest & " will be analyzed."
        End If
    End If
    
    strSheetGraphName = "Graph_" + strSheetDataName
    Worksheets.Add().Name = strSheetGraphName
    Set sheetGraph = Worksheets(strSheetGraphName)
    sheetGraph.Activate
    Cells(2, 1).Value = "PE shifts"
    Cells(2, 2).Value = pe
    Cells(5, 1).Value = "Start PE"
    Cells(6, 1).Value = "End PE"
    Cells(7, 1).Value = "Step PE"
    Cells(8, 1).Value = "# scan"
    Cells(5, 2).Value = startEb
    Cells(6, 2).Value = endEb
    Cells(7, 2).Value = stepEk
    If numscancheck <= 0 Then
        Cells(8, 2).Value = scanNum
        [C8:C8].Value = "times"
        [B5:C8].Interior.ColorIndex = 37
    ElseIf numscancheck > 0 Then
        Cells(8, 2).Value = strTest
        [C8:C8].Value = vbNullString
        [B5:C7].Interior.ColorIndex = 37
        [B8:C8].Interior.ColorIndex = 33
    End If
    
    Cells(9, 1).Value = "Offset/multp"
    Cells(9, 2).Value = off
    Cells(9, 3).Value = multi
    Cells(10, 1).Value = "PE"
    Cells(10, 2).Value = "+shift"
    Cells(10, 3).Value = "Ab"
    strLabel = "Photon energy (eV)"
    str1 = "Pe"
    str2 = "Sh"
    str3 = "Ab"
    [C2:C2].Value = "eV"
    [C5:C7].Value = "eV"
    [A2:A2].Interior.ColorIndex = 3
    [B2:C2].Interior.ColorIndex = 38
    [A5:A8].Interior.ColorIndex = 41
    Range(Cells(9, 1), Cells(9, 1)).Interior.ColorIndex = 43
    Range(Cells(9, 2), Cells(9, 3)).Interior.ColorIndex = 35
    
    For i = 1 To numData
        If IsNumeric(D(i, 1)) = True Then
            If D(i, 1) > 0 Then
            Else
                D(i, 1) = 100
            End If
        Else
            D(i, 1) = 100
        End If
        C(i, 1) = (C(i, 1) * 1000000000000# / D(i, 1)) * 100 ' Is normalized to 100mA of Ie (unit: pA)
        If strList = "Ip" Then
        
        ElseIf strList = "FL" Or strList = "Is" Then
            U(i, 1) = (U(i, 1) / D(i, 1)) * 100 ' If normalized to 100mA of Ie (unit: counts)
        ElseIf strList = "-FL" Or strList = "-Is" Then
            U(i, 1) = (U(i, 1) * 1000000000000# / D(i, 1)) * -100 ' If normalized to 100mA of Ie (unit: pA)
            en(i, 1) = en(i, 1) * -1   ' If/Ip
        End If
        If b(i, 1) = 0 Then
            b(i, 1) = 1     ' if Ip = 0 the set Ip = 1 pico amps.
        Else
            b(i, 1) = (b(i, 1) * 1000000000000# / D(i, 1)) * 100 ' Ip normalized to 100mA of Ie (unit: pA)
        End If
    Next
    
    If strList = "-FL" Then strList = "FL"
    If strList = "-Is" Then strList = "Is"

    Range(Cells(11, 1), Cells(10 + numData, 1)) = A
    If strList = "FL" Then
        Range(Cells(11, 3), Cells(10 + numData, 3)) = en
        dblMax2 = Application.Max(en)
        dblMin2 = Application.Min(en)
        dblMax = (dblMax2 - Cells(9, 2).Value) * Cells(9, 3).Value
        dblMin = (dblMin2 - Cells(9, 2).Value) * Cells(9, 3).Value
    ElseIf strList = "Ip" Then
        Range(Cells(11, 3), Cells(10 + numData, 3)) = b
    Else
        Range(Cells(11, 3), Cells(10 + numData, 3)) = tmp
        dblMax = (dblMax1 - Cells(9, 2).Value) * Cells(9, 3).Value
        dblMin = (dblMin1 - Cells(9, 2).Value) * Cells(9, 3).Value
    End If
    
    Cells(11, 2).FormulaR1C1 = "=R2C2 + RC[-1]"
    Range(Cells(11, 2), Cells(10 + numData, 2)).FillDown
    imax = numData + 10
    Cells(10 + (imax), 1).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
    Range(Cells(10 + (imax), 1), Cells((2 * imax) - 1, 1)).FillDown
    Cells(10 + (imax), 2).FormulaR1C1 = "=R2C + R[-" & (imax - 1) & "]C[-1]"
    Range(Cells(10 + (imax), 2), Cells((2 * imax) - 1, 2)).FillDown
    Cells(10 + (imax), 3).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C[-1]) * R9C"
    Range(Cells(10 + (imax), 3), Cells((2 * imax) - 1, 3)).FillDown
    Set dataData = Range(Cells(10 + (imax), 2), Cells(10 + (imax), 2).Offset(numData - 1, 1))
    startEb = Cells(10 + (imax), 2).Value
    endEb = Cells(9 + (imax) + numData, 2).Value
    
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers 'xlXYScatterSmoothNoMarkers
    ActiveChart.SetSourceData Source:=dataData, PlotBy:=xlColumns
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetGraphName
    ActiveChart.SeriesCollection(1).Name = ActiveWorkbook.Name '"xas"
    ActiveChart.ChartTitle.Delete

    With ActiveChart.Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = strLabel
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .MinimumScale = startEb
        .MaximumScale = endEb
    End With
    With ActiveChart.Axes(xlValue)
        .HasTitle = True
        If strList = "FL" Then
            .AxisTitle.Text = "If normalized by Ip (arb. unit)"
            .MinimumScale = dblMin
            .MaximumScale = dblMax
        Else
            .AxisTitle.Text = "Is normalized by Ip (arb. unit)"
            .MinimumScale = dblMin
            .MaximumScale = dblMax
        End If
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
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
    
    SourceRangeColor1 = ActiveChart.SeriesCollection(1).Border.Color
    Range(Cells(10, 1), Cells(10, 2)).Interior.Color = SourceRangeColor1
    Range(Cells(9 + (imax), 1), Cells(9 + (imax), 2)).Interior.Color = SourceRangeColor1
    strTest = mid$(strSheetGraphName, InStr(strSheetGraphName, "_") + 1, Len(strSheetGraphName) - 6)
    Cells(8 + (imax), 2).Value = strTest + ".xlsx"
    Cells(9 + (imax), 1).Value = str1 + strTest
    Cells(9 + (imax), 2).Value = str2 + strTest
    Cells(9 + (imax), 3).Value = str3 + strTest
    
    If strList = "Ip" Then
        sheetAvg.Activate
        Cells(1, 1).Value = "PE/eV"         ' you can fit
        Cells(1, 2).Value = "Ip"
        Range(Cells(2, 1), Cells(1 + numData, 1)) = A
        Range(Cells(2, 2), Cells(1 + numData, 2)) = b
        sheetGraph.Activate
        sheetGraph.Name = "Graph_" + strSheetAvgName
        sheetData.Activate
        sheetData.Visible = False
        
        GoTo skipchkphotobl
    End If

    If NoCheck = "ON" Then
        Application.DisplayAlerts = False
        Worksheets(strSheetCheckName2).Delete
        Application.DisplayAlerts = True
        GoTo skipchkphotobl
    End If

    sheetCheck.Activate     ' Photo sheet for multiple scanned data analysis
    
    For k = 0 To iCol
        Set dataCheck = Range(Cells(11 + (k * (3 + numData)), 1), Cells(10 + numData + (k * (3 + numData)), (scanNumR + 1)))
        
        Charts.Add
        ActiveChart.ChartType = xlXYScatterLinesNoMarkers
        ActiveChart.SetSourceData Source:=dataCheck, PlotBy:=xlColumns
        ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetCheckName2
        ActiveSheet.ChartObjects(1 + k).Top = 20 + (k * (500 / windowSize))
        
        With ActiveChart.Axes(xlValue)
            .HasTitle = True
            If k = 0 Then
                .AxisTitle.Text = "Ie (mA)"
            ElseIf k = 1 Then
                .AxisTitle.Text = "Ip (pA)"
            ElseIf k = 2 Then
                .AxisTitle.Text = "Is (pA)"
            ElseIf k = 3 Then
                .AxisTitle.Text = "Is/Ip (arb. units)"
            ElseIf k = 4 Then
                If strList = "FL" Then
                    .AxisTitle.Text = "If (pA)"
                ElseIf strList = "Is" Then
                    .AxisTitle.Text = "If/Ip (arb. units)"
                Else
                    .AxisTitle.Text = "Upstream U60 gap (mm)"
                End If
            ElseIf k = 5 Then
                If strList = "FL" Then
                    .AxisTitle.Text = "If/Ip (arb. units)"
                End If
            End If
        End With
        
        For Each mySeries In ActiveChart.SeriesCollection
            mySeries.Format.Line.Weight = 1
            mySeries.ChartType = xlXYScatterLinesNoMarkers
        Next
    Next
    
    For Each myChartOBJ In ActiveSheet.ChartObjects
        With myChartOBJ
            .Left = 200
            .Width = (550 * windowRatio) / windowSize
            .Height = 500 / windowSize
        End With

        With myChartOBJ.Chart.Axes(xlCategory, xlPrimary)
            .MinorTickMark = xlOutside
            .MinimumScale = startEb
            .MaximumScale = endEb
            .HasTitle = True
            .AxisTitle.Text = "Photon energy (eV)"
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .HasMajorGridlines = True
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With myChartOBJ.Chart.Axes(xlValue)
            .AxisTitle.Font.Size = 12
            .AxisTitle.Font.Bold = False
            .MajorGridlines.Border.LineStyle = xlDot
        End With
        With myChartOBJ.Chart.Legend
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
            
        With myChartOBJ.Chart
            .PlotArea.Width = (((550 * windowRatio) - 60) / windowSize)
            .ChartArea.Border.LineStyle = 0
        End With
    Next
                
    For k = 0 To iCol
        Cells(k * (numData + 3) + 10, 1).Value = "PE"
        Cells(k * (numData + 3) + 10, 1).Interior.ColorIndex = 16
        Cells(k * (numData + 3) + 10, 2 + scanNumR).Value = "Avg"
        Cells(k * (numData + 3) + 10, 2 + scanNumR).Interior.ColorIndex = 15
    Next
    
    For i = 1 To scanNumR
        Cells(0 * (numData + 3) + 10, 1 + i).Interior.Color = ActiveSheet.ChartObjects(1).Chart.SeriesCollection(i).Border.Color
        Cells(0 * (numData + 3) + 10, 1 + i).Value = "Ie #" + CStr(i)
        ActiveSheet.ChartObjects(1).Chart.SeriesCollection(i).Name = Cells(0 * (numData + 3) + 10, 1 + i).Value
    Next
    
    For i = 1 To scanNumR
        Cells(1 * (numData + 3) + 10, 1 + i).Interior.Color = ActiveSheet.ChartObjects(2).Chart.SeriesCollection(i).Border.Color
        Cells(1 * (numData + 3) + 10, 1 + i).Value = "Ip #" + CStr(i)
        ActiveSheet.ChartObjects(2).Chart.SeriesCollection(i).Name = Cells(1 * (numData + 3) + 10, 1 + i).Value
    Next
    
    For i = 1 To scanNumR
        Cells(2 * (numData + 3) + 10, 1 + i).Interior.Color = ActiveSheet.ChartObjects(3).Chart.SeriesCollection(i).Border.Color
        Cells(2 * (numData + 3) + 10, 1 + i).Value = "Is #" + CStr(i)
        ActiveSheet.ChartObjects(3).Chart.SeriesCollection(i).Name = Cells(2 * (numData + 3) + 10, 1 + i).Value
    Next
    
    For i = 1 To scanNumR
        Cells(3 * (numData + 3) + 10, 1 + i).Interior.Color = ActiveSheet.ChartObjects(4).Chart.SeriesCollection(i).Border.Color
        Cells(3 * (numData + 3) + 10, 1 + i).Value = "Is/Ip #" + CStr(i)
        ActiveSheet.ChartObjects(4).Chart.SeriesCollection(i).Name = Cells(3 * (numData + 3) + 10, 1 + i).Value
    Next
    
    For i = 1 To scanNumR
        Cells(4 * (numData + 3) + 10, 1 + i).Interior.Color = ActiveSheet.ChartObjects(5).Chart.SeriesCollection(i).Border.Color
        If strList = "FL" Or strList = "Is" Or strList = "Ip" Then
            Cells(4 * (numData + 3) + 10, 1 + i).Value = "If #" + CStr(i)
            ActiveSheet.ChartObjects(5).Chart.SeriesCollection(i).Name = Cells(4 * (numData + 3) + 10, 1 + i).Value
        Else
            Cells(4 * (numData + 3) + 10, 1 + i).Value = "Gap #" + CStr(i)
            ActiveSheet.ChartObjects(5).Chart.SeriesCollection(i).Name = Cells(4 * (numData + 3) + 10, 1 + i).Value
        End If
    Next
    
    If strList = "FL" Then
        For i = 1 To scanNumR
            Cells(5 * (numData + 3) + 10, 1 + i).Interior.Color = ActiveSheet.ChartObjects(6).Chart.SeriesCollection(i).Border.Color
            Cells(5 * (numData + 3) + 10, 1 + i).Value = "If/Ip #" + CStr(i)
            ActiveSheet.ChartObjects(6).Chart.SeriesCollection(i).Name = Cells(5 * (numData + 3) + 10, 1 + i).Value
        Next
    End If
    
    If ExistSheet(strSheetCheckName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetCheckName).Delete
        Application.DisplayAlerts = True
    End If
    
    strSheetCheckName = "Check_" + strSheetDataName

    Worksheets.Add().Name = strSheetCheckName       ' Check sheet for averaged data analysis
    Set sheetCheck = Worksheets(strSheetCheckName)
    Range(Cells(11, 1), Cells(10 + numData, 1)) = D 'Ie
    Range(Cells(11, 2), Cells(10 + numData, 2)) = A 'Pe
    Range(Cells(11, 3), Cells(10 + numData, 3)) = b 'Ip
    Range(Cells(11, 4), Cells(10 + numData, 4)) = A 'Pe
    Range(Cells(11, 5), Cells(10 + numData, 5)) = C 'Is
    Range(Cells(11, 6), Cells(10 + numData, 6)) = tmp 'Is/Ip
    
    If strList = "FL" Then
        Range(Cells(11, 8), Cells(10 + numData, 8)) = A 'Pe
        Range(Cells(11, 9), Cells(10 + numData, 9)) = U 'If
        Range(Cells(11, 10), Cells(10 + numData, 10)) = en  'If/Ip
    ElseIf strList = "Is" Then
        Range(Cells(11, 8), Cells(10 + numData, 8)) = en  'If/Ip
    ElseIf strList = "Ip" Then
        Range(Cells(11, 8), Cells(10 + numData, 8)) = b  'Ip
    Else
        Range(Cells(11, 8), Cells(10 + numData, 8)) = U ' Gap
    End If
    
    If strList = "EY" Or strList = "Is" Then 'Cells(11, 8).Value = 0 Then      ' when fixed gap mode
        wf = 26.5
        k = 0
        If StrComp(testMacro, "debug", 1) = 0 Then
        Else
            wf = Application.InputBox(Title:="Calc. 1st har. PE", Prompt:="Input the U60 gap: mm", Default:=wf, Type:=1)
            If wf <= 0 Or Len(wf) = 0 Then
                k = 1
                char = Application.InputBox(Title:="Calc. U60 gap", Prompt:="Input the 1st har. photon energy: eV", Default:=char, Type:=1)
                If char < 20 Or char > 300 Then char = 40
            ElseIf wf < 25 Or wf > 200 Then
                wf = 26.5
            End If
        End If
    
        If k = 0 Then
            Cells(4, 3).Value = wf
            Cells(4, 6).FormulaR1C1 = "= " & a0 & " + " & a1 & " * Exp(" & a2 & " * R4C3)"    ' B (T)
            Cells(5, 6).FormulaR1C1 = "= 0.934 * " & lambda & " * (R4C6)"                          ' K
            Cells(5, 3).FormulaR1C1 = "=950 * ((" & gamma & ") ^ 2) / (((((R5C6) ^ 2) / 2) + 1) * " & lambda & ")" ' 1st har.
            char = Cells(5, 3).Value
            [B5:B5].Interior.ColorIndex = 45
            [C5:D5].Interior.ColorIndex = 44
            [E5:E5].Interior.ColorIndex = 45
            [F5:F5].Interior.ColorIndex = 44
            [B4:B4].Interior.ColorIndex = 18
            [C4:D4].Interior.ColorIndex = 38
            [E4:E4].Interior.ColorIndex = 18
            [F4:F4].Interior.ColorIndex = 38
        Else
            Cells(5, 3).Value = char
            Cells(5, 6).FormulaR1C1 = "= Sqrt((((950 *((" & gamma & ")^2))/(R5C3 * " & lambda & "))-1) * 2)"    ' K
            Cells(4, 6).FormulaR1C1 = "= R5C6/(" & lambda & " * 0.934)"                            ' B (T)
            Cells(4, 3).FormulaR1C1 = "=(Ln((R4C6 - " & a0 & ")/(" & a1 & ")))/(" & a2 & ")"   ' 1st har.
            wf = Cells(4, 3).Value
            [B5:B5].Interior.ColorIndex = 45
            [C5:D5].Interior.ColorIndex = 44
            [E5:E5].Interior.ColorIndex = 45
            [F5:F5].Interior.ColorIndex = 44
            [B4:B4].Interior.ColorIndex = 18
            [C4:D4].Interior.ColorIndex = 38
            [E4:E4].Interior.ColorIndex = 18
            [F4:F4].Interior.ColorIndex = 38
        End If
        Cells(4, 2).Value = "U60 gap"
        Cells(4, 4).Value = "mm"
        Cells(5, 2).Value = "1st har."
        Cells(5, 4).Value = "eV"
        Cells(4, 5).Value = "B (T)"
        Cells(5, 5).Value = "K"
    End If
    
    Cells(10, 1).Value = "Ie (mA)"
    Cells(10, 2).Value = "Pe (eV; Ip)"
    Cells(10, 3).Value = "Ip (pA/100mA)"
    Cells(10, 4).Value = "Pe (eV; Is)"
    Cells(10, 5).Value = "Is (pA/100mA)"
    Cells(10, 6).Value = "Is/Ip"
    Cells(10, 7).Value = "Calc. Is/Ip"
    If strList = "FL" Then
        Cells(10, 8).Value = "Pe (eV; If)"
        Cells(10, 9).Value = "If (pA/100mA)"
        Cells(10, 10).Value = "If/Ip"
        Cells(10, 11).Value = "Calc. If/Ip"
    ElseIf strList = "Is" Then
        Cells(10, 8).Value = "If/Ip"
    Else
        Cells(10, 8).Value = "Gap (mm)"
    End If
    Cells(2, 1).Value = "PE shift"
    Cells(2, 3).Value = "PE shift"
    Cells(7, 2).Value = "Offset"
    Cells(8, 2).Value = "Multiple"
    Cells(7, 4).Value = "Offset"
    Cells(8, 4).Value = "Multiple"
    Cells(2, 5).Value = "eV"
    If strList = "FL" Then
        Cells(2, 7).Value = "PE shift"
        Cells(2, 9).Value = "eV"
        Cells(7, 8).Value = "Offset"
        Cells(8, 8).Value = "Multiple"
        Cells(7, 12).Value = "amps."
        Cells(8, 12).Value = "times"
    ElseIf strList = "Is" Then
        Cells(7, 9).Value = "amps."
        Cells(8, 9).Value = "times"
    Else
        Cells(7, 8).Value = "amps."
        Cells(8, 8).Value = "times"
    End If
    
    [B2:B2].Value = 0
    [D2:D2].Value = 0
    [C7:C7].Value = 0
    [E7:G7].Value = 0
    [C8:C8].Value = 1
    [E8:G8].Value = 1
    If strList = "FL" Then
        [H2:H2].Value = 0
        [I7:K7].Value = 0
        [I8:K8].Value = 1
    ElseIf strList = "Is" Then
        [H7:H7].Value = 0
        [H8:H8].Value = 1
    End If
    
    Set dataIntCheck = Range(Cells(11, 1), Cells(numData + 10, 1))
    chkMax = Application.Max(C)
    chkMin = Application.Min(C)
    
    For k = 1 To numData
        If tmp(k, 1) = vbNullString Then
            tmp(k, 1) = 1
        Else
            tmp(k, 1) = (((tmp(k, 1) - dblMin) / (dblMax - dblMin)) * (chkMax - chkMin)) + chkMin
        End If
    Next
    dataIntCheck.Offset(0, 5) = tmp
    
    If strList = "FL" Then
        chkMax2 = Application.Max(U)
        chkMin2 = Application.Min(U)
        dblMax2 = Application.Max(en)
        dblMin2 = Application.Min(en)
        For k = 1 To numData
            en(k, 1) = (((en(k, 1) - dblMin2) / (dblMax2 - dblMin2)) * (chkMax2 - chkMin2)) + chkMin2
        Next
        dataIntCheck.Offset(0, 9) = en
    ElseIf strList = "Is" Then
        chkMax2 = Application.Max(tmp)
        chkMin2 = Application.Min(tmp)
        
        dblMax2 = Application.Max(en)
        dblMin2 = Application.Min(en)
        For k = 1 To numData
            en(k, 1) = (((en(k, 1) - dblMin2) / (dblMax2 - dblMin2)) * (chkMax2 - chkMin2)) + chkMin2
        Next
    Else
    End If
    
    Cells(11, 7).FormulaR1C1 = "=(((((RC[-2]/RC[-4]) - " & (dblMin) & ")/ " & (dblMax - dblMin) & ")* " & (chkMax - chkMin) & ")+ " & (chkMin) & ")"
    Range(Cells(11, 7), Cells(numData + 10, 7)).FillDown
    Cells(10 + (imax), 2).FormulaR1C1 = "=R[-" & (imax - 1) & "]C - R2C"
    Range(Cells(10 + (imax), 2), Cells((2 * imax) - 1, 2)).FillDown
    Cells(10 + (imax), 3).FormulaR1C1 = "=(R[-" & (imax - 1) & "]C - R7C ) * R8C"
    Range(Cells(10 + (imax), 3), Cells((2 * imax) - 1, 3)).FillDown
    Cells(10 + (imax), 4).FormulaR1C1 = "=R[-" & (imax - 1) & "]C - R2C"
    Range(Cells(10 + (imax), 4), Cells((2 * imax) - 1, 4)).FillDown
    Cells(10 + (imax), 5).FormulaR1C1 = "=(R[-" & (imax - 1) & "]C - R7C ) * R8C"
    Range(Cells(10 + (imax), 5), Cells((2 * imax) - 1, 5)).FillDown
    Cells(10 + (imax), 6).FormulaR1C1 = "=(R[-" & (imax - 1) & "]C - R7C ) * R8C"
    Range(Cells(10 + (imax), 6), Cells((2 * imax) - 1, 6)).FillDown
    Cells(10 + (imax), 7).FormulaR1C1 = "=((((((RC[-2]/RC[-4]) - " & (dblMin) & ")/ " & (dblMax - dblMin) & ")* " & (chkMax - chkMin) & ")+ " & (chkMin) & ") - R7C ) * R8C"
    Range(Cells(10 + (imax), 7), Cells((2 * imax) - 1, 7)).FillDown
    
    If strList = "FL" Then
        Cells(11, 11).FormulaR1C1 = "=(((((RC[-2]/RC[-8]) - " & (dblMin2) & ")/ " & (dblMax2 - dblMin2) & ")* " & (chkMax2 - chkMin2) & ")+ " & (chkMin2) & ")"
        Range(Cells(11, 11), Cells(numData + 10, 11)).FillDown
        Cells(10 + (imax), 11).FormulaR1C1 = "=((((((RC[-2]/RC[-8]) - " & (dblMin2) & ")/ " & (dblMax2 - dblMin2) & ")* " & (chkMax2 - chkMin2) & ")+ " & (chkMin2) & ") - R7C ) * R8C"
        Range(Cells(10 + (imax), 11), Cells((2 * imax) - 1, 11)).FillDown
        Cells(10 + (imax), 8).FormulaR1C1 = "=R[-" & (imax - 1) & "]C - R2C"
        Range(Cells(10 + (imax), 8), Cells((2 * imax) - 1, 8)).FillDown
        Cells(10 + (imax), 9).FormulaR1C1 = "=(R[-" & (imax - 1) & "]C - R7C ) * R8C"
        Range(Cells(10 + (imax), 9), Cells((2 * imax) - 1, 9)).FillDown
        Cells(10 + (imax), 10).FormulaR1C1 = "=(R[-" & (imax - 1) & "]C - R7C ) * R8C"
        Range(Cells(10 + (imax), 10), Cells((2 * imax) - 1, 10)).FillDown
    ElseIf strList = "Is" Then
        Cells(10 + (imax), 8).FormulaR1C1 = "=(R[-" & (imax - 1) & "]C - R7C ) * R8C"
        Range(Cells(10 + (imax), 8), Cells((2 * imax) - 1, 8)).FillDown
    End If
    
    Set dataCheck = Range(Cells(10 + imax, 2), Cells((2 * imax) - 1, 2))
    Set dataIntCheck = Range(Cells(11, 1), Cells(numData + 10, 1))
    
    Charts.Add
    ActiveChart.ChartType = xlXYScatterLinesNoMarkers
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetCheckName
    
    With ActiveChart.SeriesCollection(1)
        .XValues = dataCheck
        .Values = dataCheck.Offset(0, 1)
        .AxisGroup = xlPrimary
        .Border.ColorIndex = 41
        .Name = "Ip"
    End With
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(2)
        .XValues = dataCheck.Offset(0, 2)
        .Values = dataCheck.Offset(0, 3)
        .AxisGroup = xlSecondary
        .Border.ColorIndex = 45
        .Name = "Is"
    End With
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(3)
        .XValues = dataCheck.Offset(0, 2)
        .Values = dataCheck.Offset(0, 4)
        .AxisGroup = xlSecondary
        .Border.ColorIndex = 3
        .Name = "Is/Ip"
    End With

    With ActiveChart.Axes(xlCategory, xlPrimary)
        .MinorTickMark = xlOutside
        If strList = "EY" Or strList = "Is" Then
            .MinimumScale = Application.Floor(startEb, Cells(5, 3).Value)
            .MaximumScale = Application.Ceiling(endEb, Cells(5, 3).Value)
        Else
            .MinimumScale = startEb
            .MaximumScale = endEb
        End If
        .HasTitle = True
        .AxisTitle.Text = "Photon energy (eV)"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .HasMajorGridlines = True
        If strList = "EY" Or strList = "Is" Then
            .MajorUnit = Cells(5, 3).Value
        End If
    End With
    
    With ActiveChart.Axes(xlValue, xlSecondary)
        .HasTitle = True
        .AxisTitle.Text = "Is (pA/100mA) & Is/Ip (arb. unit)"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .MinimumScale = chkMin
        .MaximumScale = chkMax
    End With
    
    chkMax = Application.Max(b)
    chkMin = Application.Min(b)
    
    With ActiveChart.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Ip (pA/100mA)"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .MinimumScale = chkMin
        .MaximumScale = chkMax
    End With
    
    Charts.Add
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetCheckName
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(1)
        .ChartType = xlXYScatter
        .XValues = dataCheck.Offset(0, 2)
        .Values = dataCheck.Offset(0, 4)
        .AxisGroup = 1
        .ChartType = xlXYScatterLinesNoMarkers
        .Border.ColorIndex = 3
        .Border.Weight = xlThin
        .Name = "Is/Ip"
    End With
    
    With ActiveChart.SeriesCollection(2)
        .ChartType = xlXYScatterLinesNoMarkers
        .XValues = dataCheck.Offset(0, 2)
        .Values = dataCheck.Offset(0, 5)
        .AxisGroup = 1
        .Border.ColorIndex = 26
        .Name = "cIs/Ip"
    End With
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(3)
        .ChartType = xlXYScatterLinesNoMarkers
        .XValues = dataCheck
        .Values = dataIntCheck
        .AxisGroup = 2
        .Border.ColorIndex = 4
        .Name = "Ie"
    End With
    
    If strList = "FL" Then
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(4)
            .ChartType = xlXYScatter
            .XValues = dataCheck.Offset(0, 6)
            .Values = dataCheck.Offset(0, 8)
            .AxisGroup = 1
            .ChartType = xlXYScatterLinesNoMarkers
            .Border.ColorIndex = 31
            .Border.Weight = xlThin
            .Name = "If/Ip"
        End With
        
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(5)
            .ChartType = xlXYScatter
            .XValues = dataCheck.Offset(0, 6)
            .Values = dataCheck.Offset(0, 9)
            .AxisGroup = 1
            .ChartType = xlXYScatterLinesNoMarkers
            .Border.ColorIndex = 33
            .Name = "cIf/Ip"
        End With
    ElseIf strList = "Is" Then
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(4)
            .ChartType = xlXYScatterLinesNoMarkers
            .XValues = dataCheck
            .Values = dataCheck.Offset(0, 6)
            .AxisGroup = 1
            .Border.ColorIndex = 1
            .Name = "If/Ip"
        End With
    Else
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(4)
            .ChartType = xlXYScatterLinesNoMarkers
            .XValues = dataCheck
            .Values = dataIntCheck.Offset(0, 7)
            .AxisGroup = 2
            .Border.ColorIndex = 1
            .Name = "gap"
        End With
    End If

    With ActiveChart.Axes(xlCategory, xlPrimary)
        .MinorTickMark = xlOutside
        If strList = "EY" Or strList = "Is" Then
            .MinimumScale = Application.Floor(startEb, Cells(5, 3).Value)
            .MaximumScale = Application.Ceiling(endEb, Cells(5, 3).Value)
        Else
            .MinimumScale = startEb
            .MaximumScale = endEb
        End If
        .HasTitle = True
        .AxisTitle.Text = "Photon energy (eV)"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .HasMajorGridlines = True
        If strList = "EY" Or strList = "Is" Then
            .MajorUnit = Cells(5, 3).Value
        End If
    End With
    
    chkMax = Application.Max(dataCheck.Offset(0, 5))
    chkMin = Application.Min(dataCheck.Offset(0, 5))
    If chkMax < Application.Max(dataCheck.Offset(0, 8)) Then
        chkMax = Application.Max(dataCheck.Offset(0, 8))
    End If
    If chkMin > Application.Min(dataCheck.Offset(0, 8)) Then
        chkMin = Application.Min(dataCheck.Offset(0, 8))
    End If
    
    With ActiveChart.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Yield (arb. units)"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .MinimumScale = chkMin
        .MaximumScale = chkMax
    End With
    
    With ActiveChart.Axes(xlValue, xlSecondary)
        .HasTitle = True
        If strList <> "VG" Then
            .AxisTitle.Text = "Ie (mA)"
        Else
            .AxisTitle.Text = "Ie (mA) & Gap (mm)"
        End If
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .MajorGridlines.Border.LineStyle = xlDot
        .MinimumScale = 20
        .MaximumScale = 160
    End With
    
    With ActiveSheet.ChartObjects(1)
        .Top = 150
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
            .PlotArea.Width = (((550 * windowRatio) - 100) / windowSize)
            .ChartArea.Border.LineStyle = 0
        End With
    End With
    
    Charts.Add
    ActiveChart.Location Where:=xlLocationAsObject, Name:=strSheetCheckName
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(1)
        .ChartType = xlXYScatterLinesNoMarkers
        .XValues = dataCheck
        .Values = dataCheck.Offset(0, 1)
        .AxisGroup = xlPrimary
        .Border.ColorIndex = 41
        .Name = "Ip"
    End With
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(2)
        .ChartType = xlXYScatterLinesNoMarkers
        .AxisGroup = xlSecondary
        .Border.ColorIndex = 47
        If strList = "FL" Then
            .Name = "If"
            .XValues = dataCheck.Offset(0, 6)
            .Values = dataCheck.Offset(0, 7)
        Else
            .Name = "Is"
            .XValues = dataCheck.Offset(0, 2)
            .Values = dataCheck.Offset(0, 3)
        End If
    End With
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(3)
        .ChartType = xlXYScatterLinesNoMarkers
        .AxisGroup = xlSecondary
        .Border.ColorIndex = 31
        .Name = "If/Ip"
        If strList = "FL" Then
            .Name = "If/Ip"
            .XValues = dataCheck.Offset(0, 6)
            .Values = dataCheck.Offset(0, 8)
        Else
            .Name = "Is/Ip"
            .XValues = dataCheck.Offset(0, 2)
            .Values = dataCheck.Offset(0, 4)
        End If
    End With

    With ActiveChart.Axes(xlCategory, xlPrimary)
        .MinorTickMark = xlOutside
        If strList = "EY" Or strList = "Is" Then
            .MinimumScale = Application.Floor(startEb, Cells(5, 3).Value)
            .MaximumScale = Application.Ceiling(endEb, Cells(5, 3).Value)
        Else
            .MinimumScale = startEb
            .MaximumScale = endEb
        End If
        .HasTitle = True
        .AxisTitle.Text = "Photon energy (eV)"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .HasMajorGridlines = True
        If strList = "EY" Or strList = "Is" Then
            .MajorUnit = Cells(5, 3).Value
        End If
    End With
    
    If strList = "FL" Then
        dblMax2 = Application.Max(en)
        dblMin2 = Application.Min(en)
    Else
        dblMax2 = Application.Max(tmp)
        dblMin2 = Application.Min(tmp)
    End If
    
    With ActiveChart.Axes(xlValue, xlSecondary)
        .HasTitle = True
        If strList = "FL" Then
            .AxisTitle.Text = "If (pA/100mA) & If/Ip (arb. unit)"
        Else
            .AxisTitle.Text = "Is (pA/100mA) & Is/Ip (arb. unit)"
        End If
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .MinimumScale = dblMin2
        .MaximumScale = dblMax2

    End With
    
    chkMax = Application.Max(b)
    chkMin = Application.Min(b)
    
    With ActiveChart.Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Ip (pA/100mA)"
        .AxisTitle.Font.Size = 12
        .AxisTitle.Font.Bold = False
        .MinimumScale = chkMin
        .MaximumScale = chkMax
    End With
    
    With ActiveSheet.ChartObjects(2)
        .Top = 1 * (500 / windowSize) + 150
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
            .PlotArea.Width = (((550 * windowRatio) - 100) / windowSize)
            .ChartArea.Border.LineStyle = 0
        End With
    End With
    
    With ActiveSheet.ChartObjects(3)
        .Top = 2 * (500 / windowSize) + 150
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
            .PlotArea.Width = (((550 * windowRatio) - 100) / windowSize)
            .ChartArea.Border.LineStyle = 0
        End With
    End With
    
    Cells(10, 1).Interior.ColorIndex = 4
    Cells(10, 2).Interior.ColorIndex = 15
    Cells(10, 3).Interior.ColorIndex = 41
    Cells(10, 4).Interior.ColorIndex = 15
    Cells(10, 5).Interior.ColorIndex = 45
    Cells(10, 6).Interior.ColorIndex = 3
    Cells(10, 7).Interior.ColorIndex = 26
    If strList = "FL" Then
        Cells(10, 8).Interior.ColorIndex = 15
        Cells(10, 9).Interior.ColorIndex = 47
        Cells(10, 10).Interior.ColorIndex = 31
        Cells(10, 11).Interior.ColorIndex = 33
    Else
        Cells(10, 8).Interior.ColorIndex = 16
    End If
    
    [A2:A2].Interior.ColorIndex = 3
    [C2:C2].Interior.ColorIndex = 3
    [B2:B2].Interior.ColorIndex = 7
    [D2:E2].Interior.ColorIndex = 7
    [B7:B7].Interior.ColorIndex = 33
    [D7:D7].Interior.ColorIndex = 33
    [B8:B8].Interior.ColorIndex = 17
    [D8:D8].Interior.ColorIndex = 17
    [C7:C7].Interior.ColorIndex = 37
    [E7:H7].Interior.ColorIndex = 37
    [C8:C8].Interior.ColorIndex = 24
    [E8:H8].Interior.ColorIndex = 24
    If strList = "FL" Then
        [G2:G2].Interior.ColorIndex = 3
        [H2:I2].Interior.ColorIndex = 7
        [H7:H7].Interior.ColorIndex = 33
        [H8:H8].Interior.ColorIndex = 17
        [I7:L7].Interior.ColorIndex = 37
        [I8:L8].Interior.ColorIndex = 24
    ElseIf strList = "Is" Then
        [H7:I7].Interior.ColorIndex = 37
        [H8:I8].Interior.ColorIndex = 24
    End If
    
    Cells(1, 1).Select
skipchkphotobl:
    If scanNum = 1 Then
        Application.DisplayAlerts = False
        Worksheets(strSheetCheckName2).Delete
        Application.DisplayAlerts = True
    End If
    
    sheetGraph.Activate
    
    If strList = "Is" Then
        strSheetFitName = "Fit_" + mid$(strSheetDataName, 1, Len(strSheetDataName) - 3)
        If ExistSheet(strSheetFitName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetFitName).Delete
            Application.DisplayAlerts = True
        End If
        
        Call descriptHidden1
        Application.DisplayAlerts = False
        On Error GoTo ErrorSave1
        ActiveWorkbook.SaveCopyAs Filename:=ActiveWorkbook.Path & "\" & mid$(ActiveWorkbook.Name, 1, InStrRev(ActiveWorkbook.Name, ".") - 1) & "_Is.xlsx"
ErrorSave1:
        Application.DisplayAlerts = True
        strSheetGraphName = "Graph_" + strSheetDataName
        strSheetAnaName = "Exp_" + strSheetDataName
        Call ExportCmp
        sheetAna.Activate
        Cells(1, 1).Value = "PE/eV"
        Cells(1, 2).Value = sheetData.Name
        Call Convert2Txt
        Application.DisplayAlerts = False
        sheetAna.Delete
        Application.DisplayAlerts = True
        sheetGraph.Activate
        GoTo LoopIf
    ElseIf strList = "Ip" Then
        If strAES = "EY" Then
            strSheetFitName = "Fit_" + strSheetDataName
        Else
            strSheetFitName = "Fit_" + mid$(strSheetDataName, 1, Len(strSheetDataName) - 3)
        End If
        
        If ExistSheet(strSheetFitName) Then
            Application.DisplayAlerts = False
            Worksheets(strSheetFitName).Delete
            Application.DisplayAlerts = True
        End If
        
        Call descriptHidden1
        Application.DisplayAlerts = False
        On Error GoTo ErrorSave2
        ActiveWorkbook.SaveCopyAs Filename:=ActiveWorkbook.Path & "\" & mid$(ActiveWorkbook.Name, 1, InStrRev(ActiveWorkbook.Name, ".") - 1) & "_Ip.xlsx"
ErrorSave2:
        Application.DisplayAlerts = True
        sheetData.Visible = True
        sheetAvg.Activate
        Call Convert2Txt
        sheetGraph.Activate
        GoTo LoopIf
    End If
End Sub

Sub ScanRangeCheck()
    numscancheck = 0
    
    If Len(strscanNum) < 3 Then Exit Sub
    If mid$(strscanNum, 1, 1) = "[" And mid$(strscanNum, Len(strscanNum), 1) = "]" Then
        i = 1
        j = 0
        strTest = strscanNum
        strscanNum = mid$(strscanNum, 2, Len(strscanNum) - 2)
        For iRow = 1 To Len(strscanNum)
            strLabel = mid$(strscanNum, iRow, 1)
            If IsNumeric(strLabel) = False Then
                If strLabel = "," Or strLabel = "-" Then
                Else
                    numscancheck = 0
                    Exit Sub
                End If
            End If
        Next

        If InStr(1, strscanNum, ",", 1) > 0 Then
            A = Split(strscanNum, ",")
            For i = LBound(A) To UBound(A)
                If InStr(1, A(i), "-", 1) > 0 Then
                    C = Split(A(i), "-")
                    If CInt(C(0)) > CInt(C(1)) Then
                        q = -1
                    Else
                        q = 1
                    End If
                    For k = CInt(C(0)) To CInt(C(1)) Step q
                        If k > 0 And k <= scanNumR Then
                            ReDim Preserve ratio(j + 1)
                            ratio(j + 1) = k
                            Debug.Print ratio(j + 1), 1
                            j = j + 1
                        End If
                    Next
                Else
                    If CInt(A(i)) > 0 And CInt(A(i)) <= scanNumR Then
                        ReDim Preserve ratio(j + 1)
                        ratio(j + 1) = CInt(A(i))
                        Debug.Print ratio(j + 1), 2
                        j = j + 1
                    End If
                End If
            Next
        ElseIf InStr(1, strscanNum, "-", 1) > 0 Then
            C = Split(strscanNum, "-")
            If CInt(C(0)) > CInt(C(1)) Then
                q = -1
            Else
                q = 1
            End If
            For k = CInt(C(0)) To CInt(C(1)) Step q
                If k > 0 And k <= scanNumR Then
                    ReDim Preserve ratio(j + 1)
                    ratio(j + 1) = k
                    Debug.Print ratio(j + 1), 3
                    j = j + 1
                End If
            Next
        Else
            If CInt(strscanNum) > 0 And CInt(strscanNum) <= scanNumR Then
                ReDim Preserve ratio(j + 1)
                ratio(j + 1) = CInt(strscanNum)
                Debug.Print ratio(j + 1), 4
                j = j + 1
            End If
        End If
        numscancheck = UBound(ratio)
    Else
        numscancheck = -1
    End If
End Sub

Sub HigherOrderCheck()
    If Len(strhighpe) < 4 Then Exit Sub
    If mid$(strhighpe, 1, 1) = ";" And mid$(strhighpe, Len(strhighpe) - 2, 3) = " eV" Then
        i = 1
        j = 0
        strscanNum = mid$(strhighpe, 2, Len(strhighpe) - 4)

        For iRow = 1 To Len(strscanNum)
            strLabel = mid$(strscanNum, iRow, 1)
            If IsNumeric(strLabel) = False Then
                If strLabel = ";" Or strLabel = "." Then
                Else
                    Exit Sub
                End If
            End If
        Next

        If InStr(1, strscanNum, ";", 1) > 0 Then
            A = Split(strscanNum, ";")
            If UBound(A) > 8 Then Exit Sub  ' limit of higher order or ghost is 8
            For i = LBound(A) To UBound(A)
                If CSng(A(i)) > 0 Then
                        ReDim Preserve highpe(j + 1)
                        highpe(j + 1) = CSng(A(i))
                        Debug.Print highpe(j + 1), 1
                        j = j + 1
                End If
            Next
        Else
            If CSng(strscanNum) > 0 Then
                ReDim Preserve highpe(j + 1)
                highpe(j + 1) = CSng(strscanNum)
                Debug.Print highpe(j + 1), 2
                j = j + 1
            End If
        End If
    End If
End Sub

Sub KeBL()
    If graphexist = 0 Then
        If Cells(1, 2).Value = "AlKa" Then
            pe = 1486.6
            multi = 0.001
        ElseIf strTest = "KE/eV" Or strTest = "BE/eV" Then
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
                strTest = "AE/eV"
            End If
            multi = 1
        End If
        
        If strTest = "BE/eV" Then
            wf = 4
        Else
            wf = 4
        End If
        
        char = 0
        cae = 100
        off = 0
        ncomp = 0
    End If
    Set rng = [A:A]
    numData = Application.CountA(rng) - 1
    
    If numData > 1 And IsNumeric(Cells(2, 1).Value) Then
        Do While IsNumeric(Cells(numData + 1, 1).Value) = False
            numData = numData - 1
        Loop
        Do While IsNumeric(Cells(numData + 1, 2).Value) = False
            numData = numData - 1
        Loop
    Else
        Call GetOut
        Exit Sub
    End If
    
    startEk = Cells(2, 1).Value
    endEk = Cells(numData + 1, 1).Value
    stepEk = Cells(3, 1).Value - Cells(2, 1).Value
    scanNum = 1
    Set dataData = Range(Cells(2, 1), Cells(numData + 1, 2))
    Set dataKeData = Range(Cells(2, 1), Cells(numData + 1, 1))
    Set dataIntData = dataKeData.Offset(, 1)
    
    If strTest = "KE/eV" Or strTest = "AE/eV" Or strTest = "PE/eV" Or strTest = "ME/eV" Or strTest = "GE/eV" Then
        If startEk > endEk Then
            U = Range(Cells(2, 1), Cells(numData + 1, 3))
            
            If ExistSheet("Sort_" & strSheetDataName) Then
                Application.DisplayAlerts = False
                Worksheets("Sort_" & strSheetDataName).Delete
                Application.DisplayAlerts = True
            End If
    
            Worksheets.Add().Name = "Sort_" & strSheetDataName
            Range(Cells(2, 1), Cells(numData + 1, 3)) = U
            Range(Cells(2, 1), Cells(numData + 1, 3)).Sort Key1:=Cells(2, 1), Order1:=xlAscending
            Set dataData = Range(Cells(2, 1), Cells(numData + 1, 2))
            Set dataKeData = Range(Cells(2, 1), Cells(numData + 1, 1))
            Set dataIntData = dataKeData.Offset(, 1)
            startEk = Cells(2, 1).Value
            endEk = Cells(numData + 1, 1).Value
            stepEk = Cells(3, 1).Value - Cells(2, 1).Value
            Cells(1, 1).Value = strTest & "/sort"
            Cells(1, 2).Value = "Y/sort"
            Cells(1, 3).Value = "Ie/sort"
        End If
    ElseIf strTest = "BE/eV" Then
        If startEk < endEk Then
            U = Range(Cells(2, 1), Cells(numData + 1, 3))
            
            If ExistSheet("Sort_" & strSheetDataName) Then
                Application.DisplayAlerts = False
                Worksheets("Sort_" & strSheetDataName).Delete
                Application.DisplayAlerts = True
            End If
    
            Worksheets.Add().Name = "Sort_" & strSheetDataName
            Range(Cells(2, 1), Cells(numData + 1, 3)) = U
            Range(Cells(2, 1), Cells(numData + 1, 3)).Sort Key1:=Cells(2, 1), Order1:=xlDescending
            Set dataData = Range(Cells(2, 1), Cells(numData + 1, 2))
            Set dataKeData = Range(Cells(2, 1), Cells(numData + 1, 1))
            Set dataIntData = dataKeData.Offset(, 1)
            startEk = Cells(2, 1).Value
            endEk = Cells(numData + 1, 1).Value
            stepEk = Cells(3, 1).Value - Cells(2, 1).Value
            Cells(1, 1).Value = strTest & "/sort"
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
            
    If str1 = "Ke" Then
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

Sub EachComp()
    cae = 0
    For Each Target In OpenFileName
            
        If StrComp(Target, ActiveWorkbook.FullName, 1) = 0 Then
            cae = 1     ' in case the original file opens
            GoTo SkipOpen
        End If
        
        strTest = mid$(Target, InStrRev(Target, "\") + 1, Len(Target) - InStrRev(Target, "\"))
        
        If Not WorkbookOpen(strTest) Then
            Workbooks.Open Target
            Workbooks(strTest).Activate

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
            'On Error GoTo SkipOpen
        End If
        
        If StrComp(mid$(strAna, 1, 3), "Fit", 1) = 0 Then    ' FitAnalysis, FitComp, FitRatioAnalysis
            If strAna = "FitRatioAnalysis" Then
                strCpa = "Ana_" + mid$(Target, InStrRev(Target, "\") + 1, Len(Target) - InStrRev(Target, "\") - 5)
            ElseIf mid$(strSheetFitName, 1, 9) = "Fit_Norm_" Then
                strCpa = "Fit_Norm_" + mid$(Target, InStrRev(Target, "\") + 1, Len(Target) - InStrRev(Target, "\") - 5)
            Else
                strCpa = "Fit_" + mid$(Target, InStrRev(Target, "\") + 1, Len(Target) - InStrRev(Target, "\") - 5)
            End If
        ElseIf mid$(strSheetGraphName, 1, 11) = "Graph_Norm_" Then
            strCpa = "Graph_Norm_" + mid$(Target, InStrRev(Target, "\") + 1, Len(Target) - InStrRev(Target, "\") - 5)    ' for Graph_Norm
            strAna = "Graph_Norm"
        Else
            strCpa = "Graph_" + mid$(Target, InStrRev(Target, "\") + 1, Len(Target) - InStrRev(Target, "\") - 5)    ' for .xlsx
        End If
        
        Target = mid$(Target, InStrRev(Target, "\") + 1, Len(Target) - InStrRev(Target, "\") - 5) + ".xlsx"
        
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
        
        If q = 1 Then
            If Not Cells(2, 1).Value = "PE shifts" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        ElseIf q = 2 Then
            If Not Cells(2, 1).Value = "PE" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        ElseIf q = 3 Then
            If Not Cells(2, 1).Value = "KE shifts" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        ElseIf q = 4 Then
            If Not Cells(2, 1).Value = "Shifts" Then
                If j = 0 Then
                    Workbooks(Target).Close True
                Else
                    Workbooks(Target).Sheets(strLabel).Activate
                    j = 0
                End If
                GoTo SkipOpen
            End If
        End If
        
        If StrComp(mid$(strAna, 1, 3), "Fit", 1) = 0 Then
            If strAna = "FitRatioAnalysis" Then
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
        Else
            If StrComp(Cells(40, para + 9).Value, "Ver.", 1) = 0 Then
                iCol = para
            Else
                For iCol = 1 To 500
                    If StrComp(Cells(40, iCol + 9).Value, "Ver.", 1) = 0 Then
                        Exit For
                    ElseIf iCol = 500 Then
                        MsgBox "Graph sheet has no parameters to be compared."
                        End
                    End If
                Next
            End If
            
            If Cells(40, iCol + 10).Value >= 6.56 Then
                numData = Workbooks(Target).Sheets(strCpa).Cells(41, iCol + 12).Value
            Else
                MsgBox "Macro code used in some data comparison is obsolete!"
                If numData = 0 Then GoTo SkipOpen
            End If
        End If
        
        str4 = Cells(10, 1).Value       'check whether BE/eV or KE/eV. If BE/eV, only BE graph available
        
        If strAna = "FitAnalysis" Then
            C = Workbooks(Target).Sheets(strCpa).Range(Cells(1, 5), Cells(18 + sftfit2, 4 + g))
            b = Workbooks(Target).Sheets(strCpa).Range(Cells(1, 1), Cells(1, 3))
            
            For iCol = 0 To g - 1
                For iRow = 0 To g - 1
                    If D(3, iCol + 5) = C(1, iRow + 1) Then                                 ' Check Name of peak
                        D(5 + i, iCol + 5) = C(2, iRow + 1)                                 ' BE
                        If C(16 + sftfit2, iRow + 1) > 0 Then
                            D(5 + i + spacer + UBound(OpenFileName), iCol + 5) = C(16 + sftfit2, iRow + 1)      ' T.I.Area
                            D(5 + i + 2 * (spacer + UBound(OpenFileName)), iCol + 5) = C(17 + sftfit2, iRow + 1)  ' S.I.Area
                            D(5 + i + 3 * (spacer + UBound(OpenFileName)), iCol + 5) = C(18 + sftfit2, iRow + 1)    ' N.I.Area
                        Else
                            D(5 + i + spacer + UBound(OpenFileName), iCol + 5) = 0
                            D(5 + i + 2 * (spacer + UBound(OpenFileName)), iCol + 5) = 0  ' S.Area
                            D(5 + i + 3 * (spacer + UBound(OpenFileName)), iCol + 5) = 0    ' N.Area
                        End If
                        D(5 + i + 4 * (spacer + UBound(OpenFileName)), iCol + 5) = C(4, iRow + 1)     ' FWHM
                        Exit For
                    End If
                Next
            Next
    
            For Gnum = 0 To 4
                D(5 + (spacer + UBound(OpenFileName)) * Gnum + i, 1) = Target
                D(5 + (spacer + UBound(OpenFileName)) * Gnum + i, 2) = strCpa
                D(5 + (spacer + UBound(OpenFileName)) * Gnum + i, 4) = Workbooks(Target).Sheets(strCpa).Cells(8 + sftfit2, 2).Value
            Next
            
            For Gnum = 0 To 2
                D(5 + i, g + 6 + Gnum) = b(1, 1 + Gnum)
            Next
            
            If j = 0 Then
                Workbooks(Target).Close True
            Else
                Workbooks(Target).Sheets(strLabel).Activate
                j = 0
            End If
            
            i = i + 1
            GoTo SkipOpen
        ElseIf strAna = "FitRatioAnalysis" Then
            Dim spacera As Integer
            Dim ga As Integer
            Dim filenuma As Integer
            Dim iCola As Integer
            Dim iRowa As Integer
            
            spacera = Workbooks(Target).Sheets(strCpa).Cells(2, para + 1).Value     ' spacer
            ga = Workbooks(Target).Sheets(strCpa).Cells(3, para + 1).Value          ' # of peaks
            filenuma = Workbooks(Target).Sheets(strCpa).Cells(4, para + 1).Value    ' # of files
            C = Workbooks(Target).Sheets(strCpa).Range(Cells(1, 1), Cells((4 + spacera * 4) + 5 * filenuma, 9 + 2 * ga)) ' No check in matching among the peak names.
            b = Workbooks(Target).Sheets(strCpa).Range(Cells(4, 6 + ga), Cells(3 + filenuma, 8 + ga))
            
            D(1, g + 5) = Target
            D(2, g + 5) = strCpa
                
            For iCola = 0 To ga - 1
                For iRowa = 0 To fileNum   ' include the peak name
                    D(3 + iRowa, iCola + g + 5) = C(3 + iRowa, iCola + 5)                                 ' BE
                    D(2 + iRowa + 1 * (spacer + fileNum), iCola + g + 5) = C(2 + iRowa + 1 * (spacera + filenuma), iCola + 5)      ' P.Area
                    D(1 + iRowa + 2 * (spacer + fileNum), iCola + g + 5) = C(1 + iRowa + 2 * (spacera + filenuma), iCola + 5)  ' S.Area
                    D(0 + iRowa + 3 * (spacer + fileNum), iCola + g + 5) = C(0 + iRowa + 3 * (spacera + filenuma), iCola + 5)    ' N.Area
                    D(-1 + iRowa + 4 * (spacer + fileNum), iCola + g + 5) = C(-1 + iRowa + 4 * (spacera + filenuma), iCola + 5)     ' FWHM
                Next
            Next
            
            For Gnum = 0 To fileNum - 1
                A(Gnum + 1, i + 2) = b(1 + Gnum, 1) & b(1 + Gnum, 2) & b(1 + Gnum, 3)
            Next
            
            If j = 0 Then
                Workbooks(Target).Close True
            Else
                Workbooks(Target).Sheets(strLabel).Activate
                j = 0
            End If
            
            g = g + ga
            i = i + 1
            GoTo SkipOpen
        ElseIf strAna = "FitComp" Then
            numData = Cells(5, 101).Value
            tmp = Workbooks(Target).Sheets(strCpa).Range(Cells(20 + sftfit, 1), Cells(20 + sftfit + numData, 1)).Value
            en = Workbooks(Target).Sheets(strCpa).Range(Cells(20 + sftfit, 4), Cells(20 + sftfit + numData, 4)).Value

            sheetAna.Activate
            sheetAna.Range(Cells(10, (4 + (i * 3))), Cells(10 + numData, (4 + (i * 3)))).Value = tmp
            sheetAna.Range(Cells(10, (6 + (i * 3))), Cells(10 + numData, (6 + (i * 3)))).Value = en
            If StrComp(mid$(Cells(10, (4 + (i * 3))).Value, 1, 2), "BE", 1) = 0 Then
                str1 = "Be"
                str2 = "Sh"
                str3 = "In"
                Cells(4, (4 + (i * 3))) = "Shift"
                Cells(4, (5 + (i * 3))) = 0
                Cells(4, (6 + (i * 3))) = "eV"
                Cells(10, (5 + (i * 3))) = "Shift"
                Range(Cells(4, (4 + (i * 3))), Cells(4, (4 + (i * 3)))).Interior.ColorIndex = 3
                Range(Cells(4, (5 + (i * 3))), Cells(4, (6 + (i * 3)))).Interior.ColorIndex = 38
            ElseIf StrComp(mid$(Cells(10, (4 + (i * 3))).Value, 1, 2), "PE", 1) = 0 Then
                str1 = "Pe"
                str2 = "Sh"
                str3 = "Ab"
                Cells(2, (4 + (i * 3))).Value = "Shift"
                Cells(2, (5 + (i * 3))).Value = 0
                Cells(2, (6 + (i * 3))).Value = "eV"
                Cells(10, (5 + (i * 3))).Value = "Shift"
                Range(Cells(2, (4 + (i * 3))), Cells(2, (4 + (i * 3)))).Interior.ColorIndex = 3
                Range(Cells(2, (5 + (i * 3))), Cells(2, (6 + (i * 3)))).Interior.ColorIndex = 38
            ElseIf StrComp(mid$(Cells(10, (4 + (i * 3))).Value, 1, 2), "ME", 1) = 0 Then
                str1 = "Po"
                str2 = "Sh"
                str3 = "Ab"
                Cells(2, (4 + (i * 3))).Value = "Shift"
                Cells(2, (5 + (i * 3))).Value = 0
                Cells(2, (6 + (i * 3))).Value = "a.u."
                Cells(10, (5 + (i * 3))).Value = "Shift"
                Range(Cells(2, (4 + (i * 3))), Cells(2, (4 + (i * 3)))).Interior.ColorIndex = 3
                Range(Cells(2, (5 + (i * 3))), Cells(2, (6 + (i * 3)))).Interior.ColorIndex = 38
            End If
            strSheetGraphName = strSheetAnaName
            p = 2
        Else
            Workbooks(Target).Sheets(strCpa).Range(Cells(2, 1), Cells(10 + numData, 3)).Copy Destination:=Workbooks(wb).Sheets(strSheetGraphName).Cells(2, (4 + (i * 3)))
            Workbooks(wb).Sheets(strSheetGraphName).Activate
        End If
        
        strCasa = Cells(1, (5 + (i * 3))).Value
        Cells(1, (5 + (i * 3))).Value = Target
        Cells(9, (4 + (i * 3))).Value = "Offset/multp"
        Cells(9, (5 + (i * 3))).Value = 0

        If WorksheetFunction.Round(Cells(2, (5 + (i * 3))), 1) = 1486.6 And StrComp(mid$(strAna, 1, 3), "Fit", 1) <> 0 Then
            Cells(9, (6 + (i * 3))).Value = 0.001
        Else
            Cells(9, (6 + (i * 3))).Value = 1
        End If

        Cells(9, (4 + (i * 3))).Interior.Color = RGB(139, 195, 74)
        Range(Cells(9, (5 + (i * 3))), Cells(9, (6 + (i * 3)))).Interior.Color = RGB(197, 225, 165)
        
        If Cells(3, (4 + (i * 3))).Interior.ColorIndex = 45 Then
            Cells(3, (5 + (i * 3))).FormulaR1C1 = "=(Ln((((Sqrt((((950 *((" & gamma & ")^2))/(R4C * " & lambda & "))-1) * 2))/(" & lambda & " * 0.934)) - " & a0 & ")/(" & a1 & ")))/(" & a2 & ")" ' gap
        ElseIf Cells(4, (4 + (i * 3))).Interior.ColorIndex = 45 Then
            Cells(4, (5 + (i * 3))).FormulaR1C1 = "=950 * ((" & gamma & ") ^ 2) / (((((0.934 * " & lambda & " * (" & a0 & " + " & a1 & " * Exp(" & a2 & " * R3C))) ^ 2) / 2) + 1) * " & lambda & ")" ' 1st har.
        End If
        
        imax = numData + 10
        
        If str1 = "Ke" And str3 = "In" Then
            If str4 = "Be" Then
                Cells(11, (5 + (i * 3))).FormulaR1C1 = "=R4C + RC[-1]"
                Cells(10 + (imax), (5 + (i * 3))).FormulaR1C1 = "=R4C + R[-" & (imax - 1) & "]C[-1]"
            ElseIf str4 = "Ek" Then ' this is a trigger to handle "BE/eV" data
                Cells(11, (4 + (i * 3))).FormulaR1C1 = "=R2C[1] - RC[1] - R3C[1]"
                Cells(10 + (imax), (5 + (i * 3))).FormulaR1C1 = "=-R4C + R[-" & (imax - 1) & "]C"
            Else
                Cells(11, (5 + (i * 3))).FormulaR1C1 = "=R2C - R3C - R4C - RC[-1]"
                Cells(10 + (imax), (5 + (i * 3))).FormulaR1C1 = "=R2C - R3C - R4C - R[-" & (imax - 1) & "]C[-1]"
            End If
        ElseIf str1 = "Pe" Or str1 = "Po" Then
            Cells(11, (5 + (i * 3))).FormulaR1C1 = "=R2C + RC[-1]"
            Cells(10 + (imax), (5 + (i * 3))).FormulaR1C1 = "=R2C + R[-" & (imax - 1) & "]C[-1]"
        ElseIf str1 = "Be" Then
            If str4 = "Ke" Or str4 = "Ek" Then  ' old data used with "Ek"
                Cells(11, (5 + (i * 3))).FormulaR1C1 = "=R2C - R3C - R4C - RC[-1]"
                Cells(10 + (imax), (5 + (i * 3))).FormulaR1C1 = "=R2C - R3C - R4C - R[-" & (imax - 1) & "]C[-1]"
            Else
                Cells(11, (5 + (i * 3))).FormulaR1C1 = "=R4C + RC[-1]"
                Cells(10 + (imax), (5 + (i * 3))).FormulaR1C1 = "=R4C + R[-" & (imax - 1) & "]C[-1]"
            End If
        ElseIf str3 = "De" Then
            Cells(10 + (imax), (4 + (i * 3))).FormulaR1C1 = "=R2C2 + R[-" & (imax - 1) & "]C"
            Range(Cells(10 + (imax), (4 + (i * 3))), Cells((2 * imax) - 1, (4 + (i * 3)))).FillDown
            Cells(10 + (imax), (5 + (i * 3))).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C) *R9C[1]"
            Range(Cells(10 + (imax), (5 + (i * 3))), Cells((2 * imax) - 1, (5 + (i * 3)))).FillDown
            Cells(10 + (imax), (6 + (i * 3))).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C) *R9C"
            Range(Cells(10 + (imax), (6 + (i * 3))), Cells((2 * imax) - 1, (6 + (i * 3)))).FillDown
            GoTo AESmode
        End If
        
        If str1 = "Ke" And str3 = "In" And str4 = "Ek" Then
            Range(Cells(11, (4 + (i * 3))), Cells((imax), (4 + (i * 3)))).FillDown
        Else
            Range(Cells(11, (5 + (i * 3))), Cells((imax), (5 + (i * 3)))).FillDown
        End If
        Cells(10 + (imax), (4 + (i * 3))).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
        Range(Cells(10 + (imax), (4 + (i * 3))), Cells((2 * imax) - 1, (4 + (i * 3)))).FillDown
        Range(Cells(10 + (imax), (5 + (i * 3))), Cells((2 * imax) - 1, (5 + (i * 3)))).FillDown
        Cells(10 + (imax), (6 + (i * 3))).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C[-1])*R9C"
        Range(Cells(10 + (imax), (6 + (i * 3))), Cells((2 * imax) - 1, (6 + (i * 3)))).FillDown
AESmode:
        Range(Cells((2 * imax), (4 + (i * 3))), Cells((2 * imax), (6 + (i * 3))).End(xlDown)).Clear
        Range(Cells((imax + 1), (4 + (i * 3))), Cells((imax + 9), (6 + (i * 3)))).Clear
        Range(Cells(9, (4 + (i * 3))), Cells(9, ((4 + (i * 3))))).Interior.ColorIndex = 43
        Range(Cells(9, (5 + (i * 3))), Cells(9, ((6 + (i * 3))))).Interior.ColorIndex = 35
        
        Set dataKeGraph = Range(Cells(10 + (imax), (4 + (i * 3))), Cells((2 * imax - 1), (4 + (i * 3))))
        
        If j = 0 Then
            Workbooks(Target).Close True
        Else
            Workbooks(Target).Sheets(strLabel).Activate
            j = 0
        End If
        
        Workbooks(wb).Sheets(strSheetGraphName).Activate
        ActiveSheet.ChartObjects(1).Activate
        
        If i > ncomp - 1 Then
            ActiveChart.SeriesCollection.NewSeries
            Gnum = ActiveChart.SeriesCollection.Count
        Else
            Gnum = 1
            If Cells(42, para + 12).Value > 0 Then Gnum = Gnum + 1
            If Cells(43, para + 12).Value > 0 Then Gnum = Gnum + 1
            If Cells(44, para + 12).Value > 0 Then Gnum = Gnum + 1
            Gnum = Gnum + i + 1
            Debug.Print Gnum
        End If

        With ActiveChart.SeriesCollection(Gnum)
            .ChartType = xlXYScatterLinesNoMarkers
            If str3 = "De" Then
                .Name = "='" & ActiveSheet.Name & "'!R1C" & (5 + (i * 3)) & ""
                .XValues = dataKeGraph.Offset(0, 0)
                .Values = dataKeGraph.Offset(0, 1)
            Else
                .Name = "='" & ActiveSheet.Name & "'!R1C" & (5 + (i * 3)) & ""
                .XValues = dataKeGraph.Offset(0, 1)
                .Values = dataKeGraph.Offset(0, 2)
            End If
            SourceRangeColor1 = .Border.Color
        End With
        
        If str1 = "Ke" And (str4 = "Ke" Or str4 = "Ek") Then
            ActiveSheet.ChartObjects(2).Activate
            If i > ncomp - 1 Then
                ActiveChart.SeriesCollection.NewSeries
            End If
            With ActiveChart.SeriesCollection(Gnum)
                .ChartType = xlXYScatterLinesNoMarkers
                .Name = "='" & ActiveSheet.Name & "'!R1C" & (5 + (i * 3)) & ""
                .XValues = dataKeGraph
                .Values = dataKeGraph.Offset(0, 2)
                SourceRangeColor2 = .Border.Color
            End With
        
            Range(Cells(10, (4 + (i * 3))), Cells(10, ((4 + (i * 3))))).Interior.Color = SourceRangeColor2
            Range(Cells(9 + (imax), (4 + (i * 3))), Cells(9 + (imax), ((4 + (i * 3))))).Interior.Color = SourceRangeColor2
        Else
            Range(Cells(10, (4 + (i * 3))), Cells(10, ((4 + (i * 3))))).Interior.Color = SourceRangeColor1
            Range(Cells(9 + (imax), (4 + (i * 3))), Cells(9 + (imax), ((4 + (i * 3))))).Interior.Color = SourceRangeColor1
        End If

        Range(Cells(10, (5 + (i * 3))), Cells(10, ((5 + (i * 3))))).Interior.Color = SourceRangeColor1
        Range(Cells(9 + (imax), (5 + (i * 3))), Cells(9 + (imax), ((5 + (i * 3))))).Interior.Color = SourceRangeColor1
        strTest = mid$(Cells(1, (5 + (i * 3))).Value, 1, Len(Cells(1, (5 + (i * 3))).Value) - 5)
        Cells(8 + (imax), (5 + (i * 3))).Value = Cells(1, (5 + (i * 3))).Value
        Cells(9 + (imax), (4 + (i * 3))).Value = str1 + strTest
        Cells(9 + (imax), (5 + (i * 3))).Value = str2 + strTest
        Cells(9 + (imax), (6 + (i * 3))).Value = str3 + strTest
        i = i + 1
SkipOpen:
    Next Target
End Sub

Sub descriptGraph()
    Cells(2, 1).Value = "PE"
    Cells(3, 1).Value = "WF"
    Cells(4, 1).Value = "Char"
    Cells(5, 1).Value = "Start KE"
    Cells(6, 1).Value = "End KE"
    Cells(7, 1).Value = "Step KE"
    Cells(8, 1).Value = "# scan"
    If StrComp(Cells(2, 1).Value, "PE", 1) = 0 Then
        If UBound(highpe) > 0 Then
            Cells(2, 3).Value = strhighpe
            [C3:C7].Value = "eV"
        Else
            [C2:C7].Value = "eV"
        End If
    End If
    
    Cells(10, 1).Value = "Ke"
    Cells(10, 2).Value = "Be"
    Cells(10, 3).Value = "In"
    If g = 0 Then
        Cells(1, 2).Value = Gnum
    Else
        Cells(1, 2).Value = g
    End If
    Cells(2, 2).Value = pe
    Cells(3, 2).Value = wf
    Cells(4, 2).Value = char
    Cells(5, 2).Value = startEk
    Cells(6, 2).Value = endEk
    Cells(7, 2).Value = stepEk
    If numscancheck <= 0 Then
        Cells(8, 2).Value = scanNum
        Cells(8, 3).Value = "times"
        [B5:C8].Interior.Color = RGB(144, 202, 249)
    ElseIf numscancheck > 0 Then
        Cells(8, 2).Value = strscanNumR
        Cells(8, 3).Value = vbNullString
        [B5:C7].Interior.Color = RGB(144, 202, 249)
        [B8:C8].Interior.Color = RGB(255, 204, 128)
    End If
    Cells(9, 1).Value = "Offset/multp"
    Cells(9, 2).Value = off
    Cells(9, 3).Value = multi
    
    Call descriptHidden1
    [A2:A4].Interior.Color = RGB(244, 67, 54)
    [B2:C4].Interior.Color = RGB(244, 143, 177)
    [A5:A8].Interior.Color = RGB(3, 169, 244)
    Range(Cells(9, 1), Cells(9, 1)).Interior.Color = RGB(139, 195, 74)
    Range(Cells(9, 2), Cells(9, 3)).Interior.Color = RGB(197, 225, 165)
    imax = numData + 10
    
    If strTest = "PE/eV" Or strTest = "GE/eV" Then
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
        Cells(10 + (imax), 2).FormulaR1C1 = "=R2C2 + R[-" & (imax - 1) & "]C[-1]"
        strLabel = "Photon energy (eV)"
        str1 = "Pe"
        str2 = "Sh"
        str3 = "Ab"
    ElseIf strTest = "ME/eV" Then
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
        Cells(10 + (imax), 2).FormulaR1C1 = "=R2C2 + R[-" & (imax - 1) & "]C[-1]"
        strLabel = "Position (arb. unit)"
        str1 = "Po"
        str2 = "Sh"
        str3 = "Ab"
    ElseIf strTest = "BE/eV" Then
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
        strLabel = "Binding energy (eV)"
        str1 = "Ke"
        str2 = "Be"
        str3 = "In"
    ElseIf strTest = "AE/eV" Then
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
        strLabel = "Kinetic energy (eV)"
        str1 = "Ke"
        str2 = "Ae"
        str3 = "De"
    Else
        Cells(11, 2).FormulaR1C1 = "=R2C2 - R3C2 - R4C2 - RC[-1]"
        Cells(10 + (imax), 2).FormulaR1C1 = "=R2C - R3C - R4C - R[-" & (imax - 1) & "]C[-1]"
        strLabel = "Binding energy (eV)"
        str1 = "Ke"
        str2 = "Be"
        str3 = "In"
    End If
    
    If str3 = "De" Then
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
        Set dataIntGraph = dataKeGraph.Offset(, 2)
    Else
        If strTest = "BE/eV" Then
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
        Set dataIntGraph = dataKeGraph.Offset(, 2)
        If strTest = "BE/eV" Then
            startEk = Cells(11, 1).Value
            endEk = Cells(10 + numData, 1).Value
        End If
    End If
End Sub

Sub Gcheck()
    If IsNumeric(Cells(1, 2).Value) = False Then
        strAna = "ana"
    ElseIf Cells(1, 2).Value = 600 Then
        Gnum = 1
        g = 600
    ElseIf Cells(1, 2).Value = 1200 Then
        Gnum = 2
        g = 1200
    ElseIf Cells(1, 2).Value = 2400 Then
        Gnum = 3
        g = 2400
    ElseIf IsNumeric(Cells(1, 2).Value) = True Then
        Gnum = 0
        g = Cells(1, 2).Value
    Else
        Gnum = 0
        g = 0
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
    
    If g = 600 Then
        Gnum = 1
    ElseIf g = 1200 Then
        Gnum = 2
    ElseIf g = 2400 Then
        Gnum = 3
    Else
        Gnum = 0
    End If
    
    Cells(45, para + 12).Value = Gnum
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
    Dim tfa As Single
    Dim tfb As Single
    
    Cells(19, 101).Value = ver
    Cells(1, 1).Value = "Shirley"
    Cells(1, 2).Value = "BG"
    Cells(1, 3).Value = vbNullString
    Cells(2, 1).Value = "Tolerance"
    Cells(3, 1).Value = "Initial A"
    Cells(4, 1).Value = "Final A"
    Cells(5, 1).Value = "Iteration"
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
    
    If str1 = "Pe" Then
        Cells(20 + sftfit, 1).Value = "PE / eV"
        Cells(20 + sftfit, 2).Value = "Ab"
    ElseIf str1 = "Po" Then
        Cells(20 + sftfit, 1).Value = "ME / eV"
        Cells(20 + sftfit, 2).Value = "Ab"
    Else
        Cells(20 + sftfit, 1).Value = "BE / eV"
        Cells(20 + sftfit, 2).Value = "In"
    End If
    Cells(2, 2).Value = 0.000001
    Cells(3, 2).Value = 0.001
    Cells(15 + sftfit2, 2).Value = Gnum     ' Grating number, 0 means VersaProbe II
    If Cells(15 + sftfit2, 2).Value = 0 Then    ' VersaProbe II AlKa
        Cells(14 + sftfit2, 2).Value = 23.5
    Else
        Cells(14 + sftfit2, 2).Value = cae
    End If
                                            ' Inelastic mean free path parameter:
    Cells(16 + sftfit2, 2).Value = mfp     ' lambda is proportional to E^x, and x can be from 0.5 to 0.9.
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
    Cells(16, 100).Value = "Shirley"
    Cells(17, 100).Value = "Iteration limit"
    Cells(18, 100).Value = "Average data"
    Cells(19, 100).Value = "Ver."
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
    Cells(15, 101).Value = nomfac
    Cells(16, 101).Value = "BG"
    Cells(17, 101).Value = 10       ' limit of iteration
    Cells(18, 101).FormulaR1C1 = "=Average(R31C2:R" & (30 + numData) & "C2)"
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
    
    If Cells(15 + sftfit2, 2).Value = 1 Then   ' grating #1
        Cells(2, 103).Value = 2       ' max FWHM1 limit
        Cells(3, 103).Value = 0.1       ' min FWHM1 limit
        Cells(4, 103).Value = 2       ' max FWHM2 limit
        Cells(5, 103).Value = 0.1       ' min FWHM2 limit
        Cells(6, 103).Value = 0.999       ' max shape limit
        Cells(7, 103).Value = 0.001       ' min shape limit
        Cells(10, 101).Value = 20          ' average points for poly BG
        If str1 = "Pe" Then             ' additional BE step
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
        Cells(10, 101).Value = 10          ' average points for poly BG
        If str1 = "Pe" Then             ' additional BE step
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
    
    ' T integrated in BG-IN
    Cells(21 + sftfit + numData, 4) = IntegrationTrapezoid(Range(Cells(21 + sftfit, 1), Cells(20 + sftfit + numData, 1)), Range(Cells(21 + sftfit, 4), Cells(20 + sftfit + numData, 4)))
    ' T integrated in each peak
    For q = 1 To j
        Cells(16 + sftfit2, 4 + q) = IntegrationTrapezoid(Range(Cells(21 + sftfit, 1), Cells(20 + sftfit + numData, 1)), Range(Cells(21 + sftfit, 4 + q), Cells(20 + sftfit + numData, 4 + q)))
        Cells(17 + sftfit2, 4 + q).FormulaR1C1 = "= R" & (16 + sftfit2) & "C / (R" & (9 + sftfit2) & "C)"
        If str1 = "Pe" Or str1 = "Po" Then
        Else
            Cells(19 + sftfit2, 4 + q).FormulaR1C1 = "= (R15C101 * (1 - (0.25 * R" & (7 + sftfit2) & "C)*(3 * (cos(3.14*R24C2/180))^2 - 1)) * R" & (9 + sftfit2) & "C * ((R3C)^(R" & (16 + sftfit2) & "C2)) * R" & (14 + sftfit2) & "C2 * (((R" & (17 + sftfit2) & "C2^2)/((R" & (17 + sftfit2) & "C2^2)+((R3C)/(R" & (14 + sftfit2) & "C2))^2))^R" & (18 + sftfit2) & "C2))"
            Cells(18 + sftfit2, 4 + q).FormulaR1C1 = "= R" & (16 + sftfit2) & "C / R" & (19 + sftfit2) & "C"
        End If
    Next
    ' T integrated in sum peaks
    Cells(21 + sftfit + numData, 5 + j) = IntegrationTrapezoid(Range(Cells(21 + sftfit, 1), Cells(20 + sftfit + numData, 1)), Range(Cells(21 + sftfit, 5 + j), Cells(20 + sftfit + numData, 5 + j)))
    Range(Cells(11, 104), Cells(16, 104)).Delete
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
    Cells(16, 100).Value = "Shirley"
    Cells(16, 101).Value = "BG"
    
    If Cells(8, 101).Value = 0 Then 'Or Cells(9, 101).Value > 0 Then
        Cells(2, 2).Value = 0.000001
        If Cells(3, 2).Value > 0.1 Or Cells(3, 2).Value <= 0.0000001 Then Cells(3, 2).Value = 0.001
    ElseIf Cells(3, 2).Value >= 0.1 Or Cells(3, 2).Value <= 0.000001 Then
        Cells(3, 2).Value = 0.001
    End If
ShirleyBGagain:
    Cells(4, 2).Value = Cells(3, 2).Value
    Cells(startR, 98).FormulaR1C1 = "= (2 * RC1 - (R" & startR & "C1 + R" & endR & "C1))/(R" & endR & "C1 - R" & startR & "C1)" ' CT
    Range(Cells(startR, 98), Cells(endR, 98)).FillDown

    If Cells(20 + sftfit, 2).Value = "Ab" Then ' for PE
        If Cells(startR, 1).Value = Cells(6, 101).Value Then
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
        If Cells(endR, 1).Value = Cells(7, 101).Value Then
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
        For k = endR - 1 To startR Step -1
            Cells(k, 99).FormulaR1C1 = "= ABS(R[1]C2 - R[1]C3)"
            Cells(k, 3).FormulaR1C1 = "=R" & (endR + 1) & "C + R4C2 * SUM(R[1]C99:R" & (endR) & "C99)"
        Next
    End If

    Cells(startR, 100).FormulaR1C1 = "=((RC2 - RC3)^2)/(abs(RC3))" ' CV
    Range(Cells(startR, 100), Cells(endR, 100)).FillDown
    Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Cells(4, 2)
    SolverAdd CellRef:=Cells(4, 2), Relation:=1, FormulaText:=1  ' max
    SolverAdd CellRef:=Cells(4, 2), Relation:=3, FormulaText:=-1  ' min
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
    Cells(8, 1).Value = "Pre-edge"
    Cells(9, 1).Value = "Post-edge"
    Cells(16, 100).Value = "Victoreen"
    Cells(16, 101).Value = "BG"
    
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
    
    Cells(startR, 98).FormulaR1C1 = "= RC1 - R9C2"
    Range(Cells(startR, 98), Cells(endR, 98)).FillDown
    Cells(startR, 99).FormulaR1C1 = "= (2 * (RC1-R9C2) - (R" & startR & "C1 + R" & endR & "C1 -2*R9C2))/(R" & endR & "C1 - R" & startR & "C1)" ' PE
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
        Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    End If
        
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(6, 2))
    SolverAdd CellRef:=Range(Cells(3, 2), Cells(4, 2)), Relation:=3, FormulaText:=0
    SolverAdd CellRef:=Cells(3, 2), Relation:=1, FormulaText:=1 ' max
    ' 2nd poly should be zero if ratio is too small
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
    [A8:A9].Interior.Color = RGB(159, 168, 218)
    [B8:B9].Interior.Color = RGB(197, 202, 233)
End Sub

Sub TangentArcBG()
    Cells(1, 1).Value = "Arctan"
    Cells(1, 2).Value = "BG"
    Cells(1, 3).Value = vbNullString
    Cells(16, 100).Value = "Arctan"
    Cells(16, 101).Value = "BG"
    
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
    
    Cells(startR, 3).FormulaR1C1 = "=R2C2 + (1-R7C2) * ((R6C2 * (RC1 - R4C2))) + R7C2 * (R3C2 * ((0.5) + (1/3.14) * ATAN((RC1 - R4C2)/(R5C2 / 2))))"
    Range(Cells(startR, 3), Cells(endR, 3)).FillDown
    Cells(startR, 100).FormulaR1C1 = "=((RC2 - RC3)^2)/(abs(RC3))" ' CV
    Range(Cells(startR, 100), Cells(endR, 100)).FillDown
    Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    
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
    If StrComp(str1, "Po", 1) = 0 Then
    Else
        For k = 2 To 5
            If Cells(k, 2).Font.Bold = "True" Then
            ElseIf Cells(8, 101).Value = 0 Then
                If k = 2 Then
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
    Cells(1, 2).Value = "BG"
    Cells(1, 3).Value = vbNullString
    Cells(2, 1).Value = "a0"
    Cells(3, 1).Value = "a1"
    Cells(4, 1).Value = "a2"
    Cells(5, 1).Value = "a3"
    Cells(16, 100).Value = "Polynominal"
    Cells(16, 101).Value = "BG"
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
    Cells(6 + sftfit2, 2).FormulaR1C1 = "=(AVERAGE(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + AVERAGE(R" & endR & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(5, 2))
    
    For k = 2 To 5
        If Cells(k, 2).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
        End If
    Next
    
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
End Sub

Sub TougaardBG()    ' This is numerical convoluted Tougaard BG based on TougaardBG2 in extended version.
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
    Cells(16, 100).Value = "Tougaard"
    Cells(16, 101).Value = "BG"
    
    For k = 2 To 6
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
        End If
    Next
    
    Call descriptTConv
    Cells(startR, 100).FormulaR1C1 = "=((RC2 - RC3)^2)/(abs(RC3))" ' CV
    Range(Cells(startR, 100), Cells(endR, 100)).FillDown
    Cells(6 + sftfit2, 2).FormulaR1C1 = "= (Average(R" & startR & "C100:R" & (startR + ns - 1) & "C100) + Average(R" & endR - 1 & "C100:R" & (endR - ns + 1) & "C100)) / 2"
    SolverOk SetCell:=Cells(6 + sftfit2, 2), MaxMinVal:=2, ValueOf:="0", ByChange:=Range(Cells(2, 2), Cells(6, 2))
    SolverAdd CellRef:=Range(Cells(2, 2), Cells(5, 2)), Relation:=1, FormulaText:=5000
    SolverAdd CellRef:=Range(Cells(2, 2), Cells(5, 2)), Relation:=3, FormulaText:=0
    
    For k = 2 To 6
        If Cells(k, 2).Font.Bold = "True" Then
            SolverAdd CellRef:=Cells(k, 2), Relation:=2, FormulaText:=Cells(k, 2)
        End If
    Next
    
    SolverSolve UserFinish:=True
    SolverFinish KeepFinal:=1
    [A2:A7].Interior.Color = RGB(156, 204, 101)    '43
    [B2:B7].Interior.Color = RGB(197, 225, 165)    '35
End Sub

Sub descriptEFfit1()
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
    Cells(8 + sftfit2, 2).Value = -0.5
    Cells(9 + sftfit2, 2).Value = 0.5
    Cells(2, 2).Value = dblMax
    Cells(3, 2).Value = 1
    Cells(4, 2).Value = dblMin
    Cells(5, 2).Value = 0
    Cells(2, 5).Value = 0
    Cells(4, 5).Value = 300
    Cells(6, 2).Value = 0
    Cells(7, 2).Value = 0
    Cells(8, 2).Value = 0
    Cells(6, 5).Value = 0.1
    Cells(8, 5).Value = 1
    Cells(20 + sftfit, 3).Value = "FitEF (FD)"
    Cells(20 + sftfit, 4).Value = "Least fits (FD)"
    Cells(20 + sftfit, 5).Value = "Residual (FD)"
    Cells(20 + sftfit, 6).Value = "FitEF (GC)"
    Cells(20 + sftfit, 7).Value = "Least fits (GC)"
    Cells(20 + sftfit, 8).Value = "Residual (GC)"
    Cells(8, 101).Value = 0     ' 7.45: revised from "-1"
    Cells(16, 100).Value = "EF"
    Cells(16, 101).Value = "fit"
End Sub

Sub descriptEFfit2()
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
End Sub

Sub ProfileAnalyzer()
    Dim a0 As Integer, a1, a2, a3, a4, a5
    If IsNumeric(Cells(7, (4 + i)).Value) Then
        If Cells(7, (4 + i)).Value = 0 Then
            a1 = 0
        ElseIf Cells(7, (4 + i)).Value = 1 Then
            a1 = 1
        Else
            a1 = 2
        End If
    Else
        If Cells(7, (4 + i)).Value = "Gauss" Then
            a1 = 0
        ElseIf Cells(7, (4 + i)).Value = "Lorentz" Then
            a1 = 1
        Else
            a1 = 2
        End If
    End If
    
    If Cells(7, (4 + i)).Font.Italic Then
        a2 = 1
    Else
        a2 = 0
    End If
    
    If Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleNone Then
        a3 = 0
    ElseIf Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleSingle Then
        a3 = 1
    ElseIf Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleDouble Then
        a3 = 2
    End If
    
    If LCase(Cells(9, 103).Value) = "multipak" Then
        a4 = 1
        a5 = 1
    ElseIf LCase(Cells(9, 103).Value) = "product" Then
        a4 = -1
        a5 = 0
    Else
        a4 = 1      ' "sum"
        a5 = 0
    End If
    
    a0 = (1000 * a5 + 100 * a1 + 10 * a2 + a3) * a4
    
    If Cells(11, (4 + i)).Value = "G" And a0 <> 0 Then
        Cells(7, (4 + i)).Value = "Gauss"
        Cells(7, (4 + i)).Font.Italic = "False"
        Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleNone
    ElseIf Cells(11, (4 + i)).Value = "L" And Abs(a0) <> 100 Then
        Cells(7, (4 + i)).Value = "Lorentz"
        Cells(7, (4 + i)).Font.Italic = "False"
        Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleNone
    ElseIf Cells(11, (4 + i)).Value = "DS L" And Abs(a0) <> 110 Then
        Cells(7, (4 + i)).Value = "Lorentz"
        Cells(7, (4 + i)).Font.Italic = "True"
        Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleNone
    ElseIf Cells(11, (4 + i)).Value = "SGL" And a0 <> 200 Then
        If Not 0 < Cells(7, (4 + i)).Value < 1 Or IsNumeric(Cells(7, (4 + i)).Value) = False Then Cells(7, (4 + i)).Value = 0.2
        Cells(7, (4 + i)).Font.Italic = "False"
        If a0 > 1000 Then
            Cells(9, 103).Value = "MultiPak"
        Else
            Cells(9, 103).Value = "Sum"
        End If
        Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleNone
    ElseIf Cells(11, (4 + i)).Value = "PGL" And a0 <> -200 Then
        If Not 0 < Cells(7, (4 + i)).Value < 1 Or IsNumeric(Cells(7, (4 + i)).Value) = False Then Cells(7, (4 + i)).Value = 0.2
        Cells(7, (4 + i)).Font.Italic = "False"
        Cells(9, 103).Value = "Product"
        Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleNone
    ElseIf Cells(11, (4 + i)).Value = "TSGL" And a0 <> 1201 Then
        If Not 0 < Cells(7, (4 + i)).Value < 1 Or IsNumeric(Cells(7, (4 + i)).Value) = False Then Cells(7, (4 + i)).Value = 0.2
        Cells(7, (4 + i)).Font.Italic = "False"
        Cells(9, 103).Value = "MultiPak"
        Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleSingle
    ElseIf Cells(11, (4 + i)).Value = "GL" And a0 <> 1200 Then
        If Not 0 < Cells(7, (4 + i)).Value < 1 Or IsNumeric(Cells(7, (4 + i)).Value) = False Then Cells(7, (4 + i)).Value = 0.2
        Cells(7, (4 + i)).Font.Italic = "False"
        Cells(9, 103).Value = "MultiPak"
        Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleNone
    End If
End Sub

Sub FitEquations()
    If Cells(15 + sftfit2, 2).Value = 1 Then
        Cells(15, 101).Value = 0.01
    ElseIf Cells(15 + sftfit2, 2).Value = 2 Then
        Cells(15, 101).Value = 0.0002
    ElseIf Cells(15 + sftfit2, 2).Value = 3 Then
        Cells(15, 101).Value = 0.0001
    Else
        Cells(15, 101).Value = 0.001    ' VersaProbe II AlKa
    End If
    
    imax = 0    '# of iteration for asymmetric voigt fit
    j = Cells(8 + sftfit2, 2).Value
    q = Cells(9, 101).Value
    Range(Cells(1, (5 + j)), Cells(15 + sftfit2 + 4, 55)).Clear
    Range(Cells(20 + sftfit, 5), Cells((2 * numData + 22 + sftfit), 55)).ClearContents
    Range(Cells(1, 5), Cells(15 + sftfit2, (4 + j))).Interior.Color = RGB(178, 235, 242) '34
    Range(Cells(15 + sftfit2 + 1, 5), Cells(15 + sftfit2 + 4, (4 + j))).Interior.Color = RGB(207, 216, 220)
    
    If q < j Then
        If (j - q) Mod 2 = 0 And j Mod 2 = 0 Then
            For i = 1 To (j - q) Step 2
                Range(Cells(1, (4 + q + i)), Cells(9 + sftfit2, (4 + q + i + 1))).Value = Range(Cells(1, (4 + q - 1)), Cells(9 + sftfit2, (4 + q))).Value
                Range(Cells(14 + sftfit2, (4 + q + i)), Cells(15 + sftfit2, (4 + q + i + 1))).Value = Range(Cells(14 + sftfit2, (4 + q - 1)), Cells(15 + sftfit2, (4 + q))).Value
                Cells(1, (4 + q + i)).Value = Cells(1, 5).Value + "_" + CStr((4 + q + i - 5) / 2)
                Cells(1, (4 + q + i + 1)).Value = Cells(1, 6).Value + "_" + CStr((4 + q + i + 1 - 6) / 2)
                Cells(2, (4 + q + i)).Value = Cells(2, (4 + q - 1)).Value + i * (Cells(8, 103).Value / Cells(8 + sftfit2, 2).Value)
                Cells(2, (4 + q + i + 1)).Value = Cells(2, (4 + q)).Value + i * (Cells(8, 103).Value / Cells(8 + sftfit2, 2).Value)
                If Cells(4, 5).Font.Bold = True Then
                    Cells(4, (4 + q + i)).Font.Bold = True
                End If
                If Cells(4, 6).Font.Bold = True Then
                    Cells(4, (4 + q + i + 1)).Font.Bold = True
                End If
            Next
        Else
            For i = 1 To (j - q)
                Range(Cells(1, (4 + q + i)), Cells(9 + sftfit2, (4 + q + i))).Value = Range(Cells(1, (4 + q)), Cells(9 + sftfit2, (4 + q))).Value
                Cells(1, (4 + q + i)).Value = Cells(1, 5).Value + "s" + CStr((4 + q + i)-5)
                Cells(2, (4 + q + i)).Value = Cells(2, (4 + q)).Value + i * (Cells(8, 103).Value / Cells(8 + sftfit2, 2).Value)
            Next
        End If
        Cells(9, 101).Value = j
    ElseIf q > j Then
        Cells(9, 101).Value = j
    End If

    For i = 1 To j
        Call ProfileAnalyzer
        If IsEmpty(Cells(7, (4 + i))) = True Then Cells(7, (4 + i)) = 0
        If IsNumeric(Cells(7, (4 + i))) = False Then
            If Cells(7, (4 + i)) = "Gauss" Then
                Cells(7, (4 + i)) = 0
            ElseIf Cells(7, (4 + i)) = "Lorentz" Then
                Cells(7, (4 + i)) = 1
            ElseIf Cells(7, (4 + i)) = "Voigt" Then
                Cells(7, (4 + i)) = 0.5
            Else
                Cells(7, (4 + i)) = 0
            End If
        Else
            If Cells(7, (4 + i)) < 0 Or Cells(7, (4 + i)) > 1 Then Cells(7, (4 + i)) = 0
        End If
        If Cells(7, (4 + i)) = 0 Then
            Cells(startR, (4 + i)).FormulaR1C1 = "=R6C * EXP(-(1/2)*((RC[" & (-3 - i) & "]-R2C)/(R4C/2.35))^2)"
            Range(Cells(startR, (4 + i)), Cells(endR, (4 + i))).FillDown
            Cells(10 + sftfit2, (4 + i)).FormulaR1C1 = "=SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)"  ' Area Gauss
            Cells(11 + sftfit2, (4 + i)).FormulaR1C1 = "=SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14) / R14C" ' Area Gauss            Cells(12 + sftfit2, (4 + i)).FormulaR1C1 = "=SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14) / R24C" ' Area Gauss
            Cells(11, (4 + i)).Value = "G"
        ElseIf Cells(7, (4 + i)) = 1 Then
            If Cells(7, (4 + i)).Font.Italic = "True" Then  'Doniach-Sunjic (DS) convoluted with Lorentzian.
                For k = 1 To numData
                    If Cells((startR - 1 + k), 1).Value <= Cells(2, (4 + i)).Value Then
                        Cells((startR - 1 + k), (4 + i)).FormulaR1C1 = "= R6C * (((R4C/2)^2)/((RC[" & (-3 - i) & "]-R2C)^2 + (R4C/2)^2))"
                    Else            ' R8C used as an asymmetric parameter, R9C is second normalizing factor to connect to the other half.
                        Cells((startR - 1 + k), (4 + i)).FormulaR1C1 = "= R9C * R6C * Exp(Gammaln(1 - R8C)) * Cos((3.14 * R8C/2)+(1-R8C) * atan((RC[" & (-3 - i) & "]-R2C)/R4C)) / (((RC[" & (-3 - i) & "]-R2C)^2 + (R4C/2)^2)^((1-R8C)/2))"
                    End If
                Next
                
                Cells(10 + sftfit2, (4 + i)).FormulaR1C1 = "=(R6C * (R4C/2) * 3.14)" ' Area Lorentz
                Cells(11 + sftfit2, (4 + i)).FormulaR1C1 = "=(R6C * (R4C/2) * 3.14) / R14C"  ' Area Lorentz
                Cells(12 + sftfit2, (4 + i)).FormulaR1C1 = "=(R6C * (R4C/2) * 3.14) / R24C"  ' Area Lorentz
                Cells(11, (4 + i)).Value = "DS L"
            Else
                Cells(startR, (4 + i)).FormulaR1C1 = "= R6C * (((R4C/2)^2)/((RC[" & (-3 - i) & "]-R2C)^2 + (R4C/2)^2))"
                Range(Cells(startR, (4 + i)), Cells(endR, (4 + i))).FillDown
                Cells(10 + sftfit2, (4 + i)).FormulaR1C1 = "=(R6C * (R4C/2) * 3.14)" ' Area Lorentz
                Cells(11 + sftfit2, (4 + i)).FormulaR1C1 = "=(R6C * (R4C/2) * 3.14) / R14C"  ' Area Lorentz
                Cells(12 + sftfit2, (4 + i)).FormulaR1C1 = "=(R6C * (R4C/2) * 3.14) / R24C"  ' Area Lorentz
                Cells(11, (4 + i)).Value = "L"
            End If
        ElseIf 0 < Cells(7, (4 + i)).Value < 1 And Cells(9, 103).Value = "Sum" Then    ' GL sum form: SGL
            Cells(5, (4 + i)).Value = Cells(4, (4 + i)).Value
            If Cells(7, (4 + i)).Font.Italic = "True" And Cells(7, (4 + i)).Font.Underline <> xlUnderlineStyleSingle And Cells(7, (4 + i)).Font.Underline <> xlUnderlineStyleDouble Then          ' asymmetric GL sum function
                For k = 1 To numData
                    If Cells((startR - 1 + k), 1).Value < Cells(2, (4 + i)).Value Then
                        Cells((startR - 1 + k), (4 + i)).FormulaR1C1 = "=R6C * ((R7C)*((((R5C)/2)^2)/((RC[" & (-3 - i) & "]-R2C)^2 + ((R5C)/2)^2)) + (1- R7C)*(EXP(-(1/2)*((RC[" & (-3 - i) & "]-R2C)/(R5C/2.35))^2)))"
                    Else
                        Cells((startR - 1 + k), (4 + i)).FormulaR1C1 = "=R6C * ((R7C)*((((R4C)/2)^2)/((RC[" & (-3 - i) & "]-R2C)^2 + ((R4C)/2)^2)) + (1- R7C)*(EXP(-(1/2)*((RC[" & (-3 - i) & "]-R2C)/(R4C/2.35))^2)))"
                    End If
                Next

                Cells(10 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R4C/2) * 3.14)))"
                Cells(11 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R4C/2) * 3.14))) / R14C"
                Cells(12 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R4C/2) * 3.14))) / R24C "
                Cells(11, (4 + i)).Value = "ASGL"
            Else
                Debug.Print "else"
                Cells(startR, (4 + i)).FormulaR1C1 = "=R6C * ((R7C)*((((R5C)/2)^2)/((RC[" & (-3 - i) & "]-R2C)^2 + ((R5C)/2)^2)) + (1- R7C)*(EXP(-(1/2)*((RC[" & (-3 - i) & "]-R2C)/(R4C/2.35))^2)))"
                Range(Cells(startR, (4 + i)), Cells(endR, (4 + i))).FillDown
                Cells(10 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R5C/2) * 3.14))) "
                Cells(11 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R5C/2) * 3.14))) / R14C"
                Cells(12 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R5C/2) * 3.14))) / R24C"
                Cells(11, (4 + i)).Value = "SGL"
            End If
        ElseIf 0 < Cells(7, (4 + i)).Value < 1 And Cells(9, 103).Value = "Product" Then    ' GL product form: PGL
            Cells(5, (4 + i)).Value = Cells(4, (4 + i)).Value
            If Cells(7, (4 + i)).Font.Italic = "True" And Cells(7, (4 + i)).Font.Underline <> xlUnderlineStyleSingle And Cells(7, (4 + i)).Font.Underline <> xlUnderlineStyleDouble Then          ' asymmetric GL sum function
                For k = 1 To numData
                    If Cells((startR - 1 + k), 1).Value < Cells(2, (4 + i)).Value Then
                        Cells((startR - 1 + k), (4 + i)).FormulaR1C1 = "=R6C * ((EXP(-(1/2)*(1- R7C)*((RC[" & (-3 - i) & "]-R2C)/(R4C/2.35))^2))/((((R4C)/2)^2)/((R7C)*(RC[" & (-3 - i) & "]-R2C)^2 + ((R4C)/2)^2)))"
                    Else
                        Cells((startR - 1 + k), (4 + i)).FormulaR1C1 = "=R6C * ((EXP(-(1/2)*(1- R7C)*((RC[" & (-3 - i) & "]-R2C)/(R5C/2.35))^2))/((((R5C)/2)^2)/((R7C)*(RC[" & (-3 - i) & "]-R2C)^2 + ((R5C)/2)^2)))"
                    End If
                Next

                Cells(10 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R4C/2) * 3.14)))"
                Cells(11 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R4C/2) * 3.14))) / R14C"
                Cells(12 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R4C/2) * 3.14))) / R24C"
                Cells(11, (4 + i)).Value = "APGL"
            Else
                Debug.Print "else"
                Cells(startR, (4 + i)).FormulaR1C1 = "=R6C * ((EXP(-(1/2)*(1- R7C)*((RC[" & (-3 - i) & "]-R2C)/(R4C/2.35))^2))/((((R5C)/2)^2)/((R7C)*(RC[" & (-3 - i) & "]-R2C)^2 + ((R5C)/2)^2)))"
                Range(Cells(startR, (4 + i)), Cells(endR, (4 + i))).FillDown
                Cells(10 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R5C/2) * 3.14))) "
                Cells(11 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R5C/2) * 3.14))) / R14C"
                Cells(12 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R5C/2) * 3.14))) / R24C"
                Cells(11, (4 + i)).Value = "PGL"
            End If
        ElseIf 0 < Cells(7, (4 + i)).Value < 1 And Cells(9, 103).Value = "MultiPak" Then    ' GL multipak form: SGL and TSGL
            Cells(5, (4 + i)).Value = Cells(4, (4 + i)).Value
            If Cells(7, (4 + i)).Font.Italic = "False" And Cells(7, (4 + i)).Font.Underline = xlUnderlineStyleSingle Then   ' exponential asymmetric blend based Voigt (GL multipak)
                Cells(8, (4 + i)).Value = 0.35      ' initial parameters
                Cells(9, (4 + i)).Value = 10
                Debug.Print "non-italic underline multipak"
                For k = 1 To numData        ' ' R8C: Tail coefficient, R9C: Half Tail length at half maximum
                    If Cells((startR - 1 + k), 1).Value >= Cells(2, (4 + i)).Value Then
                        Cells((startR - 1 + k), (4 + i)).FormulaR1C1 = "=R6C * ((R7C)*((((R4C)/2)^2)/((RC[" & (-3 - i) & "]-R2C)^2 + ((R4C)/2)^2)) + (1- R7C)*(EXP(-(1/2)*((RC[" & (-3 - i) & "]-R2C)/(R4C/2.35))^2)) + (R8C * (1 - EXP(-(1/2)*((RC[" & (-3 - i) & "]-R2C)/(R4C/2.35))^2)) * exp((-6.9/R9C) * (2 * (RC[" & (-3 - i) & "] - R2C))/R4C)))"
                    Else
                        Cells((startR - 1 + k), (4 + i)).FormulaR1C1 = "=R6C * ((R7C)*((((R4C)/2)^2)/((RC[" & (-3 - i) & "]-R2C)^2 + ((R4C)/2)^2)) + (1- R7C)*(EXP(-(1/2)*((RC[" & (-3 - i) & "]-R2C)/(R4C/2.35))^2)))"
                    End If
                Next
                
                Cells(10 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R4C/2) * 3.14))) "
                Cells(11 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R4C/2) * 3.14))) / R14C"
                Cells(12 + sftfit2, (4 + i)).FormulaR1C1 = "=((1-R7C)*(SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)) + R7C*((R6C * (R4C/2) * 3.14))) / R24C"
                Cells(11, (4 + i)).Value = "TSGL"
            Else
                Debug.Print "GL multipak"     ' MultiPak GL sum form with a single FWHM for G and L
                Cells(startR, (4 + i)).FormulaR1C1 = "=R6C * ((R7C)*((((R4C)/2)^2)/((RC[" & (-3 - i) & "]-R2C)^2 + ((R4C)/2)^2)) + (1- R7C)*(EXP(-(1/2)*((RC[" & (-3 - i) & "]-R2C)/(R4C/2.35))^2)))"
                Range(Cells(startR, (4 + i)), Cells(endR, (4 + i))).FillDown
                Cells(10 + sftfit2, (4 + i)).FormulaR1C1 = "=SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14)"
                Cells(11 + sftfit2, (4 + i)).FormulaR1C1 = "=SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14) / R14C"
                Cells(12 + sftfit2, (4 + i)).FormulaR1C1 = "=SQRT(2) * (R4C/2.35) * R6C * SQRT(3.14) / R24C"
                Cells(11, (4 + i)).Value = "GL"
            End If
        End If
        
        Cells(20 + sftfit, (4 + i)).FormulaR1C1 = "=R1C" ' Peak name
        Cells(3, (4 + i)).FormulaR1C1 = "=(R12C101 - R13C101 - R14C101 - R2C)" ' KE " & (pe - wf - char) & "
        Cells((numData + 23 + sftfit), (4 + i)).FormulaR1C1 = "=R[" & (-numData - 2) & "]C + R[" & (-numData - 2) & "]C[" & -i - 1 & "]"      ' Peak + BG
        Range(Cells((numData + 23 + sftfit), (4 + i)), Cells((2 * numData + 22 + sftfit), (4 + i))).FillDown
        Cells((numData + 22 + sftfit), (4 + i)).FormulaR1C1 = "=R1C" ' Peak name"
        Cells(8 + sftfit2, (4 + i)).FormulaR1C1 = "=R6C + " & dblMin
    Next

    Cells(startR, (5 + j)).FormulaR1C1 = "=SUM(RC[" & -j & "]:RC[-1])"      ' sum of peaks
    Range(Cells(startR, (5 + j)), Cells(endR, (5 + j))).FillDown
    Cells((numData + 23 + sftfit), (5 + j)).FormulaR1C1 = "=R[" & (-numData - 2) & "]C + R[" & (-numData - 2) & "]C[" & -j - 2 & "]"    ' Sum of Peaks + BG
    Range(Cells((numData + 23 + sftfit), (4 + i)), Cells((2 * numData + 22 + sftfit), (4 + i))).FillDown
    Cells((numData + 22 + sftfit), (4 + i)).Value = "peaks+BG"
    Cells(startR, (6 + j)).FormulaR1C1 = "=((RC2 - R[" & (2 + numData) & "]C[-1])^2)/(abs(R[" & (2 + numData) & "]C[-1]))"     ' Least fits 2
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
    
    For i = 1 To j
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(3 + i)
            .ChartType = xlXYScatterLinesNoMarkers
            .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C" & (4 + i) & ""
            .XValues = rng
            .Values = rng.Offset((numData + 2), (3 + i))
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
        i = 0
        Set pts = .Points
        For Each pt In pts
            i = i + 1
            With pt.DataLabel
                .Text = Range(Cells(1, 5), Cells(1, 5).Offset(0, (j - 1))).Cells(i).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 12
            End With
        Next
    End With
    
    For i = 1 To j
        Cells(1, (4 + i)).Interior.Color = ActiveChart.SeriesCollection(i + 3).Border.Color
        Cells(1, (4 + i)).Font.ColorIndex = 2
    Next
    
    If ActiveSheet.ChartObjects.Count = 1 Then Exit Sub
    
    ActiveSheet.ChartObjects(2).Activate
    k = ActiveChart.SeriesCollection.Count
    For i = k To 2 Step -1
        ActiveChart.SeriesCollection(i).Delete
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
    
    For i = 1 To j
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(2 + i)
            .ChartType = xlXYScatterLinesNoMarkers
            .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C" & (4 + i) & ""
            .XValues = rng
            .Values = rng.Offset(0, (3 + i))
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
        i = 0
        Set pts = .Points
        For Each pt In pts
            i = i + 1
            With pt.DataLabel
                .Text = Range(Cells(1, 5), Cells(1, 5).Offset(0, (j - 1))).Cells(i).Value
                .Position = xlLabelPositionAbove
                .Font.Size = 12
            End With
        Next
    End With
    
    For i = 1 To j
        ActiveChart.SeriesCollection(i + 2).Border.Color = Cells(1, (4 + i)).Interior.Color
    Next
    
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart
        With .SeriesCollection(4 + j)
            .ChartType = xlXYScatterLinesNoMarkers
            '.ChartType = xlAreaStacked
            .Name = "='" & ActiveSheet.Name & "'!R" & (20 + sftfit) & "C" & (6 + i) & ""
            .XValues = rng
            '.Values = rng.Offset(0, (4 + j))
            .Values = rng.Offset(, (6 + j))
            .AxisGroup = xlSecondary
            .Border.ColorIndex = 44
            .Format.Line.Weight = 2
            .HasDataLabels = False
        End With
    End With
    
    ActiveChart.HasAxis(xlCategory, xlSecondary) = True
    With ActiveChart.Axes(xlCategory, xlSecondary)
        If StrComp(str1, "Pe", 1) = 0 Then
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
    For i = k To 5 + j Step -1
        ActiveChart.SeriesCollection(i).Delete
    Next
    
    ActiveSheet.ChartObjects(1).Activate
End Sub

Sub Delsheets()
    If ExistSheet(strSheetXPSFactors) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetXPSFactors).Visible = xlSheetVisible
        Worksheets(strSheetXPSFactors).Delete
        Application.DisplayAlerts = True
    End If
    If ExistSheet(strSheetAESFactors) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetAESFactors).Visible = xlSheetVisible
        Worksheets(strSheetAESFactors).Delete
        Application.DisplayAlerts = True
    End If
    If ExistSheet(strSheetPICFactors) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetPICFactors).Visible = xlSheetVisible
        Worksheets(strSheetPICFactors).Delete
        Application.DisplayAlerts = True
    End If
    If ExistSheet(strSheetChemFactors) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetChemFactors).Visible = xlSheetVisible
        Worksheets(strSheetChemFactors).Delete
        Application.DisplayAlerts = True
    End If
End Sub

Sub numMajorUnitsCheck()
    If Abs(startEk - endEk) <= 100 And Abs(startEk - endEk) > 50 Then
            numMajorUnit = 4 * windowSize
        ElseIf Abs(startEk - endEk) <= 50 And Abs(startEk - endEk) > 20 Then
            numMajorUnit = 2 * windowSize
        ElseIf Abs(startEk - endEk) > 100 Then
            numMajorUnit = 50 * windowSize
        ElseIf Abs(startEk - endEk) <= 20 Then
            numMajorUnit = 1 * windowSize
        End If
End Sub

Sub scalecheck()
    With Application
        If str3 = "De" Then
            j = 1
        Else
            j = 0
        End If
        
        startEb = Cells(20 + (numData), 2 - j).Value
        endEb = Cells(20 + (numData), 2 - j).Offset(numData - 1, 0).Value
        If Abs(startEb - endEb) <= 100 And Abs(startEb - endEb) > 50 Then
            numMajorUnit = 4 * windowSize
        ElseIf Abs(startEb - endEb) <= 50 And Abs(startEb - endEb) > 20 Then
            numMajorUnit = 2 * windowSize
        ElseIf Abs(startEb - endEb) > 100 Then
            numMajorUnit = 50 * windowSize
        ElseIf Abs(startEb - endEb) <= 20 Then
            numMajorUnit = 1 * windowSize
        End If
    
        If str1 = "Pe" Or str3 = "De" Or str1 = "Po" Then
            If startEb < 0 Then
                startEb = .Ceiling(startEb, (-1 * numMajorUnit))
            Else
                startEb = .Floor(startEb, numMajorUnit)
            End If
        ElseIf startEb > 0 Then
            startEb = .Ceiling(startEb, numMajorUnit)
        Else
            startEb = .Floor(startEb, (-1 * numMajorUnit))
        End If

        If str1 = "Pe" Or str3 = "De" Or str1 = "Po" Then
            If endEb < 0 Then
                endEb = .Floor(endEb, (-1 * numMajorUnit))
            Else
                endEb = .Ceiling(endEb, numMajorUnit)
            End If
        ElseIf endEb > 0 Then
            endEb = .Floor(endEb, numMajorUnit)
        Else
            endEb = .Ceiling(endEb, (-1 * numMajorUnit))
        End If
        
        dblMax = .Max(dataIntGraph)
        dblMin = .Min(dataIntGraph)
        
        If str3 = "De" Then
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
                
                If InStr(1, chkMax, ".") Then
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
    iniTime = Timer
    TimeC1 = TimeC2
    finTime = startTime
    strCpa = ""
    strLabel = ""
    strAna = ""
    strCasa = ""
    strAES = ""
    strErr = ""
    strErrX = ""
    strscanNum = ""
    strscanNumR = ""
    pe = 0
    off = 0
    multi = 1
    startR = 0
    endR = 0
    g = 0
    ReDim Preserve highpe(0)
    strSheetXPSFactors = "XPSFactors"
    strSheetAESFactors = "AESFactors"
    strSheetPICFactors = "PICFactors"
    strSheetChemFactors = "ChemFactors"
    
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
    Dim chkRef As Object
    With Application.AddIns
    For i = 1 To .Count
        If LCase(.Item(i).Name) = "solver.xlam" Then
            If Len(.Item(i).FullName) > 10 Then
                If AddIns("Solver Add-In").Installed = True Then
                    Exit For
                Else
                    MsgBox "No solver installed in Excel Add-in!" & vbCrLf & " Go to Excel Options - Add-Ins - Go Manage - Solver to be checked."
                    End
                End If
            End If
        ElseIf i = .Count And LCase(.Item(i).Name) <> "solver.xlam" Then
            MsgBox "No solver found in Excel Add-in!" & vbCrLf & " Go to Excel Options - Add-Ins - Go Manage - Solver.xlam to be browsed."
            End
        End If
    Next i
    End With
    
    ' SolverInstall1() or SolverInstall2() to be run for setup Solver by code.
    Call Delsheets
End Sub

Function WorkbookOpen(WorkBookName As String) As Boolean
    WorkbookOpen = False
    On Error GoTo WorkBookNotOpen
    If Len(Application.Workbooks(WorkBookName).Name) > 0 Then
        WorkbookOpen = True
        Exit Function
    End If
WorkBookNotOpen:
End Function

Sub SolverSetup()
    SolverReset ' Error due to the Solver installation! Check the Solver function correctly installed.
    SolverOptions MaxTime:=100, Iterations:=32767, Precision:=0.000001, AssumeLinear _
        :=False, StepThru:=False, Estimates:=1, Derivatives:=1, SearchOption:=1, _
        IntTolerance:=5, Scaling:=False, Convergence:=(0.0001 / Cells(3, 101).Value), AssumeNonNeg:=False
End Sub

Sub GetNormalize()
    strSheetAnaName = "Norm_" + strSheetDataName
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
    
    If Cells(1, 1).Value = "norm" Then
        i = 1   ' means data to be generated on third set of data column
        k = 1   ' means data to be normalized on first set of data column
        off = Cells(9, (5 + (i * 3)))
        multi = Cells(9, (6 + (i * 3)))
        If multi = 0 Then
            multi = 1
        End If
        
        sheetGraph.Range(Cells(1, (4 + (i * 3))), Cells((2 * (numData + 10)) - 1, (6 + (i * 3)))).Clear
        Set rng = Range(Cells(11, (k + 1 + ((0) * 3))), Cells(11, (k + 1 + (0 * 3))).End(xlDown))
        numData = Application.CountA(rng)
        Set rng = Range(Cells(11, (k + 1 + ((1) * 3))), Cells(11, (k + 1 + (1 * 3))).End(xlDown))
        iCol = Application.CountA(rng)
        C = sheetGraph.Range(Cells(11 + numData + 9, (k + 1 + (0 * 3))), Cells(11 + (numData * 2) + 8, (k + 2 + (0 * 3))))
        A = sheetGraph.Range(Cells(11 + iCol + 9, (2 + (i * 3))), Cells(11 + (iCol * 2) + 8, (3 + (i * 3))))
        D = sheetGraph.Range(Cells(11, (1 + ((i + 1) * 3))), Cells(10 + numData, (3 + ((i + 1) * 3))))
        stepEk = Cells(7, (k + 1 + (0 * 3))).Value
        endEk = Cells(7, (k + 1 + (1 * 3))).Value

        p = 1
        For q = 1 To numData
            For j = 1 To iCol
                If C(q, 1) > A(j, 1) - (endEk / 2) And C(q, 1) < A(j, 1) + (endEk / 2) Then
                    D(p, 1) = C(q, 1)
                    If A(j, 2) <> 0 Then
                            D(p, 3) = C(q, 2) / A(j, 2) ' here is normalized
                        Else
                            D(p, 3) = "NaN"
                        End If
                    p = p + 1
                    Exit For
                End If
            Next

            If j = iCol + 1 And endEk < stepEk Then
                'Debug.Print "rough mode", C(q, 1)
                For j = 1 To iCol
                    If C(q, 1) > A(j, 1) - (stepEk / 2) And C(q, 1) < A(j, 1) + (stepEk / 2) Then
                        D(p, 1) = C(q, 1)
                        If A(j, 2) <> 0 Then
                            D(p, 3) = C(q, 2) / A(j, 2) ' here is normalized
                        Else
                            D(p, 3) = "NaN"
                        End If
                        p = p + 1
                        Exit For
                    End If
                Next
            Else
                
            End If
        Next
        
        numData = p - 1
        imax = numData + 10
        sheetGraph.Range(Cells(11, (1 + ((i + 1) * 3))), Cells(10 + numData, (3 + ((i + 1) * 3)))) = D
        str1 = "Pe"
        str2 = "Sh"
        str3 = "Ab"
        strTest = strSheetDataName + "_norm"
        Cells(1, (5 + (i * 3))).Value = strTest
        Cells(8 + (imax), (5 + (i * 3))).Value = strTest
        Cells(9 + (imax), (4 + (i * 3))).Value = str1 + strTest
        Cells(9 + (imax), (5 + (i * 3))).Value = str2 + strTest
        Cells(9 + (imax), (6 + (i * 3))).Value = str3 + strTest
        Cells(2, ((4 + (i * 3)))).Value = "PE shifts"
        Cells(2, ((5 + (i * 3)))).Value = 0
        Cells(2, ((6 + (i * 3)))).Value = "eV"
        Cells(5, ((4 + (i * 3)))).Value = "Start PE"
        Cells(6, ((4 + (i * 3)))).Value = "End PE"
        Cells(7, ((4 + (i * 3)))).Value = "Step PE"
        Cells(5, ((5 + (i * 3)))).Value = Cells(11, 7).Value
        Cells(6, ((5 + (i * 3)))).Value = Cells(10 + numData, 7).Value
        Cells(7, ((5 + (i * 3)))).Value = Cells(12, 7).Value - Cells(11, 7).Value
        Range(Cells(5, 9), Cells(7, 9)) = "eV"
        Cells(9, ((4 + (i * 3)))).Value = "Offset/multp"
        Cells(9, ((5 + (i * 3)))).Value = off
        Cells(9, ((6 + (i * 3)))).Value = multi
        Cells(10, ((4 + (i * 3)))).Value = "PE"
        Cells(10, ((5 + (i * 3)))).Value = "+shift"
        Cells(10, ((6 + (i * 3)))).Value = "Ab"
        Range(Cells(5, (4 + (i * 3))), Cells(7, (4 + (i * 3)))).Interior.ColorIndex = 41
        Range(Cells(5, (5 + (i * 3))), Cells(7, (6 + (i * 3)))).Interior.ColorIndex = 37
        Range(Cells(2, (4 + (i * 3))), Cells(2, (4 + (i * 3)))).Interior.ColorIndex = 3
        Range(Cells(2, (5 + (i * 3))), Cells(2, (6 + (i * 3)))).Interior.ColorIndex = 38
        Range(Cells(9, (4 + (i * 3))), Cells(9, ((4 + (i * 3))))).Interior.ColorIndex = 43
        Range(Cells(9, (5 + (i * 3))), Cells(9, ((6 + (i * 3))))).Interior.ColorIndex = 35
        Cells(11, (5 + (i * 3))).FormulaR1C1 = "=R2C + RC[-1]"
        Cells(10 + (imax), (5 + (i * 3))).FormulaR1C1 = "=R2C + R[-" & (imax - 1) & "]C[-1]"
        Range(Cells(11, (5 + (i * 3))), Cells((imax), (5 + (i * 3)))).FillDown
        Cells(10 + (imax), (4 + (i * 3))).FormulaR1C1 = "=R[-" & (imax - 1) & "]C"
        Range(Cells(10 + (imax), (4 + (i * 3))), Cells((2 * imax) - 1, (4 + (i * 3)))).FillDown
        Range(Cells(10 + (imax), (5 + (i * 3))), Cells((2 * imax) - 1, (5 + (i * 3)))).FillDown
        Cells(10 + (imax), (6 + (i * 3))).FormulaR1C1 = "= (R[-" & (imax - 1) & "]C - R9C[-1])*R9C"
        Range(Cells(10 + (imax), (6 + (i * 3))), Cells((2 * imax) - 1, (6 + (i * 3)))).FillDown
        Set dataKeGraph = Range(Cells(10 + (imax), (4 + (i * 3))), Cells((2 * imax - 1), (4 + (i * 3))))
        ActiveSheet.ChartObjects(1).Activate
        p = ActiveChart.SeriesCollection.Count
        For j = 1 To p
            If ActiveChart.SeriesCollection(j).Name = Cells(1, 5 + (i * 3)).Value Then
                ActiveChart.SeriesCollection(j).Delete
                p = p - 1
                Exit For
            End If
        Next
        
        ActiveChart.SeriesCollection.NewSeries
        With ActiveChart.SeriesCollection(p + i)
            .ChartType = xlXYScatterLinesNoMarkers
            .Name = Cells(1, 5 + (i * 3)).Value
            .XValues = dataKeGraph.Offset(0, 1)
            .Values = dataKeGraph.Offset(0, 2)
            SourceRangeColor1 = .Border.Color
        End With
        
        Range(Cells(10, (4 + (i * 3))), Cells(10, ((4 + (i * 3))))).Interior.Color = SourceRangeColor1
        Range(Cells(9 + (imax), (4 + (i * 3))), Cells(9 + (imax), ((4 + (i * 3))))).Interior.Color = SourceRangeColor1
        Range(Cells(10, (5 + (i * 3))), Cells(10, ((5 + (i * 3))))).Interior.Color = SourceRangeColor1
        Range(Cells(9 + (imax), (5 + (i * 3))), Cells(9 + (imax), ((5 + (i * 3))))).Interior.Color = SourceRangeColor1
        sheetGraph.Range(Cells(11 + numData + 8, (5 + (i * 3))), Cells(11 + (numData * 2) + 8, (6 + (i * 3)))).Copy
        sheetAna.Cells(1, 1 + ((i - 1) * 2)).PasteSpecial Paste:=xlValues
        sheetAna.Cells(1, 1).Value = "PE/eV"
        sheetGraph.Activate
        If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    End If
    Application.CutCopyMode = False
    strErr = "skip"
End Sub

Sub ExportKratos()
    Set rng = Range(Cells(1, 1), Cells(1, 1).End(xlDown))
    iRow = Application.CountA(rng) - 4
    strSheetDataName = ActiveSheet.Name
    strCpa = ActiveWorkbook.Path
    Set sheetData = Worksheets(strSheetDataName)
    
    Do
        If InStr(strSheetDataName, " ") > 0 Then
            strSheetDataName = mid$(strSheetDataName, 1, InStr(strSheetDataName, " ") - 1) + mid$(strSheetDataName, InStr(strSheetDataName, " ") + 1, Len(strSheetDataName))
        ElseIf InStr(strSheetDataName, " ") = 0 Then
            Exit Do
        End If
    Loop
    
    strSheetAnaName = mid$(strCpa, InStrRev(strCpa, "\") + 1, Len(strCpa) - InStrRev(strCpa, "\")) + "_" + strSheetDataName
    If ExistSheet(strSheetAnaName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetAnaName).Delete
        Application.DisplayAlerts = True
    End If
    
    Worksheets.Add().Name = strSheetAnaName
    Set sheetAna = Worksheets(strSheetAnaName)
    wb = strSheetAnaName
    sheetData.Activate
    sheetData.Range(Cells(5, 2), Cells(iRow + 4, 3)).Copy
    sheetAna.Activate
    sheetAna.Cells(2, 1).PasteSpecial Paste:=xlValues
    sheetAna.Cells(1, 1).Value = "BE/eV"
    sheetAna.Cells(1, 2).Value = "AlKa"
    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    Application.CutCopyMode = False
    strErr = "skip"
End Sub

Sub ExportPHI()
    Dim numDataT As Integer
    Dim numDataF As Integer
    Dim ElemT As String
    
    Set rng = [3:3]
    iCol = Application.CountA(rng)
    Set rng = [4:4]
    iRow = Application.CountA(rng) / iCol
    strCpa = ActiveWorkbook.Path
    strSheetAnaName = ActiveSheet.Name
    Set sheetAna = Worksheets(strSheetAnaName)
    ElemT = vbNullString
    numDataF = FreeFile
    ' http://www.homeandlearn.org/write_to_a_text_file.html
    For i = 0 To iCol - 1
        Set rng = sheetAna.Range(Cells(5, 2 + (i * (1 + iRow))), Cells(5, (2 + (i * (1 + iRow)))).End(xlDown))
        numDataT = Application.CountA(rng)
        
        For p = 0 To iRow - 1
            If iRow = 1 Then
                strLabel = strSheetAnaName & "_" & sheetAna.Cells(3, 1 + (i * (1 + iRow))).Value
            Else
                strLabel = strSheetAnaName & "_" & sheetAna.Cells(3, 1 + (i * (1 + iRow))).Value & "_d" & sheetAna.Cells(4, 2 + p + (i * (1 + iRow))).Value
            End If
            strTest = strCpa & "\" & strLabel & ".txt"
            
            If Dir(strTest) <> "" And i > 0 And iRow = 1 Then
                strTest = strCpa & "\" & strLabel & i & ".txt"
            End If
            
            Open strTest For Output As #numDataF
            For j = 4 To numDataT + 4
                If j = 4 Then
                    ElemT = "BE/eV" + vbTab + "AlKa"
                Else
                    ElemT = Trim(sheetAna.Cells(j, 1 + (i * (1 + iRow))).Value) + vbTab + Trim(sheetAna.Cells(j, 2 + p + (i * (1 + iRow))).Value)
                End If
                Print #numDataF, ElemT
                ElemT = vbNullString
            Next j
            Close #numDataF
            numDataF = numDataF + 1
        Next p
    Next i
    
    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    Application.CutCopyMode = False
    strErr = "skip"
    Cells(1, 1).Value = "Exported"
End Sub

Sub TransposeSheet()      ' Cells(1, 1) = "transpose" to do this macro!
    Cells(1, 1) = "Exported to TP"
    C = ActiveSheet.UsedRange
    D = Application.Transpose(C)
    strSheetGraphName = "TP_" & strSheetDataName
    If ExistSheet(strSheetGraphName) Then
        Application.DisplayAlerts = False
        Worksheets(strSheetGraphName).Delete
        Application.DisplayAlerts = True
    End If
    
    Worksheets.Add().Name = strSheetGraphName
    Set sheetGraph = Worksheets(strSheetGraphName)
    sheetGraph.Activate
    sheetGraph.Range(Cells(1, 1), Cells(UBound(C, 2), UBound(C, 1))).Value = D
    Cells(1, 1) = vbNullString
    If StrComp(strErr, "skip", 1) = 0 Then Exit Sub
    Application.CutCopyMode = False
    strErr = "skip"
End Sub

Sub CombineLegend()
    Dim spr As String
    Dim strSheetSampleName As String, strSheetTargetName As String, strSeriesName As String
    Dim sheetSample As Worksheet, sheetTarget As Worksheet
    Dim icur As Integer, kcur As Integer
    
    If mid$(Results, 1, 1) = "i" Then
        Debug.Print "i", mid$(Results, 2, Len(Results) - InStr(1, Results, "k")), mid$(Results, InStr(1, Results, "k") + 1, Len(Results) - InStr(1, Results, "k"))
        icur = CInt(mid$(Results, 2, Len(Results) - InStr(1, Results, "k"))) ' number of comp in each comp
        kcur = CInt(mid$(Results, InStr(1, Results, "k") + 1, Len(Results) - InStr(1, Results, "k")))            ' position of comp from 0
        Results = vbNullString
    Else
        icur = -1
        kcur = -1
    End If
    
    If Cells(40, para + 9).Value = "Ver." Then
    Else
        For i = 1 To 500
            If StrComp(Cells(40, i + 9).Value, "Ver.", 1) = 0 Then Exit For
        Next
        para = i
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

            For i = 0 To ncomp
                Cells(2 + i, 1).Value = i + 1
                If i > 0 And InStr(1, sheetTarget.Cells(1, 2 + i * 3), spr) > 0 Then
                    Cells(2 + i, 2).Value = mid$(sheetTarget.Cells(1, 2 + i * 3).Value, 1, InStr(1, sheetTarget.Cells(1, 2 + i * 3).Value, spr) - 1)
                    Cells(2 + i, 4).Value = mid$(sheetTarget.Cells(1, 2 + i * 3).Value, InStr(1, sheetTarget.Cells(1, 2 + i * 3).Value, spr) + Len(spr), Len(sheetTarget.Cells(1, 2 + i * 3).Value))
                ElseIf i = 0 Then
                    sheetTarget.Activate
                    If ActiveSheet.ChartObjects.Count > 0 Then
                        ActiveSheet.ChartObjects(1).Activate
                        strSeriesName = ActiveChart.SeriesCollection(1).Name
                        If InStr(1, ActiveChart.SeriesCollection(1).Name, spr) > 0 Then
                            sheetSample.Activate
                            Cells(2 + i, 2).Value = mid$(strSeriesName, 1, InStr(1, strSeriesName, spr) - 1)
                            Cells(2 + i, 4).Value = mid$(strSeriesName, InStr(1, strSeriesName, spr) + Len(spr), Len(strSeriesName))
                        Else
                            sheetSample.Activate
                            Cells(2 + i, 2).Value = "no." & i + 1
                            Cells(2 + i, 4).Value = strSeriesName
                        End If
                    End If
                Else
                    Cells(2 + i, 2).Value = "no." & i + 1
                    Cells(2 + i, 4).Value = sheetTarget.Cells(1, 2 + i * 3).Value
                End If
                Cells(2 + i, 3).Value = spr
            Next
        Else
            Set sheetSample = Worksheets(strSheetSampleName)
            sheetSample.Activate
            
            If ncomp + 2 > sheetSample.UsedRange.Rows.Count Then
                For i = sheetSample.UsedRange.Rows.Count - 1 To ncomp
                    Cells(2 + i, 1).Value = i + 1
                    Cells(2 + i, 2).Value = "no." & i + 1
                    Cells(2 + i, 3).Value = spr
                    Cells(2 + i, 4).Value = sheetTarget.Cells(1, 2 + i * 3).Value
                Next
            ElseIf kcur >= 0 And kcur + 3 < sheetSample.UsedRange.Rows.Count Then
                For i = kcur + 1 To icur
                    Cells(2 + i, 1).Value = i + 1
                    Cells(2 + i, 2).Value = "no." & i + 1
                    Cells(2 + i, 3).Value = spr
                    Cells(2 + i, 4).Value = sheetTarget.Cells(1, 2 + i * 3).Value
                Next
            End If
        End If
        
        Set sheetSample = Worksheets(strSheetSampleName)
        sheetTarget.Activate
                
        For i = 0 To ncomp - 1
            sheetTarget.Cells(1, 5 + i * 3) = sheetSample.Cells(i + 3, 2).Value & spr & sheetSample.Cells(i + 3, 4).Value
        Next
        
        If ActiveSheet.ChartObjects.Count > 0 Then
            For i = 0 To ActiveSheet.ChartObjects.Count - 1
                ActiveSheet.ChartObjects(1 + i).Activate
                With ActiveChart.SeriesCollection(1)
                    .Name = sheetSample.Cells(2, 2).Value & spr & sheetSample.Cells(2, 4).Value
                End With
            Next
        End If
    End If
    
    sheetTarget.Activate
    Cells(1, 1).Value = "Grating"
End Sub

Sub descriptGConv()
    For k = 1 To (endR - startR + 1)
            Cells(startR, 100 + k).FormulaR1C1 = "=RC3 * Exp(-(1/2)*((RC1-R" & (startR + k - 1) & "C1)/(R6C5/2.35))^2)" ' CV
            Range(Cells(startR, 100 + k), Cells(endR, 100 + k)).FillDown
            Cells(startR + k - 1, 100).FormulaR1C1 = "=Sum(R" & (startR) & "C" & (100 + k) & ":R" & (endR) & "C" & (100 + k) & ")"
    Next k
End Sub

Sub descriptTConv()
    If Cells(20 + sftfit, 2).Value = "Ab" Then ' for PE
        For k = 1 To (endR - startR + 1)
                Cells(startR + k - 1, 110 + k).FormulaR1C1 = "=((RC2 * R2C2 * (RC1 -R" & (startR + k - 1) & "C1))/((R3C2 + (" & p & " * (RC1 -R" & (startR + k - 1) & "C1)^2))^2 + R4C2 * ((RC1 -R" & (startR + k - 1) & "C1)^2)))" ' CV
                Range(Cells(startR + k - 1, 110 + k), Cells(endR, 110 + k)).FillDown
                Cells(startR + k - 1, 109).FormulaR1C1 = "=Sum(R" & (startR + k - 1) & "C" & (110 + 1) & ":R" & (startR + k - 1) & "C" & (110 + endR - startR + 1) & ")"
                Cells(startR + k - 1, 3).FormulaR1C1 = "=R5C2 * (Sum(R" & (startR) & "C" & (109) & ":R" & (startR + k - 1) & "C" & (109) & ")/(" & (endR - startR + 1) & ") + R6C2)"
        Next k
    Else
        For k = 1 To (endR - startR + 1)
                Cells(endR - k + 1, 110 + k).FormulaR1C1 = "=((RC2 * R2C2 * (RC1 -R" & (endR - k + 1) & "C1))/((R3C2 + (" & p & " * (RC1 -R" & (endR - k + 1) & "C1)^2))^2 + R4C2 * ((RC1 -R" & (endR - k + 1) & "C1)^2)))" ' CV
                Range(Cells(startR, 110 + k), Cells(endR - k + 1, 110 + k)).FillUp
                Cells(startR + k - 1, 109).FormulaR1C1 = "=Sum(R" & (startR + k - 1) & "C" & (110 + 1) & ":R" & (startR + k - 1) & "C" & (110 + endR - startR + 1) & ")"
                Cells(endR - k + 1, 3).FormulaR1C1 = "=R5C2 * (Sum(R" & (endR - k + 1) & "C" & (109) & ":R" & (endR) & "C" & (109) & ")/(" & (endR - startR + 1) & ") + R6C2)"
        Next k
    End If
End Sub

Sub UDsamples()
    Sheets("Sheet1").Name = "XPS"
    Sheets("Sheet2").Name = "AES"
    Sheets("Sheet3").Name = "Notes"
    
    Sheets("XPS").Activate
    Cells(1, 1).Value = "Element"
    Cells(1, 2).Value = "Orbit"
    Cells(1, 3).Value = "BE (eV)"
    Cells(1, 4).Value = "ASF"
    Cells(2, 1).Value = "C"
    Cells(2, 2).Value = "1s"
    Cells(2, 3).Value = "284.6"
    Cells(2, 4).Value = "1"
    Cells(3, 1).Value = "O"
    Cells(3, 2).Value = "1s"
    Cells(3, 3).Value = "532"
    Cells(3, 4).Value = "2.93"
    Sheets("AES").Activate
    Cells(1, 1).Value = "Element"
    Cells(1, 2).Value = "Auger"
    Cells(1, 3).Value = "KE (eV)"
    Cells(1, 4).Value = "RSF"
    Cells(2, 1).Value = "C"
    Cells(2, 2).Value = "KLL"
    Cells(2, 3).Value = "266"
    Cells(2, 4).Value = "0.6"
    Cells(3, 1).Value = "O"
    Cells(3, 2).Value = "KLL"
    Cells(3, 3).Value = "506"
    Cells(3, 4).Value = "0.96"
End Sub

Sub debugAll()      ' multiple file analysis in sequence
    Dim be4all() As Variant, am4all() As Variant, fw4all() As Variant, wbX As String, shgX As Worksheet, shfX As Worksheet, strSheetDataNameX As String, numpeakX As Integer
    Dim debugMode As String, seriesnum As Integer
    
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
        End If
    Else
        modex = 0
    End If
    
    If modex >= -2 And modex <= 6 Then
    Else
        Call GetOut
        Exit Sub
    End If
    
    strErrX = ""
    If modex <= -1 Then
        ChDrive mid$(ActiveWorkbook.Path, 1, 1)
        ChDir ActiveWorkbook.Path
    End If

    If modex = -2 Then
        OpenFileName = Application.GetOpenFilename(FileFilter:="Text Files (*.xlsx), *.xlsx", Title:="Please select file(s)", MultiSelect:=True)
    Else
        OpenFileName = Application.GetOpenFilename(FileFilter:="Text Files (*.txt), *.txt,MultiPak Files (*.csv), *.csv", Title:="Please select file(s)", MultiSelect:=True)
    End If

    If IsArray(OpenFileName) Then
        If UBound(OpenFileName) >= 1 Then
            For Each Target In OpenFileName
                If mid$(Target, Len(Target) - 2, 3) = "csv" Then modex = 1
            Next
        End If
    Else
        Exit Sub
    End If
    
    If modex <= -1 Then
        wb = ActiveWorkbook.Name
        wbX = wb
        strSheetDataNameX = strSheetDataName
        Set shgX = Workbooks(wbX).Sheets("Graph_" + strSheetDataNameX)
        peX = Workbooks(wb).Sheets("Graph_" + strSheetDataName).Cells(2, 2).Value
        If debugMode = "debugFit" Or debugMode = "debugShift" Then
            Set shfX = Workbooks(wbX).Sheets("Fit_" + strSheetDataNameX)
            numpeakX = Workbooks(wb).Sheets("Fit_" + strSheetDataName).Cells(8 + sftfit2, 2).Value
        ElseIf debugMode = "debugPara" Then
            Set shfX = Workbooks(wbX).Sheets("Fit_" + strSheetDataNameX)
            tmp = Workbooks(wb).Sheets("Fit_" + strSheetDataName).Range(Cells(14 + sftfit2, 1), Cells(19 + sftfit2, 2)).Value
        End If
    End If
    
    If modex = -1 Then
        ElemX = Workbooks(wbX).Sheets("Graph_" + strSheetDataName).Cells(51, para + 9).Value
    ElseIf modex = 1 Or modex = -2 Then
    Else
        ElemX = Application.InputBox(Title:="Input atomic elements", Prompt:="Example:C,O,Co,etc ... without space!", Default:="C,O,Au", Type:=2)
    End If
    
    If modex = 1 Or modex = -2 Then
    Else
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
        strTest = mid$(Target, InStrRev(Target, "\") + 1, Len(Target) - InStrRev(Target, "\"))
        If Not WorkbookOpen(strTest) Then
            Workbooks.Open Target
        Else
            Workbooks(strTest).Activate
            strLabel = ActiveSheet.Name
        End If
        
        If modex = 1 Then
            Cells(1, 2).Value = Cells(1, 1).Value
            If IsEmpty(Cells(4, 1).Value) Then    ' Cells(4, 1).Value = 0
                Cells(1, 1).Value = "phi"
            Else
                If Cells(4, 1).Value = 0 Then
                    Cells(1, 1).Value = "contour"
                    Application.DisplayAlerts = False
                    wb = mid$(Target, 1, Len(Target) - 4)   ' file name for these images is based on csv
                    ActiveWorkbook.SaveAs Filename:=wb, FileFormat:=51
                    Workbooks(ActiveWorkbook.Name).Close SaveChanges:=False
                    Application.DisplayAlerts = True
                    GoTo SkipOpenDebug
                Else
                    Cells(1, 1).Value = "phi"
                End If
            End If
        ElseIf modex = -2 Then
            Application.DisplayAlerts = False
            strSheetDataName = mid$(Target, InStrRev(Target, "\") + 1, Len(Target) - InStrRev(Target, "\") - 5)
            Debug.Print strSheetDataName
            Workbooks(ActiveWorkbook.Name).Sheets("Fit_" + strSheetDataName).Range(Cells(14 + sftfit2, 1), Cells(19 + sftfit2, 2)) = tmp
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
        End If
        
        ' 1st Code to run in each Target
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
        
        If modex = -1 Then
            testMacro = "debug"     ' This is a trigger to run the debugAll code in sequence
            sheetGraph.Activate     ' activate Graph sheet
            shgX.Activate
            If debugMode = "debugGraphn" Then
                Set rng = [D:D]
                numpeakX = (Application.CountA(rng) - 8) / 2
                tmp = shgX.Range(Cells(1, 4), Cells(2 * numpeakX + 19, 6)).Copy
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
                tmp = shgX.Range(Cells(1, 1), Cells(10, 3))                  ' basic parameters
                en = shgX.Range(Cells(46, para + 11), Cells(47, para + 11)) ' database
                sheetGraph.Activate
                sheetGraph.Range(Cells(1, 1), Cells(10, 3)) = tmp
                sheetGraph.Range(Cells(46, para + 11), Cells(47, para + 11)) = en
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
                tmp = shfX.Range(Cells(1, 1), Cells(19 + sftfit2, 3))
                en = shfX.Range(Cells(2, 103), Cells(9, 103))
                sheetFit.Activate
                sheetFit.Range(Cells(1, 1), Cells(19 + sftfit2, 3)) = tmp
                sheetFit.Range(Cells(2, 103), Cells(9, 103)) = en
                
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
                
                tmp = shfX.Range(Cells(1, 5), Cells(15 + sftfit2 + 4, numpeakX + 4))
                sheetFit.Activate
                sheetFit.Range(Cells(1, 5), Cells(15 + sftfit2 + 4, numpeakX + 4)) = tmp
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


Function IntegrationTrapezoid(KnownXs As Variant, KnownYs As Variant) As Variant
    'Calculates the area under a curve using the trapezoidal rule.
    'KnownXs and KnownYs are the known (x,y) points of the curve.
    'By Christos Samaras
    'http://www.myengineeringworld.net
    
    Dim i As Integer
    
    'Check if the X values are range.
    If Not TypeName(KnownXs) = "Range" Then
        IntegrationTrapezoid = "Xs range is not valid"
        Exit Function
    End If
    
    'Check if the Y values are range.
    If Not TypeName(KnownYs) = "Range" Then
        IntegrationTrapezoid = "Ys range is not valid"
        Exit Function
    End If
    
    IntegrationTrapezoid = 0
    
    For i = 1 To KnownXs.Rows.Count - 1
        IntegrationTrapezoid = IntegrationTrapezoid + Abs(0.5 * (KnownXs.Cells(i + 1, 1) _
        - KnownXs.Cells(i, 1)) * (KnownYs.Cells(i, 1) + KnownYs.Cells(i + 1, 1)))
    Next i
End Function

Sub SolverInstall1()
    On Error Resume Next
    Dim wb As Workbook
    Dim SolverPath As String
    
    Set wb = ActiveWorkbook
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
    '// Dana DeLouis
    Dim wb As Workbook
    
    On Error Resume Next
    ' Set a Reference to the workbook that will hold Solver
    Set wb = ActiveWorkbook
    
    With wb.VBProject.References
        .Remove.Item ("SOLVER")
    End With
    
    With AddIns("Solver Add-In")
        .Installed = False
        .Installed = True
        wb.VBProject.References.AddFromFile .FullName
    End With
    
    ' initialize Solver
    Application.Run "Solver.xlam!Solver.Solver2.Auto_open"
End Sub






