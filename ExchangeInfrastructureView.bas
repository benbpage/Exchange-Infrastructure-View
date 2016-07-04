Sub InfraView3()
'
' InfraView Macro v 0.4
' Questions? Twitter @PaulGaljan

Application.ScreenUpdating = False
On Error Resume Next
Application.DisplayAlerts = False
Sheets("Infrastructure View").Delete
Application.DisplayAlerts = True
On Error GoTo 0
Range("JBODEvaluation").Value = "No"

'
    Sheets("Input").Select
    Sheets.Add.Name = "Infrastructure View"
    
    Range("B2").FormulaR1C1 = "# Copies"
    Range("C1").FormulaR1C1 = "PDC"
    Range("D1").FormulaR1C1 = "SDC"
    Range("B3").FormulaR1C1 = "DB Read%"
    Range("B4").FormulaR1C1 = "GCCores"
    Range("B4").FormulaR1C1 = "GC Cores"
    Range("B5").FormulaR1C1 = "Backup Capacity ("
    Range("B5").FormulaR1C1 = "Backup Capacity (1 Copy)"
    Range("B6").FormulaR1C1 = "# Servers"
    Range("B7").FormulaR1C1 = "Cores"
    Range("B8").FormulaR1C1 = "Ram"
    Range("B9").FormulaR1C1 = "Capacity"
    Range("B10").FormulaR1C1 = "DB IO"
    Range("B11").FormulaR1C1 = "BDM IO"
    Range("B12").FormulaR1C1 = "# Vols"
    Range("B13").FormulaR1C1 = "Cores"
    Range("B14").FormulaR1C1 = "Ram"
    Range("B15").FormulaR1C1 = "Capacity"
    Range("B16").FormulaR1C1 = "DB IO"
    Range("B17").FormulaR1C1 = "BDM IO"
    Range("B18").FormulaR1C1 = "# Vols"
    Range("B19").FormulaR1C1 = "Cores (incl. GC)"
    Range("B20").FormulaR1C1 = "Ram"
    Range("B21").FormulaR1C1 = "Capacity"
    Range("B22").FormulaR1C1 = "DB IO"
    Range("B23").FormulaR1C1 = "BDM IO"
    Range("B24").FormulaR1C1 = "# Vols"
 '   Range("C7").Select
    Columns("B:B").EntireColumn.AutoFit
    Range("A7:A12").Select
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Server"
    Range("A13:A18").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Site"
    Range("A19:A24").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "Site"
    Range("A13:A18").Select
    ActiveCell.FormulaR1C1 = "Copy"
    Range("A7:A24").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 90
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Columns("A:A").EntireColumn.AutoFit
    Selection.Font.Bold = True
    Range("C1:D1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    Range("B2:D6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A7:D12").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A13:D18").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A19:D24").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B2:D24").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Range("C1:D24").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Range("C13").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISERROR(C13)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("D13").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISERROR(D13)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("C19").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISERROR(C19)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("D19").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISERROR(D19)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("C5:D5").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.NumberFormat = "# ""TB"""
    Range("C9:D9,C15:D15,C21:D21").NumberFormat = "# ""TB"""
    Range("C8:D8,C14:D14,C20:D20").NumberFormat = "# ""GB"""
    Range("C2:D2,C4:D4,C6:D7,C10:D13,C16:D18,C19:D19,C22:D24").NumberFormat = "0"
    Range("C3:D3").NumberFormat = "0%"
    Range("C11:D11,C17:D17,C23:D23").NumberFormat = "# ""MB/s"""
    Columns("C:D").ColumnWidth = 10.71
    Range("A25").FormulaR1C1 = "Questions?  Twitter: @PaulGaljan"
        Range("F13").Select
    ActiveCell.FormulaR1C1 = "Sanity Check Data"
    Range("F14").Select
    ActiveCell.FormulaR1C1 = "Total Mailboxes"
    Range("F15").Select
    ActiveCell.FormulaR1C1 = "Avg Mailbox Size on Disk"
    Range("F16").Select
    ActiveCell.FormulaR1C1 = "Avg IO/Mbox"
    Range("F17").Select
    ActiveCell.FormulaR1C1 = "Mailboxes/Server"
    Range("F18").Select
    ActiveCell.FormulaR1C1 = "Mailboxes/DAG"
    Range("F19").Select
    ActiveCell.FormulaR1C1 = "Consider JBOD"
    Range("F13:G13").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
    Range("F13:G19").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("F13:G13").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Columns("F:F").EntireColumn.AutoFit
    Range("G15").NumberFormat = "#.0 ""GB"""
    Range("G16:G17").NumberFormat = "0"
    Range("G16").NumberFormat = "0.00"
    
' Finally Start with the formulas
    Range("E7").FormulaR1C1 = _
     "=IF(R[0]C[-2]=""--"",""<---Populate the SpecInt2006 Rate on the Input tab"","" "")"
    Range("C2").FormulaR1C1 = _
        "=IF(SRModel=""Active/Passive"",(NumDBCopies+numLagDBCopies)-(calcNumLagCopyInSDCActual+numDBCopiesSDC),((NumDBCopies+numLagDBCopies)/2))"
    Range("D2").FormulaR1C1 = _
        "=IF(SRModel=""Active/Passive"",(calcNumLagCopyInSDCActual+numDBCopiesSDC),((NumDBCopies+numLagDBCopies)/2))"
    Range("C3").FormulaR1C1 = "=aggRWRatio"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "=aggRWRatio"
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=IF(ValidationCheck=FALSE,calcRecGCCoresPDC,""--"")"
    Range("D4").FormulaR1C1 = "=IF(ValidationCheck=FALSE,calcRecGCCoresSDC,""--"")"
    Range("C5:D5").FormulaR1C1 = _
        "=((('Volume Requirements'!R[110]C[2]*'Role Requirements'!R[270]C)+('Role Requirements'!R[271]C*'Volume Requirements'!R[110]C[2]))/1024)*(R[1]C/R[-3]C)"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(SRModel=""Active/Passive"",(NumDAGServersPDC*NumDAGsEnv),((NumDAGServersPDC+NumDAGServersSDC)*NumDAGsEnv/2))"
    Range("D6").FormulaR1C1 = _
        "=IF(SRModel=""Active/Passive"",(NumDAGServersSDC*NumDAGsEnv),(((NumDAGServersPDC+NumDAGServersSDC)*NumDAGsEnv)/2))"
    Range("C7").FormulaR1C1 = _
        "=IF(AND(ValidationCheck=FALSE,SiteResilienceEnabled=""Yes"",numMCyclesPerCorePDC<>0),ROUNDUP(calcReqMBXCoresPDCServer+IF(calcMultiRoleEnabled=""Yes"",calcReqCASCoresPDCServer,0),0),""--"")"
    Range("D7").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(AND(ValidationCheck=FALSE,SiteResilienceEnabled=""Yes"",numMCyclesPerCoreSDC<>0),ROUNDUP(calcReqMBXCoresSDCServer+IF(calcMultiRoleEnabled=""Yes"",calcReqCASCoresSDCServer,0),0),""--"")"
    Range("C8").FormulaR1C1 = "=RecRAMMBXPDC"
    Range("D8").FormulaR1C1 = "=RecRAMMBXSDC"
    Range("C9").FormulaR1C1 = _
        "=(DBVolDiskSpaceReplicaSS+ResVolDiskSpaceNodeSS+TotLogVolSpace)/1024"
    Range("D9").FormulaR1C1 = _
        "=(DBVolDiskSpaceReplicaSS+ResVolDiskSpaceNodeSS+TotLogVolSpace)/1024"
    Range("C10").FormulaR1C1 = "=DBIOPSReplicaSS"
    Range("D10").FormulaR1C1 = "=DBIOPSReplicaSS"
    Range("C11").FormulaR1C1 = "=TotNumDBCopiesServer"
    Range("D11").FormulaR1C1 = "=TotNumDBCopiesServer"
    Range("C12").FormulaR1C1 = "=TotExchVolumes"
    Range("D12").FormulaR1C1 = "=TotExchVolumes"
    Range("C13").FormulaR1C1 = "=(R[-6]C*R[-7]C)/R[-11]C"
    Range("D13").FormulaR1C1 = "=(R[-6]C*R[-7]C)/R[-11]C"
    Range("C14").FormulaR1C1 = "=R[-6]C*R[-8]C/R[-12]C"
    Range("D14").FormulaR1C1 = "=R[-6]C*R[-8]C/R[-12]C"
    Range("C15").FormulaR1C1 = "=R[6]C/R[-13]C"
    Range("D15").FormulaR1C1 = "=R[6]C/R[-13]C"
    Range("C16").FormulaR1C1 = "=R[6]C/R[-14]C"
    Range("D16").FormulaR1C1 = "=R[6]C/R[-14]C"
    Range("C17").FormulaR1C1 = "=R[7]C/R[-15]C"
    Range("D17").FormulaR1C1 = "=R[7]C/R[-15]C"
    Range("C18").FormulaR1C1 = "=R[5]C/R[-16]C"
    Range("D18").FormulaR1C1 = "=R[5]C/R[-16]C"
    Range("C19").FormulaR1C1 = "=(R[-12]C*R[-13]C)+R[-15]C"
    Range("D19").FormulaR1C1 = "=(R[-12]C*R[-13]C)+R[-15]C"
    Range("C20").FormulaR1C1 = "=R[-12]C*R[-14]C"
    Range("D20").FormulaR1C1 = "=R[-12]C*R[-14]C"
    Range("C21").FormulaR1C1 = "=R[-12]C*R[-15]C"
    Range("D21").FormulaR1C1 = "=R[-12]C*R[-15]C"
    Range("C22").FormulaR1C1 = "=R[-12]C*R[-16]C"
    Range("D22").FormulaR1C1 = "=R[-12]C*R[-16]C"
    Range("C23").FormulaR1C1 = "=R[-11]C*R[-17]C"
    Range("D23").FormulaR1C1 = "=R[-11]C*R[-17]C"
    Range("C24").FormulaR1C1 = "=R[-13]C*R[-18]C"
    Range("D24").FormulaR1C1 = "=R[-13]C*R[-18]C"
    Range("G14").FormulaR1C1 = "=TotalMBX"
    Range("G15").FormulaR1C1 = "=(R[-10]C[-4]*1024)/R[-1]C"
    Range("G16").FormulaR1C1 = "=RC[-4]/R[-2]C"
    Range("G17").FormulaR1C1 = "=calcTotNumMBXPerSvr"
    Range("G18").FormulaR1C1 = "=TotMBXPerDAG"
    Range("G19").FormulaR1C1 = "=JBODEvaluation"
    Range("H19").FormulaR1C1 = _
        "=IF(JBODEvaluation=""Yes"",""<---Turn off JBOD Evaluation if deploying on SAN for more accurate size and IO estimation"","" "")"

    Range("A1").Select

End Sub
