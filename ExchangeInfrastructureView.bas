Attribute VB_Name = "Module2"
Sub InfraView2()
'
' InfraView2 Macro v 0.1
' Questions? Twitter @PaulGaljan

'
'Startup and Cleanup
Application.ScreenUpdating = False
On Error Resume Next
Application.DisplayAlerts = False
Sheets("Infrastructure View").Delete
Application.DisplayAlerts = True
On Error GoTo 0


'Change JBODEvaluation to No
Range("JBODEvaluation").Value = "No"

'Add headings
    Sheets("Input").Select
    Sheets.Add.Name = "Infrastructure View"
    Columns("A:A").ColumnWidth = 3.14
    Range("C1").FormulaR1C1 = "Site 1"
    Range("D1").FormulaR1C1 = "Site 2"
    Range("B2").FormulaR1C1 = "# Copies"
    Range("B3").FormulaR1C1 = "DB Read %"
    Range("B4").FormulaR1C1 = "# Servers"
    Range("B5").FormulaR1C1 = "CPU Cores"
    Range("B6").FormulaR1C1 = "RAM"
    Range("B7").FormulaR1C1 = "Storage Capacity"
    Range("B8").FormulaR1C1 = "DB IO"
    Range("B9").FormulaR1C1 = "BDM IO"
'Calculations
    Range("C2").FormulaR1C1 = "=(NumDBCopies+numLagDBCopies)-(calcNumLagCopyInSDCActual+numDBCopiesSDC)"
    Range("D2").FormulaR1C1 = "=(calcNumLagCopyInSDCActual+numDBCopiesSDC)"
    Range("C3").FormulaR1C1 = "=aggRWRatio"
    Range("D3").FormulaR1C1 = "=aggRWRatio"
    Range("C4").FormulaR1C1 = "=NumDAGServersPDC*NumDAGsEnv"
    Range("D4").FormulaR1C1 = "=NumDAGServersSDC*NumDAGsEnv"
    Range("D5").FormulaR1C1 = _
        "=IF(AND(ValidationCheck=FALSE,SiteResilienceEnabled=""Yes"",numMCyclesPerCoreSDC<>0),ROUNDUP(calcReqMBXCoresSDCServer+IF(calcMultiRoleEnabled=""Yes"",calcReqCASCoresSDCServer,0),0),""--"")"
    Range("C5").FormulaR1C1 = _
        "=IF(AND(ValidationCheck=FALSE,SiteResilienceEnabled=""Yes"",numMCyclesPerCorePDC<>0),ROUNDUP(calcReqMBXCoresPDCServer+IF(calcMultiRoleEnabled=""Yes"",calcReqCASCoresPDCServer,0),0),""--"")"
    Range("C6").FormulaR1C1 = "=RecRAMMBXPDC"
    Range("D6").FormulaR1C1 = "=RecRAMMBXSDC"
    Range("C7").FormulaR1C1 = "=(DBVolDiskSpaceReplicaSS+ResVolDiskSpaceNodeSS)/1024"
    Range("D7").FormulaR1C1 = "=(DBVolDiskSpaceReplicaSS+ResVolDiskSpaceNodeSS)/1024"
    Range("C8").FormulaR1C1 = "=DBIOPSReplicaSS"
    Range("D8").FormulaR1C1 = "=DBIOPSReplicaSS"
    Range("C9").FormulaR1C1 = "=TotNumDBCopiesServer"
    Range("D9").FormulaR1C1 = "=TotNumDBCopiesServer"
'Extrapolations
    Range("B5:B9").Select
    Range("B9").Activate
    Selection.Copy
    Range("B10").Select
    ActiveSheet.Paste
    Range("B15").Select
    ActiveSheet.Paste
    Range("A5:A9").Select
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

'Apply initial formatting
    ActiveCell.FormulaR1C1 = "Server"
    Range("A10:A14").Select
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
    ActiveCell.FormulaR1C1 = "Copy"
    Range("A15:A19").Select
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
    Range("A5:A19").Select
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
    Selection.Font.Bold = True
    Columns("B:B").ColumnWidth = 15.14
    Range("A5:D9").Select
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
    Range("A10:D14").Select
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
    Range("A15:D19").Select
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
    Range("C1:D19").Select
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
    Range("B2:D4").Select
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
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
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
    Range("A5:A19").Select
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
    Range("C10").Select
    ActiveCell.FormulaR1C1 = "=(R[-5]C*R[-6]C)/R[-8]C"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = "=R[-5]C*R[-7]C/R[-9]C"
    Range("C12").Select
    ActiveCell.FormulaR1C1 = "=R[5]C/R[-10]C"
    Range("C13").Select
    ActiveCell.FormulaR1C1 = "=R[5]C/R[-11]C"
    Range("C14").Select
    ActiveCell.FormulaR1C1 = "=R[5]C/R[-12]C"
    Range("C15").Select
    ActiveCell.FormulaR1C1 = "=R[-10]C*R[-11]C"
    Range("C16").Select
    ActiveCell.FormulaR1C1 = "=R[-10]C*R[-12]C"
    Range("C17").Select
    ActiveCell.FormulaR1C1 = "=R[-10]C*R[-13]C"
    Range("C18").Select
    ActiveCell.FormulaR1C1 = "=R[-10]C*R[-14]C"
    Range("C19").Select
    ActiveCell.FormulaR1C1 = "=R[-10]C*R[-15]C"
    Range("C10:C19").Select
    Selection.Copy
    Range("D10").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("C3:D3").Style = "Percent"

' Conditional Formatting to handle lack of proc information
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISERROR(C10)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("D10").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISERROR(D10)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("D15").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISERROR(D15)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("C15").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISERROR(C15)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Range("C10").Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ISERROR(C10)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

'Comment for specint rates
    Range("C5").AddComment
    Range("C5").comment.Visible = False
    Range("C5").comment.Text Text:= _
        "pgaljan:" & Chr(10) & "Enter SpecInt2006 Rate values on the Input tab to calculate cores."
    Range("C2").Select

' Format with units
    Range("C7:D7,C12:D12,C17:D17").Select
    Range("C17").Activate
    Selection.NumberFormat = "#.0 ""TB"""
    Range("C9:D9,C14:D14,C19:D19").Select
    Range("C19").Activate
    Selection.ClearComments
    Selection.NumberFormat = "# ""MB/s"""
'Make it pretty
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 1.71
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("D9:E9").Select
    Selection.NumberFormat = "0"
Range("G1").FormulaR1C1 = "Questions?  Twitter: @PaulGaljan"
Range("A1").Activate
End Sub


