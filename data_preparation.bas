Attribute VB_Name = "data_preparation"
Sub main()

    ' The main sub is the one that should be executed.
    ' You don't need to execute any other sub.

    DeleteShapes
    DeleteHeaders
    DeleteBlankLines
    LoopDeleteBlankCols
    CreateTable
    CreateFormulas
    CreateTotalLine
    DeactivateTableFilter
    AlignTableHeader
    AdjustColsWidth
    
    Range("A1").Activate
    
End Sub

Sub AlignTableHeader()

    ' Center align every content of the table header.

    ActiveSheet.ListObjects(1).HeaderRowRange.Select
    
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
    
    Range("A1").Activate

End Sub

Sub DeactivateTableFilter()

    ' Deactivate the filter buttons of the table.
    
    ActiveSheet.ListObjects(1).ShowAutoFilterDropDown = False

End Sub

Sub AdjustColsWidth()

    ' Adjust the width of every column.

    Cells.Select
    Cells.EntireColumn.AutoFit

End Sub

Sub CreateTable()

    ' Create table.

    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(GetLastRowNum(6000), GetUltColNum(1000))), , xlYes).Name = "tab_data"

End Sub

Sub CreateFormulas()

    ' Create formulas in the last two columns ("TOTAL HE" and "TOTAL VHE").

    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects(1)
    
    tbl.ListColumns("TOTAL HE").DataBodyRange.Select
    Selection.FormulaR1C1 = "=[@[HE 50]] + [@[HE 100]] + [@[HE 150]]"
    
    tbl.ListColumns("TOTAL VHE").DataBodyRange.Select
    Selection.FormulaR1C1 = "=[@[VHE 50]] + [@[VHE 100]] + [@[VHE 150]]"

End Sub

Sub CreateTotalLine()
    
    ' Create total formulas in the total line for each column
    ' of the resulting table.
    
    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects(1)
    
    tbl.ShowTotals = True
            
    Dim stringForm As String
    
    For Each CL In tbl.HeaderRowRange
        stringForm = "=SUBTOTAL(109,["
        
        If CL.Value <> "ID" And CL.Value <> "Nome" Then
            stringForm = stringForm & CL.Value & "])"
            tbl.TotalsRowRange(tbl.ListColumns(CL.Value).Index).Select
            Selection.FormulaR1C1 = stringForm
        End If
    Next CL
    
End Sub

Sub DeleteShapes()
    
    ' Delete every shape (including every image)
    ' from the active sheet.
    
    For Each oPic In ActiveSheet.Shapes
        oPic.Delete
    Next oPic

End Sub

Sub DeleteHeaders()

    ' Clear the contents of every header but the first header.
    
    Range("A1").Activate
    
    Do While ActiveCell.Value = ""
        ActiveCell.Offset(1, 0).Activate
    Loop
    
    Dim firstHeader As Range
    Set firstHeader = Range(ActiveCell.Address)

    Dim lastRow As Range
    Set lastRow = GetLastRow()
    
    For Each CL In Range("A1", lastRow)
        
        CL.Activate
        
        If CL.Value = "ID" And CL.Address <> firstHeader.Address Then
            
            Rows(ActiveCell.Row).Select
            Selection.ClearContents
            
        End If
        
    Next CL

End Sub

Sub DeleteBlankLines()
    
    ' Delete every blank row.
    
    Dim lastRow As Range
    Set lastRow = GetLastRow()

    Dim lastRowNum As Integer
    lastRowNum = lastRow.Row

    Range("A1").Activate

    Do While ActiveCell.Value = ""
        ActiveCell.Offset(1, 0).Activate
    Loop

    Do While ActiveCell.Row < lastRowNum
        If ActiveCell.Value = "" Then
            ActiveCell.EntireRow.Delete
            lastRowNum = lastRowNum - 1
        Else
            ActiveCell.Offset(1, 0).Activate
        End If
    Loop
    
    ' Delete the blank rows above the first header.
    Range("A1").Activate
    
    If Range("A1").Value <> "ID" Then
        Do While ActiveCell.Value = ""
            ActiveCell.Offset(1, 0).Activate
        Loop
    End If
    
    If ActiveCell.Value = "ID" And ActiveCell.Row <> 1 Then
        Range("A1", ActiveCell.Offset(-1, 0)).Select
        Selection.EntireRow.Delete
    End If
    
    Range("A1").Activate

End Sub

Sub DeleteBlankCols()

    ' Delete blank cols between contents.

    For col = 1 To GetUltColNum(1000)
    
        If Cells(1, col).Value = "" And col < GetUltColNum(1000) Then
            Columns(Cells(1, col).Column).EntireColumn.Delete
        End If
    
    Next col
    
    Range("A1").Select

End Sub

Sub LoopDeleteBlankCols()
       
    ' Execute the DeleteBlankCols for the necessary number of times
    ' until there is no more blank columns.
       
    Do
        If CountBlankCols() > 0 Then
            DeleteBlankCols
        End If
    
        blankCols = CountBlankCols()
        
    Loop While blankCols > 0

End Sub

' -----------------------------------------------------------------
' AUXILIARY FUNCTIONS
' -----------------------------------------------------------------

Public Function CountBlankCols() As Integer

    ' Count blank columns and return this number as an Integer.

    blankCols = 0

    For Each CL In Range(Cells(1, 1), Cells(1, GetUltColNum(1000)))
        If CL.Value = "" Then
            blankCols = blankCols + 1
        End If
    Next CL

    CountBlankCols = blankCols

End Function

Public Function GetLastRow() As Range

    ' Return the last row as a Range object.

    Range("A60000").Select
    Selection.End(xlUp).Select
    
    Set GetLastRow = ActiveCell

End Function

Public Function GetLastRowNum(limRow As Integer) As Integer

    ' Return the number of the last row as an Integer.

    Cells(limRow, 1).Select
    Selection.End(xlUp).Select
    
    GetLastRowNum = ActiveCell.Row

End Function

Public Function GetUltColNum(limCol As Integer) As Integer

    ' Return the number of the last column as an Integer.

    Cells(1, limCol).Select
    Selection.End(xlToLeft).Select
    
    Dim lastCol As Range
    Set lastCol = Selection
    
    GetUltColNum = lastCol.Column

End Function
