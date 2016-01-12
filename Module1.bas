Attribute VB_Name = "Module1"
Option Explicit

' This section turns on high overhead operations
Sub AppTrue()
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

' This section increases performance by turning off high overhead operations
Sub AppFalse()
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

End Sub

Sub BuildAllocatedTable(Table As ListObject, AllocatedTableName As String, AllocatedTableSheet As String, ApplyInflation As Boolean)
    
    Call AppFalse
    Sheets(AllocatedTableSheet).Activate
    Dim EntryCollection As EntryCollection
    Set EntryCollection = New EntryCollection
    Dim Entry As Entry
    Dim ListRow As Excel.ListRow
    Dim RowNum As Long
    Dim Column As Integer
    Dim i As Integer
    Dim j As Long
    Dim ServiceArray() As Variant
    Dim PALSOSArray() As Variant
    Dim RowRange As Range
    Dim Row As ListRow
    Dim AllocatedTableRowCount As Long
    Dim AllocatedTableStartColumn As Integer
    Dim AllocatedTableStartRow As Integer
    Dim AllocatedNumberofColumns As Integer
    Dim ServiceShareRule As Range
    Dim PALS_OS_Rule As Range
    Dim SplitsServiceShareRule As Range
    Dim SplitsPALS_OS_Rule As Range

    'fill entry data from the heads table (ID, PALS/O&S array, Service array)
    For Each ListRow In Table.ListRows
        Set Entry = New Entry
        Entry.ID = Intersect(Table.ListColumns("ID").Range, ListRow.Range).value
        i = 0
        j = 0
        
        Set ServiceShareRule = Intersect(ListRow.Range, Table.ListColumns("Service Share Rule").Range)
        Set PALS_OS_Rule = Intersect(ListRow.Range, Table.ListColumns("PALS/O&S Split").Range)
        
        ReDim ServiceArray(WorksheetFunction.CountIf([SplitTable[Split Name]], Intersect(ListRow.Range, Table.ListColumns("Service Share Rule").DataBodyRange).value))
        ReDim PALSOSArray(WorksheetFunction.CountIf([SplitTable[Split Name]], Intersect(ListRow.Range, Table.ListColumns("PALS/O&S Split").DataBodyRange).value))
        
        'loop through Split table to fill arrays
        For Each Row In [SplitTable].ListObject.ListRows

            If Sheets("Splits").Cells(Intersect(Row.Range, [SplitTable[Split Name]]).Row, Intersect(Row.Range, [SplitTable[Split Name]]).Column) = ServiceShareRule And ServiceShareRule <> "" Then
                ServiceArray(i) = Sheets("Splits").Cells(Intersect(Row.Range, [SplitTable[Service]]).Row, Intersect(Row.Range, [SplitTable[Service]]).Column)
                i = i + 1
            ElseIf Sheets("Splits").Cells(Intersect(Row.Range, [SplitTable[Split Name]]).Row, Intersect(Row.Range, [SplitTable[Split Name]]).Column) = PALS_OS_Rule And PALS_OS_Rule <> "" Then
                PALSOSArray(j) = Sheets("Splits").Cells(Intersect(Row.Range, [SplitTable[PALS/O&S]]).Row, Intersect(Row.Range, [SplitTable[PALS/O&S]]).Column)
                j = j + 1
            End If
        Next Row
        
        'assign arrays to object property and add entry object to collection
        Entry.Service = ServiceArray
        Entry.PALSOS = PALSOSArray
        EntryCollection.Entries.Add Entry
    Next ListRow
    
    'determine needed size for Allocated table
    AllocatedTableRowCount = 0
    For Each Entry In EntryCollection.Entries
        AllocatedTableRowCount = AllocatedTableRowCount + (UBound(Entry.PALSOS) * UBound(Entry.Service))
    Next Entry
    
    Dim ListObject As ListObject
    Dim AllocatedTable As ListObject
    For Each ListObject In Sheets(AllocatedTableSheet).ListObjects
        If ListObject.Name = AllocatedTableName Then
            Set AllocatedTable = ListObject
        End If
    Next ListObject
    
    'determine Absolute row (sheet row, instead of listobject row) of first row in table
    Dim AbsRow As Long
    AbsRow = AllocatedTable.HeaderRowRange.Row + 1
    
    'delete all table rows
    If AllocatedTable.ListRows.Count > 0 Then
        'clear table, add one row, get row value
        AllocatedTable.DataBodyRange.Rows.Delete
    End If
    
    'assign number values of header row, first table column, number of column
    AllocatedTableStartRow = AllocatedTable.HeaderRowRange.Row
    AllocatedTableStartColumn = AllocatedTable.HeaderRowRange.Column
    AllocatedNumberofColumns = AllocatedTable.HeaderRowRange.Columns.Count
    
    'dimension field column variables
    Dim Allocated_IDCol As Integer
    Dim PALS_OSCol As Integer
    Dim ServiceCol As Integer
    Dim IDCol As Integer
    
    'assign column values to variables
    IDCol = AllocatedTable.ListColumns("ID").Range.Column
    Allocated_IDCol = AllocatedTable.ListColumns(Table.Name & "_ID").Range.Column
    PALS_OSCol = AllocatedTable.ListColumns("PALS/O&S").Range.Column
    ServiceCol = AllocatedTable.ListColumns("Service").Range.Column
    
    'convert table to range because filling cells in a range is MUCH faster than in a table
    AllocatedTable.Unlist
       
    Dim k As Long
    k = 1
    'fill ID, PALS/O&S, and Service columns
    For Each Entry In EntryCollection.Entries
        For i = 1 To UBound(Entry.PALSOS)
            For j = 1 To UBound(Entry.Service)
            Cells(AbsRow, Allocated_IDCol) = Entry.ID
            Cells(AbsRow, PALS_OSCol) = Entry.PALSOS(i - 1)
            Cells(AbsRow, ServiceCol) = Entry.Service(j - 1)
            Cells(AbsRow, IDCol) = k
            AbsRow = AbsRow + 1
            k = k + 1
            Next j
        Next i
    Next Entry
    
    Dim AllocatedTableWorksheet As Worksheet
    Set AllocatedTableWorksheet = Sheets(AllocatedTableSheet)
    'convert back to table
    Call ConvertToTable(AllocatedTableWorksheet, Range(Cells(AllocatedTableStartRow, AllocatedTableStartColumn), Cells(AllocatedTableStartRow + AllocatedTableRowCount, AllocatedTableStartColumn + AllocatedNumberofColumns - 1)), AllocatedTableName)
    
    Dim YearFormula As String
    Dim IndexTableFormula As String
    Dim FromYearFormula As String
    YearFormula = "LOOKUP(" & AllocatedTableName & "[@[" & Table.Name & "_ID]:[" & Table.Name & "_ID]]," & Table.Name & "[[ID]:[ID]]," & Table.Name & "[2009])*SUMPRODUCT((" & AllocatedTableName & "[@[PALS/O&S Split]:[PALS/O&S Split]] = SplitTable[[Split Name]:[Split Name]])*(" & AllocatedTableName & "[@[PALS/O&S]:[PALS/O&S]] = SplitTable[[PALS/O&S]:[PALS/O&S]])*SplitTable[2009])*SUMPRODUCT((" & AllocatedTableName & "[@[Service Share Rule]:[Service Share Rule]]=SplitTable[[Split Name]:[Split Name]])*(" & AllocatedTableName & "[@[Service]:[Service]]=SplitTable[[Service]:[Service]])*SplitTable[2009])"
    
    If ApplyInflation Then
        IndexTableFormula = "IF([@[PALS/O&S]:[PALS/O&S]]=" & Chr(34) & "PALS" & Chr(34) & ",[@[PALS Table]:[PALS Table]],[@[O&S Table]:[O&S Table]])"
        FromYearFormula = "IF([@[From Type]:[From Type]]=" & Chr(34) & "TY" & Chr(34) & "," & AllocatedTableName & "[[#Headers],[2009]],[@[From Year]:[From Year]])"
        YearFormula = "=inflation(" & YearFormula & ",[@[Index Year]:[Index Year]]," & IndexTableFormula & "," & FromYearFormula & ",[@[From Type]:[From Type]],[@[To Year]:[To Year]],[@[ToType]:[ToType]])"
    Else
        YearFormula = "=" & YearFormula
    End If
    
    Call TransferTableData(Table, Sheets(AllocatedTableSheet).ListObjects(AllocatedTableName), YearFormula)
    Call RemoveFormulas(Sheets(AllocatedTableSheet).ListObjects(AllocatedTableName))
    Call AppTrue
    
End Sub

Sub BuildCostTable()
    
    Call AppFalse
    [CostTable].ListObject.Parent.Activate
    'clear table
    If [CostTable].ListObject.ListRows.Count > 0 Then
        [CostTable].ListObject.DataBodyRange.Rows.Delete
    End If
    
    'determine number of rows for CostTable
    Dim NumRows As Long
    NumRows = [AllocatedHeads].ListObject.DataBodyRange.Rows.Count
    
    'determine number of columns for CostTable
    Dim NumColumns As Long
    NumColumns = [AllocatedHeads].ListObject.DataBodyRange.Columns.Count
    
    'find CostTable starting location
    Dim FirstRow As Integer
    Dim FirstColumn As Integer
    FirstRow = [CostTable].ListObject.HeaderRowRange.Row
    FirstColumn = [CostTable].ListObject.HeaderRowRange.Column
    
    'fill AllocatedHeadsID, ID
    Dim Row As Variant
    Dim CurrentCell As Range
    Dim IDCell As Range
    Dim i As Long
    Dim IDColumn As Integer
    Dim Allocated_IDColumn As Integer
    IDColumn = [CostTable].ListObject.HeaderRowRange.Find("ID", , , xlWhole).Column
    Allocated_IDColumn = [CostTable].ListObject.HeaderRowRange.Find("AllocatedHeads_ID", , , xlWhole).Column
    i = 1
    
    'convert table to range (much faster to fill data)
    [CostTable].ListObject.Unlist
    
    For i = 1 To NumRows
        Set CurrentCell = Cells(FirstRow + i, IDColumn)
        Set IDCell = Cells(FirstRow + i, Allocated_IDColumn)
        CurrentCell = i
        IDCell = i
    Next
    
    'convert back to table
    Call ConvertToTable(Sheets("Cost"), Range(Cells(FirstRow, FirstColumn), Cells(FirstRow + NumRows, FirstColumn + NumColumns - 1)), "CostTable")
    
    'resize CostTable
    [CostTable].ListObject.Resize Range(Cells(FirstRow, FirstColumn), Cells(FirstRow + NumRows, FirstColumn + NumColumns - 1))
    
    'fill Array with field names of data to transfer between tables
    'assumption: all fields to the left of total transfer except all ID fields
    Dim YearFormula As String
    YearFormula = "=LOOKUP(CostTable[@[AllocatedHeads_ID]:[AllocatedHeads_ID]],AllocatedHeads[[ID]:[ID]],AllocatedHeads[2009])*SUMPRODUCT((CostTable[@[Rate]:[Rate]] = RatesTable[[Description]:[Description]])*(CostTable[@[PALS/O&S]:[PALS/O&S]] = RatesTable[[PALS/O&S]:[PALS/O&S]])*RatesTable[2009])"
            
    Call TransferTableData([AllocatedHeads].ListObject, [CostTable].ListObject, YearFormula)
        
    Call AppendTable([AllocatedDollarInputs].ListObject, [CostTable].ListObject)
    Call RemoveFormulas([CostTable].ListObject)
    Call AppTrue
End Sub

Sub TransferTableData(FromTable As ListObject, ToTable As ListObject, YearFormula As String)
    
    Dim Cell As Range
    Dim j As Integer
    j = 0
    Dim FieldNameArray() As String
    For Each Cell In FromTable.HeaderRowRange
        If Cell = "Total" Then
            Exit For
        End If
        
        If InStr(1, Cell, "ID") = 0 Then
            ReDim Preserve FieldNameArray(j)
            FieldNameArray(j) = Cell
            j = j + 1
        End If
    Next Cell
    
    Dim FieldName As Variant
    'resume next to prevent errors caused by fields in FromTable not being present in ToTable
    On Error Resume Next
    For Each FieldName In FieldNameArray
        If ToTable.ListColumns(FieldName).DataBodyRange.Formula(1, 1) = "" Then
            ToTable.ListColumns(FieldName).DataBodyRange = "=LOOKUP(" & ToTable.Name & "[@[" & FromTable.Name & "_ID]:[" & FromTable.Name & "_ID]]," & FromTable.Name & "[[ID]:[ID]]," & FromTable.Name & "[[" & FieldName & "]:[" & FieldName & "]])"
        End If
    Next FieldName
    'turn of error handling so other errors still come up
    On Error GoTo 0
    
    ToTable.ListColumns("2009").DataBodyRange = YearFormula
    
    Dim FirstCell As Range
    Dim FillRange As Range
    Set FirstCell = ToTable.HeaderRowRange.Find("2009")
    Set FirstCell = Cells(FirstCell.Row + 1, FirstCell.Column)
    'Could not paste inflation into years past 2075 because it creates a value error (divide by 0) which for some reason
    'could not be bypassed by a OnError Resume Next Statement
    Set FillRange = Range(FirstCell.Address, Cells(FirstCell.Row, ToTable.Range.SpecialCells(xlLastCell).Column - 5))
    FirstCell.AutoFill FillRange, xlFillDefault
    FillRange.Copy
    FirstCell.PasteSpecial xlPasteFormulas
    'create calculated column in Total column
    ToTable.ListColumns("Total").DataBodyRange = "=SUM(" & FirstCell.Address(False, False) & ":" & Cells(FirstCell.Row, ToTable.Range.SpecialCells(xlLastCell).Column).Address(False, False) & ")"
  
End Sub

Sub RemoveFormulas(Table As ListObject)
   
    'convert to values, formulas make the table slow and the file size bigger.
    Table.DataBodyRange.Copy
    Range(Cells(Table.DataBodyRange.Row, Table.DataBodyRange.Column), Cells(Table.DataBodyRange.Row, Table.DataBodyRange.Column)).PasteSpecial xlPasteValues
       
End Sub

Sub AppendTable(Table As ListObject, ToTable As ListObject)
    
    Dim Cell As Range
    Dim Field As Field
    Dim Values() As Variant
    Dim FieldCollection As FieldCollection
    Set FieldCollection = New FieldCollection
    Dim NumRowsTable As Long
    Dim NumRowsToTable As Long
    Dim TableSheet As Worksheet
    Dim ToTableSheet As Worksheet
    
    Set TableSheet = Table.Parent
    Set ToTableSheet = ToTable.Parent
    
    NumRowsTable = Table.ListRows.Count
    NumRowsToTable = ToTable.ListRows.Count
    
    For Each Cell In Table.HeaderRowRange
        
        TableSheet.Activate
        Values() = TableSheet.Range(Cells(Cell.Row + 1, Cell.Column), Cells(Cell.Row + NumRowsTable, Cell.Column))

        Set Field = New Field
        Field.Header = Cell
        Field.FieldValues = Values()
        FieldCollection.Fields.Add Field
        
    Next Cell
    
    ToTableSheet.Activate
    Dim FirstCell As Range
    Set FirstCell = Cells(ToTable.HeaderRowRange.Row, ToTable.HeaderRowRange.Column)
    Dim LastCell As Range
    Set LastCell = Cells(FirstCell.Row + Table.ListRows.Count + ToTable.ListRows.Count, FirstCell.Column + ToTable.HeaderRowRange.Columns.Count)
    
    Dim ToHeaderRange As Range
    Set ToHeaderRange = ToTable.HeaderRowRange
    Dim ToTableName As String
    ToTableName = ToTable.Name
    Dim ToTableRange As Range
    
    Set ToTableRange = ToTableSheet.Range(FirstCell, LastCell)
    
    ToTable.Unlist
    Dim AppendRange As Range
    Dim Cell2 As Range
    Dim i As Integer
    i = 1
    
    
    
    For Each Cell In ToHeaderRange
        Set AppendRange = ToTableSheet.Range(Cells(Cell.Row + NumRowsToTable + 1, Cell.Column), Cells(Cell.Row + NumRowsToTable + UBound(FieldCollection.Fields(1).FieldValues), Cell.Column))
        If Cell = "ID" Then
            For Each Cell2 In AppendRange
                Cell2 = NumRowsToTable + i
                i = i + 1
            Next Cell2
        End If
        
        For Each Field In FieldCollection.Fields
            Set AppendRange = ToTableSheet.Range(Cells(Cell.Row + NumRowsToTable + 1, Cell.Column), Cells(Cell.Row + NumRowsToTable + UBound(Field.FieldValues), Cell.Column))

            If Cell = Table.Name & "_ID" And Field.Header = "ID" Then
                AppendRange = Field.FieldValues
                Exit For
            ElseIf Field.Header = Cell And Cell <> "ID" Then
                AppendRange = Field.FieldValues
                Exit For
            End If
        Next Field
    Next Cell
    
    Call ConvertToTable(ToTableSheet, ToTableRange, ToTableName)
    
End Sub

Sub ConvertToTable(Sheet As Worksheet, Range As Range, TableName As String)
    'convert back to table
    With Sheet.ListObjects.Add(xlSrcRange, Range, , xlYes)
    .Name = TableName
    .TableStyle = "TableStyleMedium7"
    End With
    
End Sub

