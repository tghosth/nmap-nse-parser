
' Select the data table in Excel before you run this.
Sub FixNewLines()

    Dim cel As Range
    Dim selectedRange As Range

    Set selectedRange = Application.Selection

    For Each cel In selectedRange.Cells
        cel.FormulaR1C1 = cel.FormulaR1C1
    Next cel

End Sub
