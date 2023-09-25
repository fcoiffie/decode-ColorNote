' author: Michael Biggs
' created: 2022-12-08
' description: LibreOffice Basic macro to highlight rows as per
'              ColorNote colors.


' A. Using LibreOffice Calc to format csv data:
'     1.  open '.csv' document file in LibreOffice Calc
'     2.  manually select entire cell range containing dates
'         Format->Cells...->Category->Date->Format Code
'         type 'YYYY-MM-DD HH:MM:SS' then click 'OK'
'     3.  manually select entire populated cell range including headers
'         Data->AutoFilter
'     4.  with entire populated cell range still selected
'         Format->Text->Wrap Text
'     5.  to adjust column widths:
'             select one or more populated columns, then
'         either auto adjust using:
'             Format->Columns->Optimal Width...->OK
'         or manually adjust using:
'             Format->Columns->Width...
'                 type in number, say '7.5' (inches) then click 'OK'
'     6.  to adjust row heights:
'              select one or more populated rows, then
'         either auto adjust using:
'             Format->Rows->Optimal Height...->OK
'         or manually adjust using:
'             Format->Rows->Height...
'                 type in number, say '1.5' (inches) then click 'OK'
'     7.  save document file as '.ods'
'


' B. To enable macros in LibreOffice Calc - will prompt on opening:
'     1.  Tools->Options...->Security->Macro Security...->
'             select 'Medium'
'     2.  OK->OK
'


' C. To add 'ColorNoteRows' macro to LibreOffice Calc document:
'     1.  open '.bas' macro file in a text editor
'         Edit->Select All
'         Edit->Copy
'         close '.bas' macro file and text editor
'     2.  open '.csv', '.ods' or '.xlsx' document file in LibreOffice Calc
'         Tools->Macros->Organize Macros->Basic...->Macros From
'             click on file name
'             click 'New' to open 'Module1' in LibreOffice Basic editor
'         Edit->Select All
'         Edit->Cut
'         Edit->Paste
'         File->Close
'     3.  right-click on sheet name
'         Sheet Events...->Macro...
'             click '+' next to file name to reveal 'Standard'
'             click '+' next to 'Standard' to reveal 'Module1'
'             click on 'Module1' then macro name that appears opposite
'     4.  OK->OK
'     5.  click '+' next to sheet name (to add another sheet)
'         click on original sheet name (to activate sheet and run macro)
'     6.  save document file as '.ods'


'-------------------------------------------------------------------------------
Public Sub ColorNoteRows()
    ' Color each row based on value of cell in color_index column
    '     - mimics colors used by ColorNote
    '     - requires correct value for color_index_table_column
    '       where first column of table = 1, etc
    Dim used_range As Object
    Dim first_row As Long
    Dim first_column As Long
    Dim last_row As Long
    Dim last_column As Long
    Dim color_index_column As Long
    Dim color_index_column_offset As Long
    Dim i As Long

    color_index_table_column = 2 ' MAKE SURE THIS IS CORRECT

    sheet = ThisComponent.CurrentController.ActiveSheet

    cursor = sheet.createCursor()
    cursor.gotoStartOfUsedArea(False)   ' move to cell at start of used area
    cursor.gotoEndOfUsedArea(True)      ' expand to end of used area
    used_range = cursor.RangeAddress    ' cursor.RangeAddress is the used range
    first_row = used_range.StartRow
    first_column = used_range.StartColumn
    last_row = used_range.EndRow
    last_column = used_range.EndColumn

    color_index_column = first_column + color_index_table_column - 1

    ' getCellByPosition(X, Y) ie (col, row)
    ' getCellRangeByPosition(X1 Y1, X2, Y2) ie (col1, row1, col2, row2)
    For i = first_row To last_row
        With sheet.getCellRangeByPosition(first_column, i, last_column, i)
            Select Case sheet.getCellByPosition(color_index_column, i).Value
                Case 1
                    .CellBackColor = RGB(255, 230, 233) ' red
                Case 2
                    .CellBackColor = RGB(255, 235, 216) ' orange
                Case 3
                    .CellBackColor = RGB(254, 248, 186) ' yellow
                Case 4
                    .CellBackColor = RGB(229, 248, 220) ' green
                Case 5
                    .CellBackColor = RGB(232, 233, 254) ' blue
                Case 6
                    .CellBackColor = RGB(239, 224, 255) ' purple
                Case 7
                    .CellBackColor = RGB(204, 204, 204) ' black
                Case 8
                    .CellBackColor = RGB(238, 238, 238) ' grey
                Case 9
                    .CellBackColor = RGB(255, 255, 255) ' white
                Case Else
                    .CellBackColor = -1
            End Select
        End With
    Next i
End Sub
