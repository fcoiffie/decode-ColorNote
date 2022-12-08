' author: Michael Biggs
' created: 2022-12-08
' description: Microsoft Visual Basic macro to highlight rows as per
'              ColorNote colors.


' A. Using Microsoft Excel to format csv data:
'     1.  open '.csv' document file in Microsoft Excel
'     2.  manually select entire cell range containing dates
'         HOME->Cells->Format->Format Cells->Custom->Type
'         type 'yyyy-mm-dd hh:mm:ss' then click 'OK'
'     3.  manually select entire populated cell range including headers
'         INSERT->Table
'         check checkbox 'My table has headers' then click 'OK'
'     4.  with entire populated cell range still selected
'         HOME->Alignment->Wrap Text
'     5.  to adjust column widths:
'             select one or more populated columns, then
'         either auto adjust using:
'             HOME->Cells->Format->Autofit Column Width
'         or manually adjust using:
'             HOME->Cells->Format->Column Width...->Column width
'                 type in number, say '75' then click 'OK'
'     6.  to adjust row heights:
'              select one or more populated rows, then
'         either auto adjust using:
'             HOME->Cells->Format->Autofit Row Height
'         or manually adjust using:
'             HOME->Cells->Format->Row Height...->Row height
'                 type in number, say '15' then click 'OK'
'     7.  save document file as '.xlsx'
'


' B. To enable macros in Microsoft Excel - will prompt on opening:
'     1.  File->Options->Trust Center->Trust Center Settings...->Macro Settings->
'             select 'Disable all macros with notification'
'     2.  OK->OK
'


' C. To add 'ColorNoteRows' macro to Microsoft Excel document:
'     1.  open '.vb' macro file in a text editor
'         Edit->Select All
'         Edit->Copy
'         close '.vb' macro file and text editor
'     2.  open '.csv' or '.xlsx' document file in Microsoft Excel
'         right-click on sheet name
'         click View Code to open Microsoft Visual Basic editor
'         Edit->Paste
'         File->Close and Return to Microsoft Excel
'     3.  click '+' next to sheet name (to add another sheet)
'         click on original sheet name (to activate sheet and run macro)
'     4.  save document file as '.xlsm'


'-------------------------------------------------------------------------------
Private Sub Worksheet_Activate()
    Call ColorNoteRows
End Sub


'-------------------------------------------------------------------------------
Public Sub ColorNoteRows()
    ' Color each row based on value of cell in color_index column
    '     - mimics colors used by ColorNote
    '     - requires correct value for color_index_table_column
    '       where first column of table = 1, etc
    Dim used_range As Range
    Dim first_row As Long
    Dim first_column As Long
    Dim last_row As Long
    Dim last_column As Long
    Dim color_index_column As Long
    Dim color_index_column_offset As Long
    Dim i As Long

    color_index_table_column = 2 ' MAKE SURE THIS IS CORRECT

    Set used_range = ActiveSheet.UsedRange

    first_row = used_range(1).Row
    first_column = used_range(1).Column
    last_row = used_range(used_range.Cells.Count).Row
    last_column = used_range(used_range.Cells.Count).Column

    color_index_column = first_column + color_index_table_column - 1

    For i = first_row To last_row
        With ActiveSheet.Range(Cells(i, first_column), Cells(i, last_column))
            Select Case Cells(i, color_index_column).Value
                Case 1
                    .Interior.Color = RGB(255, 230, 233) ' red
                Case 2
                    .Interior.Color = RGB(255, 235, 216) ' orange
                Case 3
                    .Interior.Color = RGB(254, 248, 186) ' yellow
                Case 4
                    .Interior.Color = RGB(229, 248, 220) ' green
                Case 5
                    .Interior.Color = RGB(232, 233, 254) ' blue
                Case 6
                    .Interior.Color = RGB(239, 224, 255) ' purple
                Case 7
                    .Interior.Color = RGB(204, 204, 204) ' black
                Case 8
                    .Interior.Color = RGB(238, 238, 238) ' grey
                Case 9
                    .Interior.Color = RGB(255, 255, 255) ' white
                Case Else
                    .Interior.Color = xlNone
            End Select
        End With
    Next i
End Sub
