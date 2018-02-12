@@ -0,0 +1,275 @@
﻿' History of changes
' 02/12/2018 ns
'   - updated CreateExcelFileFromDataTable for new reporting service
' 12/15/2015 dl
'   - add try/catch/finally logic around filling data set from spreadsheet to handle errors being thrown.

Imports Microsoft.VisualBasic
Imports System.Data
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet


Public Module modExcelHelper

    Public Function LoadStringFromExcelFile(ByVal strPath As String, ByVal strFilename As String) As String

        Dim DS As System.Data.DataSet = Nothing
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        Dim MyConnection As System.Data.OleDb.OleDbConnection

        Dim strSQL As String = ""

        'load the contents of the excel file to a dataset
        MyConnection = New System.Data.OleDb.OleDbConnection( _
        "provider=Microsoft.ACE.OLEDB.12.0; " & _
        "data source=" & strPath & strFilename & "; " & _
        "Extended Properties=Excel 12.0;")

        MyConnection.Open()

        Dim dtTables As DataTable = MyConnection.GetOleDbSchemaTable(OleDb.OleDbSchemaGuid.Tables, Nothing)
        Dim tableName As String = dtTables.Rows(0).Item("Table_Name").ToString()

        'if a sheet was programatically named it may not have the $ in the sheet name
        strSQL = "SELECT * FROM [" & tableName

        If Right(strSQL, 1) = "'" Then
            If Mid(strSQL, strSQL.Length - 1, 1) <> "$" Then
                strSQL = Mid(strSQL, 1, strSQL.Length - 1) & "$'"
            End If
        Else
            If Right(strSQL, 1) <> "$" Then
                strSQL &= "$"
            End If
        End If

        strSQL &= "]"

        MyCommand = New System.Data.OleDb.OleDbDataAdapter(strSQL, MyConnection)

		Dim ReadSpreadsheet As Boolean = False
		DS = New System.Data.DataSet
		Try
			MyCommand.Fill(DS)
			ReadSpreadsheet = True
		Catch ex As Exception
			LogMessage("Exception thrown while filling dataset from spreadsheet '" & strFilename & "'  Exception message:" & ex.Message)
		Finally
			MyConnection.Close()
		End Try

		' If reading the spreadsheet failed (such as having bad data or bad name)
		If (ReadSpreadsheet = False) Then
			Return String.Empty
		End If

        Dim tbl As New DataTable
        tbl = DS.Tables(0)

        Dim sb As New System.Text.StringBuilder(1000)
        Dim sw As New System.IO.StringWriter(sb)

        DS.WriteXml(sw, System.Data.XmlWriteMode.WriteSchema)

        Return sb.ToString()

    End Function

    Public Sub CreateExcelFileFromDataTable(ByVal FilePath As String, myDT As DataTable)

        ' Create a spreadsheet document by supplying the filepath.
        ' By default, AutoSave = true, Editable = true, and Type = xlsx.
        Dim spreadsheetDocument As SpreadsheetDocument = spreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook)

        ' Add a WorkbookPart to the document.
        Dim workbookpart As WorkbookPart = spreadsheetDocument.AddWorkbookPart
        workbookpart.Workbook = New Workbook

        ' Add a WorksheetPart to the WorkbookPart.
        Dim worksheetPart As WorksheetPart = workbookpart.AddNewPart(Of WorksheetPart)()
        worksheetPart.Worksheet = New Worksheet(New SheetData())

        ' Add Sheets to the Workbook.
        Dim sheets As Sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())

        ' Append a new worksheet and associate it with the workbook.
        Dim sheet As Sheet = New Sheet
        sheet.Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart)
        sheet.SheetId = 1
        sheet.Name = "Duplicate Document"

        sheets.Append(sheet)

        'get the sheetData object so we can add the data table to it
        Dim sheetData As SheetData = worksheetPart.Worksheet.GetFirstChild(Of SheetData)()

        'add the data table
        AddDataTable(myDT, sheetData)

        'save the workbook
        workbookpart.Workbook.Save()

        ' Close the document.
        spreadsheetDocument.Close()

        ' -----------------------------------

    End Sub

    Private Sub AddDataTable(ByRef exportData As DataTable, ByRef sheetdata As SheetData)

        'add column names to the first row   
        Dim Header As Row = New Row()
        Header.RowIndex = 1

        For Each col As DataColumn In exportData.Columns
            Dim headerCell As Cell = createTextCell(exportData.Columns.IndexOf(col) + 1, Convert.ToInt32(Header.RowIndex.Value), col.ColumnName)
            Header.AppendChild(headerCell)
        Next

        sheetdata.AppendChild(Header)

        'loop through each data row   
        Dim contentRow As DataRow
        Dim intStartRow As Int32 = 2

        For intLoop = 0 To exportData.Rows.Count - 1
            contentRow = exportData.Rows(intLoop)
            sheetdata.AppendChild(createContentRow(contentRow, intLoop + intStartRow))
        Next

    End Sub

    Private Function InsertCellInWorksheet(ByVal columnName As String, ByVal rowIndex As UInteger, ByVal worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = (columnName + rowIndex.ToString())
        ' If the worksheet does not contain a row with the specified row index, insert one.        
        Dim row As Row
        If (sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).Count() <> 0) Then
            row = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).First()
        Else
            row = New Row()
            row.RowIndex = rowIndex
            sheetData.Append(row)
        End If
        ' If there is not a cell with the specified column name, insert one.          
        If (row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = columnName + rowIndex.ToString()).Count() > 0) Then
            Return row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.            
            Dim refCell As Cell = Nothing
            For Each cell As Cell In row.Elements(Of Cell)()
                If (String.Compare(cell.CellReference.Value, cellReference, True) > 0) Then
                    refCell = cell
                    Exit For
                End If
            Next
            Dim newCell As Cell = New Cell
            newCell.CellReference = cellReference
            row.InsertBefore(newCell, refCell)
            worksheet.Save()
            Return newCell
        End If
    End Function

    Private Function createTextCell(ByVal columnIndex As Integer, ByVal rowIndex As Integer, ByVal cellValue As Object) As Cell

        Dim cell As New Cell

        cell.DataType = CellValues.InlineString
        cell.CellReference = getColumnName(columnIndex) & rowIndex.ToString

        Dim inlineSTring As New InlineString()

        Dim t As New Text()

        t.Text = cellValue.ToString()
        inlineSTring.AppendChild(t)
        cell.AppendChild(inlineSTring)

        Return cell

    End Function

    Private Function createContentRow(ByVal dataRow As DataRow, ByVal rowIndex As Integer) As Row

        Dim row As New Row

        For i As Integer = 0 To dataRow.Table.Columns.Count - 1
            Dim datacell As Cell = createTextCell(i + 1, rowIndex, dataRow(i))
            row.AppendChild(datacell)
        Next

        Return row

    End Function


    Private Function getColumnName(ByVal columnIndex As Integer) As String

        Dim strCol As String = ""

        Select Case columnIndex
            Case 1
                strCol = "A"
            Case 2
                strCol = "B"
            Case 3
                strCol = "C"
            Case 4
                strCol = "D"
            Case 5
                strCol = "E"
            Case 6
                strCol = "F"
            Case 7
                strCol = "G"
            Case 8
                strCol = "H"
            Case 9
                strCol = "I"
            Case 10
                strCol = "J"
            Case 11
                strCol = "K"
            Case 12
                strCol = "L"
            Case 13
                strCol = "M"
            Case 14
                strCol = "N"
            Case 15
                strCol = "O"
            Case 16
                strCol = "P"
            Case 17
                strCol = "Q"
            Case 18
                strCol = "R"
            Case 19
                strCol = "S"
            Case 20
                strCol = "T"
            Case 21
                strCol = "U"
            Case 22
                strCol = "V"
            Case 23
                strCol = "W"
            Case 24
                strCol = "X"
            Case 25
                strCol = "Y"
            Case 26
                strCol = "Z"
        End Select

        Return strCol

    End Function


End Module
