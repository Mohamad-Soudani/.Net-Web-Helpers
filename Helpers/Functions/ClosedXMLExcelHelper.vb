Namespace Helpers

    Public Class ClosedXMLExcelHelper
        Public Function ImportExcel(filePath As String, Optional sheetName As Object = "") As DataTable
            'Save the uploaded Excel file.

            'Open the Excel file using ClosedXML.
            Using workBook As New XLWorkbook(filePath)
                'Read the first Sheet from Excel file.
                Dim workSheet As IXLWorksheet = workBook.Worksheet(sheetName)
                'Create a new DataTable.
                Dim dt As New DataTable()
                'Loop through the Worksheet rows.
                Dim firstRow As Boolean = True
                For Each row As IXLRow In workSheet.Rows()
                    'Use the first row to add columns to DataTable.
                    If firstRow Then
                        For Each cell As IXLCell In row.Cells()
                            Dim filteredColName = cell.Value.ToString()
                            If (filteredColName.IndexOf("/") > -1) Then
                                filteredColName = filteredColName.RemoveMultiple({"/"})
                            End If
                            dt.Columns.Add(filteredColName)
                        Next
                        firstRow = False
                    Else
                        'Add rows to DataTable.
                        Dim toInsert As DataRow = dt.NewRow()
                        Dim i As Integer = 0
                        For Each cell As IXLCell In row.Cells(1, dt.Columns.Count)
                            toInsert(i) = cell.Value.ToString()
                            i += 1
                        Next
                        dt.Rows.Add(toInsert)
                    End If
                Next
                Return dt
            End Using
        End Function

        Public Function GetWorkBookSheets(filePath As String, Optional excludeHeader As Boolean = True) As List(Of String)
            Dim sheetList As New List(Of String)
            Using workBook As New XLWorkbook(filePath)
                For Each sheet As IXLWorksheet In workBook.Worksheets
                    sheetList.Add(sheet.Name)
                Next
            End Using
            Return sheetList
        End Function


    End Class
End Namespace
