Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.Math
Module Module1

    Sub Main()
        sheetCopy()
    End Sub


    Private exitCount As Integer = 10
    Private SharedStringItem As UInt32Value
    Private p
    'EXCELからEXCELシートコピー
    Private Sub sheetCopy()

        Dim path1 As String = "C:\Users\81904\Desktop\新しいフォルダー\Book1.xlsx" 'コピー元
        Dim path2 As String = "C:\Users\81904\Desktop\新しいフォルダー\Book2.xlsm" 'コピ先

        Using spreadsheetDocument As SpreadsheetDocument _
                 = SpreadsheetDocument.Open(path1, True)
            Using spreadsheetDocument2 As SpreadsheetDocument _
              = SpreadsheetDocument.Open(path2, True)

                ' シート名からWorksheetオブジェクトを取得する
                Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
                Dim sheet As Sheet _
                  = workbookPart.Workbook.GetFirstChild(Of Sheets)().Elements(Of Sheet)() _
                      .FirstOrDefault(Function(s) s.Name = "コピー元")
                Dim relationshipId As String = sheet.Id.Value
                Dim worksheetPart As WorksheetPart = workbookPart.GetPartById(relationshipId)
                Dim worksheet As Worksheet = worksheetPart.Worksheet


                ' シート名からWorksheetオブジェクトを取得する
                Dim workbookPart2 As WorkbookPart = spreadsheetDocument2.WorkbookPart
                Dim sheet2 As Sheet _
                  = workbookPart2.Workbook.GetFirstChild(Of Sheets)().Elements(Of Sheet)() _
                      .FirstOrDefault(Function(s) s.Name = "コピー先")
                Dim relationshipId2 As String = sheet2.Id.Value
                Dim worksheetPart2 As WorksheetPart = workbookPart2.GetPartById(relationshipId2)
                Dim worksheet2 As Worksheet = worksheetPart2.Worksheet

                Dim youso As Integer = worksheet.Count

                Dim theCell As Cell
                Dim theCell2 As Cell

                Dim exRowCnt As Integer = 0
                Dim exColCnt As Integer = 0

                Dim bNotSet As Integer = 0
                Dim bSet As Boolean = False

                For row As Integer = 1 To 1048576 - 1
                    'コピー元セルをループします。



                    For col As Integer = 1 To 16384 - 1
                        Dim colName As String = ColumnName(col)
                        Dim addressName As String = colName + row.ToString()

                        theCell = worksheetPart.Worksheet.Descendants(Of Cell).
                            Where(Function(c) c.CellReference = addressName).FirstOrDefault
                        If (theCell IsNot Nothing) Then

                            Dim val As String = getCell(workbookPart, theCell, addressName)

                            theCell2 = worksheetPart2.Worksheet.Descendants(Of Cell).
                            Where(Function(c) c.CellReference = addressName).FirstOrDefault
                            If (theCell2 Is Nothing) Then
                                Dim shareStringPart As SharedStringTablePart

                                If (spreadsheetDocument2.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).Count() > 0) Then
                                    shareStringPart = spreadsheetDocument2.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).First()
                                Else
                                    shareStringPart = spreadsheetDocument2.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
                                End If

                                ' Insert the text into the SharedStringTablePart.
                                Dim index As Integer = InsertSharedStringItem(val, shareStringPart)

                                Dim cell As Cell = InsertCellInWorksheet(colName, row, worksheetPart2)
                                cell.CellValue = New CellValue(index.ToString)
                                cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
                            Else
                                Dim shareStringPart As SharedStringTablePart

                                If (spreadsheetDocument2.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).Count() > 0) Then
                                    shareStringPart = spreadsheetDocument2.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).First()
                                Else
                                    shareStringPart = spreadsheetDocument2.WorkbookPart.AddNewPart(Of SharedStringTablePart)()
                                End If

                                ' Insert the text into the SharedStringTablePart.
                                Dim index As Integer = InsertSharedStringItem(val, shareStringPart)

                                theCell2.CellValue = New CellValue(index.ToString)
                                theCell2.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)
                            End If
                            bSet = True
                        Else
                            exColCnt += 1
                        End If

                        If (exColCnt > exitCount) Then
                            exColCnt = 0
                            Exit For
                        End If
                    Next

                    If (bSet) Then
                        bNotSet = 0
                    Else
                        bNotSet += 1

                        If (bNotSet > 10) Then
                            Exit For
                        End If
                    End If
                    bSet = False
                Next

                worksheetPart2.Worksheet.Save()

            End Using
        End Using

    End Sub


    'EXCELからCSV出力
    Private Sub exceltocsv()



    End Sub

    Private Function getCell(wbPart As WorkbookPart, theCell As Cell, addressName As String)
        Dim value As String = Nothing

        If theCell IsNot Nothing Then
            value = theCell.InnerText

            If theCell.DataType IsNot Nothing Then
                Select Case theCell.DataType.Value
                    Case CellValues.SharedString
                        Dim stringTable = wbPart.
                          GetPartsOfType(Of SharedStringTablePart).FirstOrDefault()

                        If stringTable IsNot Nothing Then
                            Dim oe As OpenXmlElement = stringTable.SharedStringTable.
                            ElementAt(Integer.Parse(value))
                            value = oe.InnerText

                        End If

                    Case CellValues.Boolean
                        Select Case value
                            Case "0"
                                value = "FALSE"
                            Case Else
                                value = "TRUE"
                        End Select
                End Select
            Else
                value = theCell.CellValue.Text.ToString()
            End If

        End If
        Return value
    End Function

    Private Function InsertSharedStringItem(ByVal text As String, ByVal shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
    End Function

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

            Return newCell
        End If
    End Function

    ' アルファベット文字数(number of alphabet characters
    Private Const NoAC As Integer = 26

    ''' <summary>
    ''' 表計算ワークシート列名作成
    ''' </summary>
    ''' <param name="column">1以上の列番号</param>
    ''' <returns>列名</returns>
    ''' <remarks></remarks>
    Public Function ColumnName(column As Integer) As String
        Dim name As String = String.Empty
        Dim v = column
        Do
            name = Chr(Asc("A"c) + (v - 1) Mod NoAC) & name
            If v <= NoAC Then
                Exit Do
            Else
                v = CInt(System.Math.Ceiling((v - NoAC) / NoAC))
            End If
        Loop
        Return name
    End Function

End Module
