Sub ExportTableToTxt()
    Dim tbl As Table
    Dim row As Integer, col As Integer
    Dim outputFile As String
    Dim fileNumber As Integer
    Dim cellContent As String

    ' Set output file name
    outputFile = ThisDocument.Path & "\" & Left(ThisDocument.Name, InStrRev(ThisDocument.Name, ".") - 1) & ".qst"
    fileNumber = FreeFile

    ' Open the output file for writing
    Open outputFile For Output As #fileNumber

    ' Get the first table in the document
    If ThisDocument.Tables.Count > 0 Then
        Set tbl = ThisDocument.Tables(1)

        ' Loop through each row in the table
        For row = 1 To tbl.Rows.Count
            ' Get the question from the first column
            cellContent = Trim(tbl.Cell(row, 1).Range.Text)
            Print #fileNumber, "? " & cellContent

            ' Get the correct answer from the second column
            cellContent = Trim(tbl.Cell(row, 2).Range.Text)
            Print #fileNumber, "+ " & cellContent

            ' Loop through the remaining columns for incorrect answers
            For col = 3 To tbl.Columns.Count
                cellContent = Trim(tbl.Cell(row, col).Range.Text)
                If Len(cellContent) > 0 Then
                    Print #fileNumber, "- " & cellContent
                End If
            Next col
        Next row
    Else
        MsgBox "No tables found in the document!", vbExclamation
    End If

    ' Close the output file
    Close #fileNumber

    MsgBox "Data successfully exported to " & outputFile, vbInformation
End Sub
