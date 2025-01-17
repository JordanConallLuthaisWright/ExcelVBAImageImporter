Attribute VB_Name = "Module1"
Sub ImportImagesWithResizingAndBordersFixed()
    Dim ws As Worksheet
    Dim folderPath As String
    Dim img As String
    Dim rowNum As Integer
    Dim cell As Range
    Dim fileCount As Integer
    Dim pic As Shape
    Const PicMaxWidth As Double = 100 ' Set a fixed width for all images
    Const PicMaxHeight As Double = 100 ' Set a fixed height for all images

    ' Set the worksheet and folder path
    Set ws = ThisWorkbook.Sheets(1) ' Adjust to your desired sheet
    folderPath = "C:\Users\jorda\Downloads\House defects\House defects\" ' Confirm this path is correct

    ' Check if the folder exists
    If Dir(folderPath, vbDirectory) = "" Then
        MsgBox "Folder path is invalid!"
        Exit Sub
    End If

    ' Clear worksheet and reset row number
    rowNum = 1
    ws.Cells.Clear
    fileCount = 0

    ' Set uniform column width and row height for consistency
    ws.Columns("A").ColumnWidth = 25
    ws.Columns("B").ColumnWidth = 30
    ws.Rows.RowHeight = PicMaxHeight + 10

    ' Loop through all image types
    Dim fileTypes As Variant
    fileTypes = Array("*.jpg", "*.jpeg", "*.png", "*.bmp", "*.gif")

    Dim i As Integer
    For i = LBound(fileTypes) To UBound(fileTypes)
        img = Dir(folderPath & fileTypes(i))
        Do While img <> ""
            fileCount = fileCount + 1
            Set cell = ws.Cells(rowNum, 1)

            ' Insert the image as a shape instead of Picture object
            Set pic = ws.Shapes.AddPicture(folderPath & img, _
                    msoFalse, msoCTrue, _
                    cell.Left + 5, cell.Top + 5, PicMaxWidth, PicMaxHeight)

            ' Lock aspect ratio and size properly
            With pic
                .LockAspectRatio = msoTrue
                ' Adjust height if it exceeds the limit while keeping ratio
                If .Height > PicMaxHeight Then
                    .Height = PicMaxHeight
                End If
                If .Width > PicMaxWidth Then
                    .Width = PicMaxWidth
                End If
            End With

            ' Add a comment cell next to the image
            ws.Cells(rowNum, 2).Value = "Enter comment here"

            ' Apply a professional border around the image and comment
            With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, 2)).Borders
                .LineStyle = xlContinuous
                .Color = RGB(0, 0, 0) ' Black border
                .Weight = xlMedium
            End With

            ' Move to the next row
            rowNum = rowNum + 1
            img = Dir
        Loop
    Next i

    ' Completion message
    If fileCount = 0 Then
        MsgBox "No images were found in the specified folder. Please check the folder path and image formats."
    Else
        MsgBox "Import complete! " & fileCount & " images imported, resized, and bordered."
    End If
End Sub


