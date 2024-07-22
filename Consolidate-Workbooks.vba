Sub Consolidate-WBs()

    'Variables
    Dim folderPath As String
    Dim destinationPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim destWb As Workbook
    Dim destWs As Worksheet
    Dim lastRow As Long
    Dim destLastRow As Long

    'Destinations
    folderPath = "C:\Users\seu_usuário\"
    destinationPath = "C:\Users\seu_usuário\arquivo_final.xlsx"

    'Checks if the target file already exists and deletes
    If Dir(destinationPath) <> "" Then
        Kill destinationPath
    End If

    'Create the Workbook of destinationPath
    Set destWb = Workbooks.Add
    Set destWs = destWb.Sheets(1)
    destWs.Name = "arquivo_final"

    'Creating the loop structure in the folder
    fileName = Dir(folderPath & "*.*")
    Do While fileName <> ""
        ' Checks if the file is an Excel file on the formats xlsx, xlsm, xlsx and xlsb
        If Right(fileName, 4) = "xlsx" Or Right(fileName, 4) = "xlsm" Or Right(fileName, 4) = "xlsb" Or Right(fileName, 3) = "xls" Then
            ' Open the Excel file with links disabled
            Set wb = Workbooks.Open(folderPath & fileName, UpdateLinks:=0)

            ' Select the sheet "Data"
            On Error Resume Next
            Set ws = wb.Sheets("Data")
            On Error GoTo 0

            If Not ws Is Nothing Then
                ' If the sheet is hidden, then show it
                If ws.Visible = xlSheetHidden Then ws.Visible = xlSheetVisible

                ' Copy the data from the tab "Data"
                ' Searching the last row of the sheet
                lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

                ' If the sheet destination is empty, paste on A1
                If destWs.Cells(1, 1).Value = "" Then
                    ' Setting the range of copy
                    ws.Range("B7:BG" & lastRow).Copy
                    ' Pasting special (only values)
                    destWs.Range("A1").PasteSpecial Paste:=xlPasteValues
                
                ' Else, finding the next row empty
                Else
                    destLastRow = destWs.Cells(destWs.Rows.Count, "A").End(xlUp).Row + 1
                    ws.Range("B7:BG" & lastRow).Copy
                    destWs.Range("A" & destLastRow).PasteSpecial Paste:=xlPasteValues
                End If

                ' Disabling the CutCupyMode
                Application.CutCopyMode = False
            End If

            ' Closing the file without changes
            wb.Close SaveChanges:=False

            ' Restarting the variable ws
            Set ws = Nothing
        End If

        ' Next File
        fileName = Dir
    Loop

    ' Save the destination file
    destWb.SaveAs destinationPath
    destWb.Close SaveChanges:=True

    MsgBox "Sucess!"


End Sub