Sub data_to_text_file(sFileName)

'variables that you need to use in the code
Dim TextFile As Integer
Dim iCol As Integer
Dim myRange As Range
Dim cVal As Range
Dim i As Integer
Dim myFile As String

'define the range that you want to writeSheets("Main").Select
Set myRange = Sheets("Main").Range("K4:L40")
iCol = myRange.Rows.Count

'path to the text file MAKE SURE TO CHANGE THE LOCATION !!!
myFile = "C:\temp\OutPut" & sFileName & ".txt"

'define FreeFile to the variable file number
TextFile = FreeFile

'using append command to add text to the end of the file
Open myFile For Output As TextFile

'loop to add data to the text file
For i = 1 To iCol
Print #TextFile, myRange(i, 1),
Print #TextFile, myRange(i, 2)
Next i

'close command to close the text file after adding data
Close #TextFile

End Sub
