# Read text file into string
Dim strFilename As String: strFilename = "C:\temp\yourfile.txt"
Dim strFileContent As String
Dim iFile As Integer: iFile = FreeFile
Open strFilename For Input As #iFile
strFileContent = Input(LOF(iFile), iFile)
Close #iFile

# Read text file line by line
Open "C:\tester.txt" For Input As #1
    r = 0
    Do Until EOF(1)
        Line Input #1, Data
        Worksheets("UI").Range("H12").Offset(r, 0) = Data
        r = r + 1
    Loop
    Close #1
