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
    
# Creating an exe or package of a vba based workbook
Solution 1
Sub workbook_open()
    Dim exe As Excel.Application
    Set exe = Application
    exe.Visible = False
End Sub

Solution 2
-- Use a VBScript rather than an exe file.  Open Notepad, and enter the following in the text file:
Set XL=CreateObject("Excel.Application")
XL.Visible=True
XL.Workbooks.Open "C:\A.xls"
XL.Run "A.xls!TheMacro"

-- Change "C:\A.xls" to the appropriate file name. Change "A.xls!!TheMacro" to the workbook name and macro name.

-- Save the file as a vbs file, e.g., RunIt.vbs rather than a txt file.
