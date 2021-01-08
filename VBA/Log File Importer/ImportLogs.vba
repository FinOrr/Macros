''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FILENAME :        ImportLogs.             DESIGN REF: FMCM00
'
' DESCRIPTION :
' 		Loops through a directory, and pulls files with a specific extension into a workbook
' 		Each file is written to a seperate worksheet; the worksheet name is set to the file name
' 		The user is prompted to select a folder using the windows file dialog. 
' 		The second parent directory is used to save the excel workbook. For example:
' 			> Selected folder is "C:\Arduino\MyProject\Testing\SerialTest\TestRun0\"
'			> The file will be saved as "C:\Arduino\MyProject\Testing\SerialTest\TestRun0\SerialTest.Results.xlsx" when the macro finishes
'
' PUBLIC FUNCTIONS :
'		ImportLogs
'
' AUTHOR :    Fin Orr        START DATE :    Nov 2020
'
' CHANGES :
'
' REF NO	VERSION		DATE			WHO		DETAIL
' 0  		0.02		08/01/2020		FO		Made more generic by removing application specifics and added guidance through comments
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub ImportLogs()
    Dim MyFile As String
    Dim MyAdr As String
    Dim Counter As Long
    Dim MyFolder As String
    Dim DirArray() As String
    Dim SheetCounter As Integer: SheetCounter = 1
    
    Dim Heap As Integer: Heap = FreeFile
    Dim Lines() As String, i As Long
    
    Application.ScreenUpdating = False
    
	' [[ USER CONFIGURABLE SETTING ]]
    ' Create a dynamic array variable that supports up to 100 files
	' If you're expecting to import more than 100 files into one excel worksheet, then update the size of the array to >100
    Dim DirectoryListArray() As String
    ReDim DirectoryListArray(100)
    
    ' Get the user to select the folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = -1 Then
            MyFolder = .SelectedItems(1)
        End If
    End With
            
    ' When the user selects a folder
    If MyFolder <> "" Then
        ' Loop through all the files in the directory with the Dir$ function, looking for files with the extension ".log"
        MyFile = Dir$(MyFolder & "\*.log")
        'Debug.Print "MyFile = " & MyFile
        Do While MyFile <> ""
            DirectoryListArray(Counter) = MyFile
            MyFile = Dir$
            Counter = Counter + 1
        Loop
        
        ' Reset the size of the array, while maintaining the contents
        ReDim Preserve DirectoryListArray(Counter - 1)
       
		' After we've pulled all the files into the 
		For Counter = 0 To UBound(DirectoryListArray)
            MyAdr = MyFolder & "\" & DirectoryListArray(Counter)
            
            ' Open the log file and read it into the Lines() string array
            Open (MyAdr) For Input As #Heap
                Lines = Split(Input$(LOF(Heap), #Heap), vbNewLine)
            Close #Heap
            
            ' Write the values to the cells
            If UBound(Lines) > 0 Then
                For i = LBound(Lines) To UBound(Lines)
                    '   Debug.Print Lines(i)
                     ActiveSheet.Cells(i + 1, 1).Value = Lines(i)
                Next i
                Columns("A").Select
				
				' [[ USER CONFIGURABLE SETTING ]]
				' How are your log files delimited? If importing a CSV, then Comma:=True etc...
                Selection.TextToColumns _
                    Destination:=Range("A:A"), _
                    DataType:=xlDelimited, _
                    TextQualifier:=xlDoubleQuote, _
                    ConsecutiveDelimiter:=False, _
                    Tab:=True, _
                    Semicolon:=False, _
                    Comma:=False, _
                    Space:=False, _
                    Other:=False
            End If          
            			
			' [[ USER CONFIGURABLE SETTING ]]
			' If you want to run another subroutine dependent on the type of results file, the InStr function is used.
			' This function looks for a keyword (example below uses "input" and "output") in the file name and runs functions on the worksheet dependent on what it finds.
			' For example if you're logging a system's input and output and want to format them differently, here's an example:
			If InStr(DirectoryListArray(Counter), "Input") Then
                'Debug.Print "Formatting [" & DirectoryListArray(Counter) & "] as PI log file"
                FormatInputMacro
            Else
                'Debug.Print "Could not find 'Input' in: " & DirectoryListArray(Counter)
            End If
            
            If InStr(DirectoryListArray(Counter), "Output") Then
                'Debug.Print "Formatting [" & DirectoryListArray(Counter) & "] as Attingimus log file"
				FormatOutputMacro		' Call FormatOutputMacro()
            Else
            End If

			' Update the name of the active worksheet to the name of the file we're reading from
            Worksheets(SheetCounter).Name = DirectoryListArray(Counter)
            
            ' Check if this is the last log file in the directory
            If Counter = UBound(DirectoryListArray) Then
                ' Activate the first worksheet if we're finished
                Worksheets(1).Activate
            Else
                ' Else add a new worksheet at the end of the workbook and set it as active
                Sheets.Add After:=Sheets(Sheets.Count)
                SheetCounter = SheetCounter + 1
            End If
        Next Counter
    End If
    
    ' Split the directory of the current excel spreadsheet into individual array elements
    DirArray = Split(MyFolder, "\")
    ' The test clause is saved as the 2nd parent folder (.../ 2ND PARENT FOLDER / 1ST PARENT FOLDER / LOG FILE.LOG)
    TestClause = DirArray(UBound(DirArray) - 1)
    ' Save the excel file as the test clause results
    ActiveWorkbook.SaveAs Filename:=MyFolder & "\" & TestClause & ".Results.xlsx"
    
    Application.ScreenUpdating = True
End Sub
