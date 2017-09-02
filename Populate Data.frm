Dim CurPath As String
Dim ra As Range
Dim WA As Object
Dim RowValue, ColValue As Integer
Dim i, j As Integer
Dim from_text(), to_text As String
Dim Emp_Name As Range
Dim EName As String
Dim strFolder As String
Dim Word_File_Names(), Excel_File_Names() As Variant

Sub Submit_button()
  CurPath = ActiveWorkbook.Path     ' Holds the path of the excel file
   
  Erase Word_File_Names, Excel_File_Names
  
  Pass = Get_File_Names("*.doc?", Word_File_Names)
  Pass = Get_File_Names("*.xls?", Excel_File_Names)
  
  RowValue = 1
  ColValue = 0
  
  On Error Resume Next
    Dim fs, cf, strFolder
    EName = Sheets("Data_Entry").Range("B2:B2").Value
    strFolder = CurPath & "\" & EName
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(strFolder) = True Then
       MsgBox "Folder already exists!"
    Else
       Set cf = fs.CreateFolder(strFolder)

       For Each element In Word_File_Names
       Pass = Joining_Report(element)
       Next

       For Each element In Excel_File_Names
       Pass = Personal_Details(element)
       Next
    End If
  
  
   Dim RetVal
  Dim Cmd As String
  Dim FName As String
  
  Dim FileName_ As String
    FileName_ = Dir(CurPath & "\" & EName & "\")
          Do While FileName_ <> vbNullString
                'ReDim Preserve Array_Name(j)
                 FName = FileName_
                 'Cmd = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe Reader XI.lnk /t " & FName
                 'RetVal = Shell(Cmd, 1)
                 'Array_Name(j) = fileName
                 
             FileName_ = Dir
             j = j + 1
          Loop
  
End Sub

Function Get_File_Names(ByVal extension As String, ByRef Array_Name() As Variant) As String
    j = 0
    Dim FileName As String
    FileName = Dir(CurPath & "\Templates\" & extension)
          
          Do While FileName <> vbNullString
                ReDim Preserve Array_Name(j)
                Array_Name(j) = FileName
             FileName = Dir
             j = j + 1
          Loop
          
End Function

Function Joining_Report(ByVal File_Name As String)
   Erase from_text()
    i = 0

    Set WA = CreateObject("Word.Application")
    WA.Documents.Open (CurPath & "\Templates\" & File_Name)
    WA.Visible = True
                        With WA.ActiveDocument
                        Set myRange = .Content
                            With myRange.Find
                                .Execute FindText:="%*%", MatchWildcards:=True
                            End With

                          While myRange.Find.Found
                                FoundText = myRange.Text
                                FoundText = Replace(FoundText, "%", "")
                                ReDim Preserve from_text(i)
                                from_text(i) = FoundText
                                myRange.Find.Execute
                                i = i + 1
                          Wend
                        End With
          i = 0
        For Each Search_Text In from_text()
            Set ra = Cells.Find(What:=Search_Text, LookAt _
                    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
                    False, SearchFormat:=False)
    
            If ra Is Nothing Then
                i = i + 1
            Else
                to_text = ra.Offset(RowValue, ColValue)
    
                            With WA.ActiveDocument
                            Set myRange = .Content
                                With myRange.Find
                                    .Execute FindText:="%" & from_text(i) & "%", ReplaceWith:=to_text, Replace:=2
                                End With
                            End With
                        i = i + 1
            End If
        Next

Position = InStr(File_Name, ".")
'
File_Name_Trimed = Left(File_Name, Position - 1)

WA.ActiveDocument.SaveAs (CurPath & "\" & EName & "\" & File_Name)

'WA.PrintOut

''''''''''''WA.ActiveDocument.SaveAs2 CurPath & "\" & EName & "\" & File_Name_Trimed & ".pdf", 17
 
' With WA.ActiveDocument
'        .ExportAsFixedFormat OutputFileName:=Mid(.FullName, 1, InStrRev(.FullName, “.”)) & “.pdf”, _
'        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, _
'        OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, _
'        Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
'        DocStructureTags:=True, _
'        BitmapMissingFonts:=True, UseISO19005_1:=False
' End With
'
 'WA.ActivePrinter = previousPrinter
 'WA.PrintOut
 'WA.ExportAsFixedFormat CurPath & "\" & EName & "\" & File_Name & ".pdf"
 'WA.ExportAsFixedFormat OutputFileName:="C:\Ashik\Project HR Forms\Me_JoiningKit_test\Ashik\File.pdf", ExportFormat:=wdExportFormatPDF

'ActivePrinter = "Acrobat PDFWriter"
'ActiveDocument.PrintOut

 WA.Close
End Function

Function Personal_Details(ByVal File_Name As String)

Erase from_text
Dim counter As Integer
counter = 0
i = 0
   Dim Ch As Worksheet
   Dim CLoc As Range

   Dim Sh As Worksheet
   Dim Loc As Range
   
   Dim Fa As Worksheet
   
   Dim appXL As New Excel.Application
   appXL.Workbooks.Open (CurPath & "\Templates\" & File_Name)
   'appXL.Workbooks.Open ("C:\Ashik\Project HR Forms\Me_JoiningKit_test\Templates\02_RLE_Personal Details Form.xls")
   'C:\Ashik\Project HR Forms\Me_JoiningKit_test\Templates\02_RLE_Personal Details Form.xls
   appXL.ActiveWorkbook.Windows(1).Visible = True
   appXL.Visible = True
                             
   For Each Fa In appXL.ActiveWorkbook.Worksheets
   
   Dim rng As Range
   Set rng = Fa.UsedRange
   Dim cell As Range

        For Each cell In rng
        
        If counter = 900 Then
        Exit For
        Else
                If InStr(1, cell.Text, "%", vbTextCompare) > 0 Then
                    to_text = cell.Offset(0, 0)
                     FoundText = Replace(to_text, "%", "")
                                        ReDim Preserve from_text(i)
                                        from_text(i) = FoundText
                                        i = i + 1
                End If
            counter = counter + 1
            End If
        Next cell
   Next
                             
   i = 0
   For Each Sh In appXL.ActiveWorkbook.Worksheets
   For Each element In from_text
        For Each Ch In ThisWorkbook.Worksheets
                     With Ch.UsedRange
                      Set CLoc = .Cells.Find(What:=from_text(i))
                          If Not CLoc Is Nothing Then
                           Dim Replace_with As String
                            Replace_with = CLoc.Offset(1, 0)
                          End If
                     End With
        Next
   
   With Sh.UsedRange
       Set Loc = .Cells.Find(What:="%" & from_text(i) & "%")
    
        If Not Loc Is Nothing Then
        
                Set Loc = .FindNext(Loc)
                Dim To_find As String
                To_find = Loc.Offset(0, 0)
       
            appXL.Cells.Replace What:=To_find, Replacement:=Replace_with, LookAt:=xlWhole, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
             i = i + 1
         End If
    End With

    Next
    Next
   
    'appXL.ActiveWorkbook.SaveAs CurPath & "\" & EName & "\" & File_Name & ".pdf"
    'appXL.ActiveWorkbook.ExportAsFixedFormat xlTypePDF, CurPath & "\" & EName & "\" & File_Name & ".pdf"
    'appXL.ActiveSheet.PrintOut
    'ActiveWorkbook.SaveAs2 CurPath & "\" & EName & "\" & File_Name & ".pdf"
    
    
    appXL.ActiveWorkbook.SaveAs CurPath & "\" & EName & "\" & File_Name
    appXL.Close
                                                                                         'appXL.PrintOut
End Function

Sub Print_All_Files()
 
  CurPath = ActiveWorkbook.Path     ' Holds the path of the excel file
'
  Erase Word_File_Names, Excel_File_Names
'
'  Pass = Get_File_Names("*.doc?", Word_File_Names)
'  Pass = Get_File_Names("*.xls?", Excel_File_Names)
'
'  RowValue = 1
'  ColValue = 0
'
  On Error Resume Next
    Dim fs, cf, strFolder
    EName = Sheets("Data_Entry").Range("B2:B2").Value
    strFolder = CurPath & "\" & EName
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(strFolder) = True Then
       Dim Path            As String
        Dim FName           As String
         
        Path = CurPath & "\" & EName & "\"
'        FName = Dir(Path & "*.doc*")
'        Do While FName <> ""
'             Set WA = CreateObject("Word.Application")
'             WA.Documents.Open (Path & FName)
'             WA.PrintOut
'              WA.Close
'            'Application.PrintOut FileName:=Path & FName
'            FName = Dir()
'        Loop
        
        FName = Dir(Path & "*.xls*")
        Do While FName <> ""
             Dim appXL As New Excel.Application
             appXL.Workbooks.Open (Path & FName)
             appXL.ActiveSheet.PrintOut
             appXL.Close
            'Application.PrintOut FileName:=Path & FName
            FName = Dir()
        Loop
     
     End If
'    Else
'       Set cf = fs.CreateFolder(strFolder)
'
''       For Each element In Word_File_Names
''       Pass = Joining_Report(element)
''       Next
'
'       For Each element In Excel_File_Names
'       Pass = Personal_Details(element)
'       Next
'    End If
'
'
'   Dim RetVal
'  Dim Cmd As String
'  Dim FName As String
'
'  Dim FileName_ As String
'    FileName_ = Dir(CurPath & "\" & EName & "\")
'          Do While FileName_ <> vbNullString
'                'ReDim Preserve Array_Name(j)
'                 FName = FileName_
'                 'Cmd = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Adobe Reader XI.lnk /t " & FName
'                 'RetVal = Shell(Cmd, 1)
'                 'Array_Name(j) = fileName
'
'             FileName_ = Dir
'             j = j + 1
'          Loop
'
End Sub
