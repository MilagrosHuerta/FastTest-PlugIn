Attribute VB_Name = "FTP_Problems"
' ----------------------------------------------------------------------- '
' --- MACROS TO USE IN THE FastTest PlugIn PROBLEM SHEET              --- '
' ---                           CREATED BY                            --- '
' ---             MILAGROS HUERTA GÓMEZ DE MERODIO                    --- '
' ----------------------------------------------------------------------- '
Option Explicit
Dim MyTable As ListObject
Dim NumColumn As Long
Dim NumFile As Long
Dim RuteSaved As String
Dim TempFilePath As String
Dim TempFileName As String
Dim Width_Image As Integer
Dim Height_Image As Integer
Dim SHOW_RESULTS As String
Dim i As Long
Dim Height_File As Long
Sub Paste_Results()
' ---------------------------------------------------------- '
' --- Macro to paste the results from FIXED DATA --- '
' ---------------------------------------------------------- '
Dim Colu_Ini As Integer
Dim Colu_Fin As Integer
Dim Fila_Ini As Integer

    Set MyTable = [Image].ListObject

    Colu_Ini = [Solu_ini].Column
    Colu_Fin = [Image].Column - 1
    Fila_Ini = [Solu_ini].Row

    [D_Fixed] = [_OK]
    
    For i = 1 To MyTable.ListRows.Count
        [Name_Data] = MyTable.ListRows(i).Range.Cells(1, 1)
        ActiveSheet.Range(Cells(Fila_Ini, Colu_Ini), Cells(Fila_Ini, Colu_Fin)).Select
        Selection.Copy
        MyTable.ListRows(i).Range.Cells(1, Colu_Ini).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    Next i
    MyTable.ListRows(1).Range.Cells(1, 2).Select
    MsgBox "END"
End Sub
Sub Images_Save_Fixed_Data()
' --------------------------------------------------------------------------------- '
' --- Macro to save images, depending on the Fixed Data                         --- '
' --- When saving images to the hard drive, if it goes too fast,                --- '
' --- it doesn't save them properly                                             --- '
' ---  SELECT THE ROW AND PRESS F5 TO EXECUTE SLOWLY                            --- '
'       .Chart.Paste ' is in the next macro                                     --- '
' --------------------------------------------------------------------------------- '
' Sheet name FIXED DATA defined as [name] in case someone changes it
    Sheets([N_Sheet_Data]).Select
' The NAME of the TABLE to which the IMAGE cell belongs, in case someone changes the table name
    Set MyTable = [Image].ListObject
    NumColumn = [Image].Column
    NumFile = [Image].Row
    [See_Data] = [_OK]
    SHOW_RESULTS = MsgBox("Show results in the image?", vbYesNo)
' Asks whether to show the results in the IMAGE or not
    If SHOW_RESULTS = vbYes Then [See_Results] = [_OK] Else [See_Results] = [_NO]
' Loop from the first row of the DATA table to the last
    For i = 1 To MyTable.ListRows.Count
    ' Change the problem data
        Sheets([N_Sheet_Pb]).Select
        [Name_Data] = MyTable.ListRows(i).Range.Cells(1, 1)
    ' This one only works if it's a single figure, not a set of grouped images
    '    ActiveSheet.ChartObjects("FIGURA_PROBLEMA").Copy
    ' Select the problem image - If it's a set of grouped images...
        ActiveSheet.Shapes.Range(Array([N_Figure])).Select
        Selection.Copy
    ' Paste the image in the corresponding row
        Sheets([N_Sheet_Data]).Select
        Cells(i + NumFile, NumColumn).Select
        Height_File = ActiveCell.EntireRow.Height        ' En pixeles
        ActiveSheet.Pictures.Paste.Select
        Selection.ShapeRange.Height = Height_File * 0.95 ' to 95%
        Selection.Placement = xlMove                   ' To move but no change size with cell
        Selection.ShapeRange.Name = MyTable.ListRows(i).Range.Cells(1, 1)
' ----------------------------------------------------------- '
        Call Save_Image_File
' ----------------------------------------------------------- '
        Next i
    MsgBox "END"
End Sub
Sub Save_Image_File()
' ------------------------------------------------------------------------------ '
' --- This macro is for saving images generated with data as jpg files       --- '
' --- DO NOT RUN ALONE IF AN IMAGE HAS NOT BEEN SELECTED                     --- '
' ------------------------------------------------------------------------------ '
Dim FileName As String
    ' El ancho y alto de la imagen, por 4, para pegar la imagen en CHART
        Width_Image = Selection.ShapeRange.Width * [Multiple_Img] ' 2.5  ' change between 2 and 4
        Height_Image = Selection.ShapeRange.Height * [Multiple_Img] ' 2.5  ' change between 2 and 4
    ' Get the path of the current file and the Sub_Folder
        RuteSaved = ActiveWorkbook.Path & "\" & [SUB_FOLDER] & "\"
    ' If the path doesn't exist, create it
        If Dir(RuteSaved, vbDirectory) = "" Then
            MkDir RuteSaved
        End If
    ' Create a temporary path and a unique temporary file name
        TempFilePath = Environ("TEMP") & "\"
        TempFileName = "TempImage" & Format(Now, "yyyyMMddhhmmss") & ".png"
    ' Save the image to the file
    ' The width depends on the image
        With ActiveSheet.ChartObjects.Add(Left:=0, Width:=Width_Image, Top:=0, Height:=Height_Image)
'----------------------------------------------------------- '
'---   MARK THE ROW AND PRESS F5 TO EXECUTE SLOWLY        ---'
            .Chart.Paste   ' ------------------------------- '
'----------------------------------------------------------- '
            .Chart.Export TempFilePath & TempFileName, "PNG"
            .Delete
        End With
    'Move the temporary file to the final location'
        FileName = [Name_Data] & ".png"
        FileCopy TempFilePath & TempFileName, RuteSaved & FileName
    'Delete the temporary file'
        Kill TempFilePath & TempFileName
    MsgBox "END"
End Sub
Sub List_Defined_Names()
    Dim MyName As Name
    Dim ws As Worksheet
    Dim Sheet_N As String
    Dim List_Values_Formula As Long
    
    On Error Resume Next
    Sheet_N = "NAMES"
    If Worksheets(Sheet_N) Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = Sheet_N
        ws.Move After:=Sheets(Sheets.Count)
        Range("'" & Sheet_N & "'!B2").Select
        ActiveWindow.FreezePanes = True
    Else
        Worksheets(Sheet_N).Cells.ClearContents

    End If
    Worksheets(Sheet_N).Select
    Range("A1") = "Defined Name"
    Range("B1") = "TEXT Reference"
    Range("C1") = "TEXT Formula"
    Range("D1") = "Type"
    List_Values_Formula = MsgBox("Do you want to see the result of the formula? This may generate circular references, in which case it would be advisable to delete them later.", vbYesNo)
    If List_Values_Formula = vbYes Then
        Range("E1") = "Reference"
        Range("F1") = "Formula"
    End If
    
    i = 1
    For Each MyName In ActiveWorkbook.Names
        If Left(MyName.Name, 6) <> "_xlfn." Then
            i = i + 1
            Range("A" & i) = MyName.Name
            Range("B" & i) = "'" & MyName.RefersToLocal
            Range("C" & i) = "'" & IIf(MyName.RefersToRange.HasFormula, MyName.RefersToRange.FormulaLocal, MyName.RefersToRange.Value)  'MyName.Formula
            Range("D" & i) = TypeName(MyName.RefersToRange)
            If List_Values_Formula = vbYes Then
                Range("E" & i) = MyName.RefersToLocal
                Range("F" & i) = IIf(MyName.RefersToRange.HasFormula, MyName.RefersToRange.Formula, "")  'MyName.Formula
            End If
        End If
    Next MyName
    Cells.Select
    Selection.WrapText = False
    ActiveWindow.DisplayHeadings = False
    Cells(2, 1).Select
    
    Columns("A:A").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("B:C").Select
    Selection.ColumnWidth = 50
    Range("B2").Select
    MsgBox "END"
End Sub
Sub Create_Defined_Names_from_Sheet()
    Dim ws As Worksheet
    Dim Rng As Range
    Dim Fila As Long
    Dim Name As String
    Dim Reference As String
    Dim Tipo As String
    Dim Formula As String
    Dim DefinedName As Name
    Dim Not_Created As String
    
    On Error Resume Next
    Set ws = Worksheets("NAMES")
    
    If ws.Cells(Rows.Count, 1).End(xlUp).Row < 2 Then
        MsgBox "There is no data to create defined names."
        Exit Sub
    End If
    
    Recuperar_Formulas = MsgBox("When creating defined names, do you also want to recover the formulas from the cells with the defined names?", vbYesNo)
    Set Rng = ws.Range("A2").Resize(ws.Cells(Rows.Count, 1).End(xlUp).Row - 1, 4)
    
    For Fila = 2 To Rng.Rows.Count + 1
        Name = Rng.Cells(Fila, 1).Value
        Reference = Rng.Cells(Fila, 2).Value
        Formula = Rng.Cells(Fila, 3).Value
        Tipo = Rng.Cells(Fila, 4).Value
        ActiveWorkbook.Names.Add Name:=Name, RefersTo:=Reference
        If Err.Number <> 0 Then
            Not_Created = Name & Newline & Name
            Err.Clear
        End If
        If Recuperar_Formulas = vbYes Then
            Set DefinedName = ActiveWorkbook.Names(Name)
            DefinedName.RefersToRange.Formula = Formula
        End If
        Next Fila
    If Not_Created <> "" Then MsgBox Not_Created
    Cells.Select
    Selection.WrapText = False
    MsgBox "END"
End Sub
Sub Save_Formulas_Values()
' ------------------------------------------------------------------- '
' --- Macro to save all formulas and values from a workbook --- '
' ------------------------------------------------------------------- '
    Dim ws As Worksheet
    Dim NewSheet As Worksheet
    Dim Celda As Range
    Dim FilaDestino As Long
    Dim ColActual As Long
    Dim FilActual As Long
    Dim N_Columns As Long
    Dim N_Files As Long
    Dim Saved_Formula As Byte

    Dim Sheet_F As String
    Sheet_F = "FORMULA"
    
    On Error Resume Next
    Set NewSheet = ActiveWorkbook.Sheets(Sheet_F)
    If NewSheet Is Nothing Then
        Set NewSheet = Worksheets.Add
        NewSheet.Name = Sheet_F
        NewSheet.Move After:=Sheets(Sheets.Count)
        Range("'" & Sheet_F & "'!B2").Select
        ActiveWindow.FreezePanes = True
    Else
        NewSheet.Cells.ClearContents
    End If
    
    If NewSheet.Cells(1, 1).Value = "" Then
        NewSheet.Cells(1, 1).Value = "Sheet"
        NewSheet.Cells(1, 2).Value = "Cell"
        NewSheet.Cells(1, 3).Value = "Formula"
        NewSheet.Cells(1, 4).Value = "Value"
        NewSheet.Cells(1, 5).Value = "Cell Name"
    End If
    
    FilaDestino = NewSheet.Cells(NewSheet.Rows.Count, 1).End(xlUp).Row + 1
    
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name <> NewSheet.Name And ws.Name <> "NAMES" And ws.Name <> "INSTRUCTIONS" Then
        Saved_Formula = MsgBox("Do you want to save the formulas of the " & ws.Name & " sheet?", vbYesNo)
            If Saved_Formula = vbYes Then
                N_Columns = ws.UsedRange.Columns.Count
                For ColActual = 1 To N_Columns
                    N_Files = ws.Cells(ws.Rows.Count, ColActual).End(xlUp).Row
                    Set Celda = ws.Cells(1, ColActual)
                    For FilActual = 1 To N_Files
                        If Celda.HasFormula Then
                            NewSheet.Cells(FilaDestino, 1).Value = ws.Name
                            NewSheet.Cells(FilaDestino, 2).Value = Celda.Address
                            NewSheet.Cells(FilaDestino, 3).Value = "'" & Celda.FormulaLocal
                            NewSheet.Cells(FilaDestino, 4).Value = ""
                            If Not Celda.Name Is Nothing Then
                                NewSheet.Cells(FilaDestino, 5).Value = Celda.Name.Name
                            End If
                            FilaDestino = FilaDestino + 1
                        ElseIf Not IsEmpty(Celda.Value) Then
                            NewSheet.Cells(FilaDestino, 1).Value = ws.Name
                            NewSheet.Cells(FilaDestino, 2).Value = Celda.Address
                            NewSheet.Cells(FilaDestino, 3).Value = ""
                            NewSheet.Cells(FilaDestino, 4).Value = Celda.Value
                            If Not Celda.Name Is Nothing Then
                                NewSheet.Cells(FilaDestino, 5).Value = Celda.Name.Name
                            End If
                                FilaDestino = FilaDestino + 1
                        End If
                        Set Celda = Celda.Offset(1, 0)
                    Next FilActual
                Next ColActual
            End If
        End If
    Next ws
    Cells.Select
    Selection.WrapText = False
    
    Columns("A:B").EntireColumn.AutoFit
    Columns("E:E").EntireColumn.AutoFit
    Columns("C:D").Select
    Selection.ColumnWidth = 50
    Range("B2").Select
    
    On Error GoTo 0
    MsgBox "END"
End Sub
Sub Last_Cell_per_SHEETS()
' ---------------------------------------------------------------------------------------- '
' --- Macro to determine the last cell on each sheet, in case it extends too much...   --- '
' ---------------------------------------------------------------------------------------- '
    Dim ws As Worksheet
    Dim N_Columnas As Long
    Dim N_Filas As Long
    Dim MENSAJE As String
    
    For Each ws In ActiveWorkbook.Sheets
        N_Columnas = ws.UsedRange.Columns.Count
        N_Filas = ws.UsedRange.Rows.Count

        MENSAJE = MENSAJE & vbNewLine & Cells(N_Filas, N_Columnas).Address & vbTab & "->" & _
                    vbTab & N_Columnas & " Columnas" & vbTab & " -> " & "Hoja  " & ws.Name
                    
    Next ws
    MsgBox MENSAJE, vbOKOnly, "FastTest PlugIn"
End Sub
Sub PRINT_PDF()
    Dim ws As Worksheet
    Dim numElements As Integer
    Dim bSuccess As Boolean
    Dim strPDFs() As String
    Dim PRINT_ALL As Integer
    Dim Data_Print As String
    Dim Column_Print As ListColumn
    Dim Cell_Print As Range
    Dim continue As Integer
    
    Set ws = ActiveWorkbook.Sheets([N_Sheet_Pb])
    Set MyTable = [Image].ListObject
    Set Column_Print = MyTable.ListColumns("PRINT")
    
    [Line_Break] = ""
    continue = MsgBox("Do you want to print all the problems GROUPED in a single PDF file?", vbYesNo)
    If [D_Fixed] = [_NO] Then
        numElements = [N_Pb_Random]
        If continue = vbYes Then
            ReDim strPDFs(1 To 2) As String
            For i = 1 To numElements
                strPDFs(2) = ActiveWorkbook.Path & "\TEMPORAL.pdf"
                ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(2), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                     IgnorePrintAreas:=False, OpenAfterPublish:=False
                If strPDFs(1) <> "" Then
                    bSuccess = MergePDFs(strPDFs, strPDFs(1))
                Else
                    strPDFs(1) = ActiveWorkbook.Path & "\Random - " & [File_Name] & " - " & [Results] & " (" & [numElements] & " pb).pdf"
                    ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(1), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                         IgnorePrintAreas:=False, OpenAfterPublish:=False
                End If
                Calculate
            Next i
            Kill strPDFs(2)
        Else
            ReDim strPDFs(1 To 1) As String
            For i = 1 To numElements
                strPDFs(1) = ActiveWorkbook.Path & "\" & [Name_Pb] & " - " & [Results] & ".pdf"
                ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(1), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                     IgnorePrintAreas:=False, OpenAfterPublish:=False
                Calculate
            Next i
        End If
    Else
        numElements = MyTable.ListRows.Count
        PRINT_ALL = MsgBox("Print all the problems from the Fixed Data table?", vbYesNo)
        If PRINT_ALL = vbYes Then [Print_OK] = [_OK] Else [Print_OK] = [_NO]
    
        If continue = vbYes Then
            ReDim strPDFs(1 To 2) As String
            If PRINT_ALL = vbYes Then
                [Name_Data] = MyTable.ListRows(1).Range.Cells(1, 1)
                strPDFs(1) = ActiveWorkbook.Path & "\" & [File_Name] & " - " & [Results] & " (" & numElements & " pb).pdf"
                ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(1), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                     IgnorePrintAreas:=False, OpenAfterPublish:=False
                i = 1
            Else
                i = 0
                If [D_Print] > 0 And [D_Fixed] = [_OK] Then continue = MsgBox("You have selected " & [D_Print] & " problem. " & vbNewLine & vbNewLine _
                            & "Are you sure you have selected ALL the rows of the data you want to print in the Fixed Data table?", vbYesNo)
                If continue = vbNo Or [D_Print] = 0 Then
                    MsgBox "You must mark with an X the data you want to print in the PRINT column."
                    Exit Sub
                End If
            End If
            For Each Cell_Print In Column_Print.DataBodyRange
                Data_Print = Cell_Print
                i = i + 1
                If (PRINT_ALL = vbYes And i <= numElements) Or (PRINT_ALL = vbNo And Data_Print <> "" And i <= numElements) Then
                    [Name_Data] = MyTable.ListRows(i).Range.Cells(1, 1)
                    strPDFs(2) = ActiveWorkbook.Path & "\TEMPORAL.pdf"
                    ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(2), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                         IgnorePrintAreas:=False, OpenAfterPublish:=False
                    If strPDFs(1) <> "" Then
                        bSuccess = MergePDFs(strPDFs, strPDFs(1))
                    Else
                        strPDFs(1) = ActiveWorkbook.Path & "\" & [File_Name] & " - " & [Results] & " (" & [D_Print] & " pb).pdf"
                        ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(1), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                             IgnorePrintAreas:=False, OpenAfterPublish:=False
                    End If
                End If
            Next Cell_Print
            Kill strPDFs(2)
        Else
            i = 0
            ReDim strPDFs(1 To 1) As String
            For Each Cell_Print In Column_Print.DataBodyRange
                Data_Print = Cell_Print
                i = i + 1
                If (PRINT_ALL = vbYes And i <= numElements) Or (PRINT_ALL = vbNo And Data_Print <> "" And i <= numElements) Then
                    [Name_Data] = MyTable.ListRows(i).Range.Cells(1, 1)
                    strPDFs(1) = ActiveWorkbook.Path & "\" & [Name_Pb] & " - " & [Results] & ".pdf"
                    ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(1), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                         IgnorePrintAreas:=False, OpenAfterPublish:=False
                End If
            Next Cell_Print
        End If
    End If
    MsgBox "Process completed, the files are in the folder:" & vbNewLine & ActiveWorkbook.Path, vbInformation
End Sub
Private Function MergePDFs(arrFiles() As String, strSaveAs As String) As Boolean
'---------------------------------------------------------------------------------------------------
'---PROGRAM: MergePDFs------------------------------------------------------------------------------
'---DEVELOPER: Ryan Wells (wellsr.com)--------------------------------------------------------------
'---DATE: 09/2017-----------------------------------------------------------------------------------
'---DESCRIPTION: This function uses Adobe Acrobat (won't work with just the Reader!) to-------------
'--- combine PDFs into one PDF and save the new PDF with its own file name.-------------
'---INPUT: The function requires two arguments.-----------------------------------------------------
'--- 1) arrFiles is an array of strings containing the full path to each PDF you want to------
'--- combine in the order you want them combined.------------------------------------------
'--- 2) strSaveAs is a string containing the full path you want to save the new PDF as.-------
'---REQUIREMENTS: 1) Must add a reference to "Adobe Acrobat X.0 Type Library" or "Acrobat"----------
'--- under Tools > References. This has been tested with Acrobat 6.0 and 10.0.------
'---CAUTION: This function won't work unless you have the full Adobe Acrobat. In other words,-------
' Adobe Reader will not work.------------------------------------------------------------
'---------------------------------------------------------------------------------------------------
 
Dim objCAcroPDDocDestination As Acrobat.CAcroPDDoc
Dim objCAcroPDDocSource As Acrobat.CAcroPDDoc
Dim i As Integer
Dim iFailed As Integer
 
On Error GoTo NoAcrobat:
'Initialize the Acrobat objects
Set objCAcroPDDocDestination = CreateObject("AcroExch.PDDoc")
Set objCAcroPDDocSource = CreateObject("AcroExch.PDDoc")
 
'Open Destination, all other documents will be added to this and saved with
'a new filename
objCAcroPDDocDestination.Open (arrFiles(LBound(arrFiles))) 'open the first file
 
'Open each subsequent PDF that you want to add to the original
  'Open the source document that will be added to the destination
    For i = LBound(arrFiles) + 1 To UBound(arrFiles)
        objCAcroPDDocSource.Open (arrFiles(i))
        If objCAcroPDDocDestination.InsertPages(objCAcroPDDocDestination.GetNumPages - 1, objCAcroPDDocSource, 0, objCAcroPDDocSource.GetNumPages, 0) Then
          MergePDFs = True
        Else
          'failed to merge one of the PDFs
          iFailed = iFailed + 1
        End If
        objCAcroPDDocSource.Close
    Next i
    
    objCAcroPDDocDestination.Save 1, strSaveAs 'Save it as a new name
    objCAcroPDDocDestination.Close
    Set objCAcroPDDocSource = Nothing
    Set objCAcroPDDocDestination = Nothing
     
NoAcrobat:
    If iFailed <> 0 Then
        MergePDFs = False
    End If
    On Error GoTo 0
End Function


