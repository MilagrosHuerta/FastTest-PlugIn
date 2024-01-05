Attribute VB_Name = "FTP_Problemas"
' ----------------------------------------------------------------------- '
' --- MACROS PARA UTILIZAR EN LA HOJA DE PROBLEMAS DE FastTest PlugIn --- '
' ---           CREADAS POR MILAGROS HUERTA GÓMEZ DE MERODIO          --- '
' ----------------------------------------------------------------------- '
Option Explicit
Dim MiTabla As ListObject
Dim NumeroColumna As Long
Dim NumeroFila As Long
Dim RutaGuardado As String
Dim TempFilePath As String
Dim TempFileName As String
Dim Ancho_Imagen As Integer
Dim Alto_Imagen As Integer
Dim MOSTRAR_RESULTADOS As String
Dim i As Long
Dim Alto_Fila As Long
Sub Imagenes_Guarda_Datos_Fijos()
' --------------------------------------------------------------------------------- '
' --- Macro para guardar las imágenes, en función de los Datos Fijos            --- '
' --- Al guardar las imagenes en el disco duro, si va rápido no las guarda bien --- '
' ---  MARCAR LA FILA Y PULSAR F5 PARA EJECUTAR DESPACIO  --- '
'       .Chart.Paste ' que está en la macro siguiente --- '
' --------------------------------------------------------------------------------- '
' Nombre hoja DATOS FIJOS definido como [nombre] por si alguien lo cambia
    Sheets([N_Hoja_Datos]).Select
' El NOMBRE de la TABLA a la que pertenece la celda IMAGEN, por si alguien cambia el nombre de la tabla
    Set MiTabla = [Imagen].ListObject
    NumeroColumna = [Imagen].Column ' Columna en la que están las imágenes
    NumeroFila = [Imagen].Row ' Fila INICIAL en la que están las imágenes
    [Ver_Datos] = [_OK]
    MOSTRAR_RESULTADOS = MsgBox("¿Mostrar los resultados en la imagen?", vbYesNo)
' Pregunta si mostrar o no los resultados en la IMAGEN
    If MOSTRAR_RESULTADOS = vbYes Then [Ver_Resultados] = [_OK] Else [Ver_Resultados] = [_NO]
' Bucle desde la primera fila de la tabla de DATOS hasta la úlitma
    For i = 1 To MiTabla.ListRows.Count  '+ NumeroFila To MiTabla.ListRows.Count + NumeroFila
    ' Cambia los datos del problema
        Sheets([N_Hoja_Pb]).Select    ' Nombre hoja PROBLEMA definido como [nombre] por si alguien lo cambia
        [Nombre_Datos] = MiTabla.ListRows(i).Range.Cells(1, 1)
    ' Este solo sirve si una figura suelta, no un conjunto de imagenes agrupadas
    '    ActiveSheet.ChartObjects([N_Figura]).Copy
    ' Selecciona la imagen del problema - Si es un conjunto de imágenes agrupadas...
            ActiveSheet.Shapes.Range(Array([N_Figura])).Select
        Selection.Copy
    ' Pega la Imagen en la fila correspondiente
        Sheets([N_Hoja_Datos]).Select
        Cells(i + NumeroFila, NumeroColumna).Select
        Alto_Fila = ActiveCell.EntireRow.Height        ' En pixeles
        ActiveSheet.Pictures.Paste.Select
        Selection.ShapeRange.Height = Alto_Fila * 0.95 ' Alto de la fila al 95%
        Selection.Placement = xlMove                   ' Mover pero no cambiar tamaño con celda
        Selection.ShapeRange.Name = MiTabla.ListRows(i).Range.Cells(1, 1)
' ----------------------------------------------------------- '
        Call Imagenes_Guarda_Archivo
' ----------------------------------------------------------- '
    Next i
End Sub
Sub Imagenes_Guarda_Archivo()
' ------------------------------------------------------------------------------------ '
' --- Esta macro es para guardar las imágenes generadas con datos en archivos jpg  --- '
' --- NO EJECUTAR SOLA SI NO SE HA SELECCIONADO UNA IMAGEN                         --- '
' ------------------------------------------------------------------------------------ '
Dim NombreArchivo As String
    ' El ancho y alto de la imagen, por 4, para pegar la imagen en CHART
        Ancho_Imagen = Selection.ShapeRange.Width * [Multiplo_Img] ' 2.5  ' Cambiar entre 2 y 4
        Alto_Imagen = Selection.ShapeRange.Height * [Multiplo_Img] ' 2.5  ' Cambiar entre 2 y 4
    ' Obtener la ruta del archivo actual y la Sub_Carpeta
        RutaGuardado = ActiveWorkbook.Path & "\" & [SUB_CARPETA] & "\"
    ' Si la ruta no existe, la crea
        If Dir(RutaGuardado, vbDirectory) = "" Then
            MkDir RutaGuardado
        End If
    ' Crear una ruta temporal y un nombre de archivo temporal único
        TempFilePath = Environ("TEMP") & "\"
        TempFileName = "TempImage" & Format(Now, "yyyyMMddhhmmss") & ".png"
    ' Guardar la imagen en el archivo
    ' El ancho va en función de la imágen
        With ActiveSheet.ChartObjects.Add(Left:=0, Width:=Ancho_Imagen, Top:=0, Height:=Alto_Imagen)
'----------------------------------------------------------- '
'---  MARCAR LA FILA Y PULSAR F5 PARA EJECUTAR DESPACIO  --- '
'----------------------------------------------------------- '
            .Chart.Paste   ' ------------------------------- '
'----------------------------------------------------------- '
'----------------------------------------------------------- '
            .Chart.Export TempFilePath & TempFileName, "PNG"
            .Delete     ' Borra el archivo que se ha creado en la hoja de Excel
        End With
    ' Mover el archivo temporal a la ubicación final
        NombreArchivo = [Nombre_Datos] & ".png"
        FileCopy TempFilePath & TempFileName, RutaGuardado & NombreArchivo
    ' Eliminar el archivo temporal
        Kill TempFilePath & TempFileName
End Sub
Sub Nombres_Definidos_Listar()
    Dim MiNombre As Name
    Dim ws As Worksheet
    Dim Hoja As String
    Dim Listar_Valores_Formulas As Long
    
    On Error Resume Next
    Hoja = "NOMBRES"
    ' Agregar los resultados a una nueva hoja y ponerla al final
    If Worksheets(Hoja) Is Nothing Then
        Set ws = Worksheets.Add
        ws.Name = Hoja
        ws.Move After:=Sheets(Sheets.Count)
    ' Para inmobilizar la fila superior
        Range("'" & Hoja & "'!B2").Select
        ActiveWindow.FreezePanes = True
    Else
        Worksheets(Hoja).Cells.ClearContents

    End If
' Pone el encabezado para cada columna
    Worksheets(Hoja).Select
    Range("A1") = "Nombre Definido"
    Range("B1") = "Referencia TEXTO"
    Range("C1") = "Fórmula TEXTO"
    Range("D1") = "Tipo"
    Listar_Valores_Formulas = MsgBox("¿Deseas ver el resultado de la fórmula? Esto puede generar referencias circulares, en cuyo caso sería recomendable borrarlas luego.", vbYesNo)
    If Listar_Valores_Formulas = vbYes Then
    ' Estas columnas, al ser fórmulas en una hoja diferente a la que se ha diseñado, pueden generar referencias circulares
        Range("E1") = "Referencia"
        Range("F1") = "Fórmula"
    End If
    
    i = 1
    For Each MiNombre In ActiveWorkbook.Names
        If Left(MiNombre.Name, 6) <> "_xlfn." Then
            i = i + 1
            Range("A" & i) = MiNombre.Name
            Range("B" & i) = "'" & MiNombre.RefersToLocal
            Range("C" & i) = "'" & IIf(MiNombre.RefersToRange.HasFormula, MiNombre.RefersToRange.FormulaLocal, MiNombre.RefersToRange.Value)  'MiNombre.Formula
            Range("D" & i) = TypeName(MiNombre.RefersToRange)
            If Listar_Valores_Formulas = vbYes Then
                ' Estas columnas, al ser fórmulas en una hoja diferente a la que se ha diseñado, pueden generar referencias circulares
                Range("E" & i) = MiNombre.RefersToLocal
                Range("F" & i) = IIf(MiNombre.RefersToRange.HasFormula, MiNombre.RefersToRange.Formula, "")  'MiNombre.Formula
            End If
        End If
    Next MiNombre
' Para quitar AJUSTAR TEXTO
    Cells.Select
    Selection.WrapText = False
    ActiveWindow.DisplayHeadings = False  ' No muestra los encabezados
    Cells(2, 1).Select
    
    Columns("A:A").EntireColumn.AutoFit
    Columns("D:D").EntireColumn.AutoFit
    Columns("B:C").Select
    Selection.ColumnWidth = 50
    Range("B2").Select
    MsgBox "FIN"
End Sub
Sub Nombres_Definidos_Crear_Desde_Hoja()
' Esta macro es para recuperar los nombres que se tengan en la hoja, por si se hubieran perdido
    Dim ws As Worksheet
    Dim Rng As Range
    Dim Fila As Long
    Dim Nombre As String
    Dim Referencia As String
    Dim Tipo As String
    Dim Formula As String
    Dim NombreDefinido As Name
    Dim No_Creados As String
    
    On Error Resume Next
    ' Cambia "Nombres Definidos" al nombre de tu hoja
    Set ws = Worksheets("NOMBRES")
    
    ' Asegurarse de que hay datos en la columna A
    If ws.Cells(Rows.Count, 1).End(xlUp).Row < 2 Then
        MsgBox "No hay datos para crear nombres definidos."
        Exit Sub
    End If
    
    Recuperar_Formulas = MsgBox("Al crear los nombres definidos, ¿deseas también recuperar las fórmulas de las celdas con los nombres definidos?", vbYesNo)
    Set Rng = ws.Range("A2").Resize(ws.Cells(Rows.Count, 1).End(xlUp).Row - 1, 4)
    
    For Fila = 2 To Rng.Rows.Count + 1
        Nombre = Rng.Cells(Fila, 1).Value
        Referencia = Rng.Cells(Fila, 2).Value
        Formula = Rng.Cells(Fila, 3).Value
        Tipo = Rng.Cells(Fila, 4).Value
        ActiveWorkbook.Names.Add Name:=Nombre, RefersTo:=Referencia ', Type:=Evaluate("xl" & Tipo)
        If Err.Number <> 0 Then
            No_Creados = Nombre & Newline & Nombre
            Err.Clear
        End If
        If Recuperar_Formulas = vbYes Then
            Set NombreDefinido = ActiveWorkbook.Names(Nombre) ' Cambia al nombre del nombre definido
            NombreDefinido.RefersToRange.Formula = Formula
        End If
        Next Fila
    If No_Creados <> "" Then MsgBox No_Creados
    
' Para quitar AJUSTAR TEXTO selecciona todas las celdas
    Cells.Select
    Selection.WrapText = False
    MsgBox "FIN"
End Sub
Sub Pegar_Resultados()
' ---------------------------------------------------------- '
' --- Macro para pegar los resultados de los DATOS FIJOS --- '
' ---------------------------------------------------------- '
Dim Colu_Ini As Integer
Dim Colu_Fin As Integer
Dim Fila_Ini As Integer

    Set MiTabla = [Imagen].ListObject

    Colu_Ini = [Solu_ini].Column        ' Columna en la que está la primera solución
    Colu_Fin = [Imagen].Column - 1      ' Columna en la que está la última solución
    Fila_Ini = [Solu_ini].Row           ' Fila en la que están los datos (Los encabezados)

    ' Asegura que la opción de DATOS FIJOS es que sí, para poder pegar los valores
    [D_Fijos] = [_OK]
    
    For i = 1 To MiTabla.ListRows.Count
    ' Cambia los datos del problema, para pegar las soluciones una a una
        [Nombre_Datos] = MiTabla.ListRows(i).Range.Cells(1, 1)
        ActiveSheet.Range(Cells(Fila_Ini, Colu_Ini), Cells(Fila_Ini, Colu_Fin)).Select
        Selection.Copy
        MiTabla.ListRows(i).Range.Cells(1, Colu_Ini).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    Next i
    MiTabla.ListRows(1).Range.Cells(1, 2).Select
    MsgBox "FIN"
End Sub
Sub Guardar_Formulas_Valores()
' ------------------------------------------------------------------- '
' --- Macro para guardar todas las fórmulas y valores de un libro --- '
' ------------------------------------------------------------------- '
    Dim ws As Worksheet
    Dim NuevaHoja As Worksheet
    Dim Celda As Range
    Dim FilaDestino As Long
    Dim ColActual As Long
    Dim FilActual As Long
    Dim N_Columnas As Long
    Dim N_Filas As Long
    Dim Guardar_Formulas As Byte

    Dim Hoja As String
    Hoja = "FORMULAS"
    ' Agregar los resultados a una nueva hoja y ponerla al final
    
    On Error Resume Next
    Set NuevaHoja = ActiveWorkbook.Sheets(Hoja)
    ' Verificar si ya existe una hoja con el nombre "FORMULAS LIBRO" en el libro activo
    If NuevaHoja Is Nothing Then
        Set NuevaHoja = Worksheets.Add
        NuevaHoja.Name = Hoja
        NuevaHoja.Move After:=Sheets(Sheets.Count)
    ' Para inmobilizar la fila superior
        Range("'" & Hoja & "'!B2").Select
        ActiveWindow.FreezePanes = True
    Else
        NuevaHoja.Cells.ClearContents
    End If
    
    ' Agregar encabezados si es necesario
    If NuevaHoja.Cells(1, 1).Value = "" Then
        NuevaHoja.Cells(1, 1).Value = "Hoja"
        NuevaHoja.Cells(1, 2).Value = "Celda"
        NuevaHoja.Cells(1, 3).Value = "Fórmula"
        NuevaHoja.Cells(1, 4).Value = "Valor"
        NuevaHoja.Cells(1, 5).Value = "Nombre Celda"
    End If
    
    FilaDestino = NuevaHoja.Cells(NuevaHoja.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Recorrer todas las hojas del libro activo
    For Each ws In ActiveWorkbook.Sheets
        If ws.Name <> NuevaHoja.Name And ws.Name <> "NOMBRES" And ws.Name <> "INSTRUCCIONES" Then
        Guardar_Formulas = MsgBox("¿Quieres guardar las fórmulas de la hoja " & ws.Name & " ?", vbYesNo)
            If Guardar_Formulas = vbYes Then
                ' Recorrer las celdas de la primera columna, luego las de la segunda, etc.
                N_Columnas = ws.UsedRange.Columns.Count
                For ColActual = 1 To N_Columnas 'ws.UsedRange.Columns.Count
                    N_Filas = ws.Cells(ws.Rows.Count, ColActual).End(xlUp).Row
                    Set Celda = ws.Cells(1, ColActual)
                    For FilActual = 1 To N_Filas
                        If Celda.HasFormula Then
                            ' Agregar detalles de la celda y la fórmula en la nueva hoja
                            NuevaHoja.Cells(FilaDestino, 1).Value = ws.Name
                            NuevaHoja.Cells(FilaDestino, 2).Value = Celda.Address
                            NuevaHoja.Cells(FilaDestino, 3).Value = "'" & Celda.FormulaLocal
                            NuevaHoja.Cells(FilaDestino, 4).Value = ""
                            If Not Celda.Name Is Nothing Then
                                NuevaHoja.Cells(FilaDestino, 5).Value = Celda.Name.Name
                            End If
                            FilaDestino = FilaDestino + 1
                        ElseIf Not IsEmpty(Celda.Value) Then
                            ' Agregar detalles de la celda y el valor en la nueva hoja
                            NuevaHoja.Cells(FilaDestino, 1).Value = ws.Name
                            NuevaHoja.Cells(FilaDestino, 2).Value = Celda.Address
                            NuevaHoja.Cells(FilaDestino, 3).Value = ""
                            NuevaHoja.Cells(FilaDestino, 4).Value = Celda.Value
                            If Not Celda.Name Is Nothing Then
                                NuevaHoja.Cells(FilaDestino, 5).Value = Celda.Name.Name
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
    MsgBox "FIN"
End Sub
Sub Ultima_Celda_por_HOJAS()
' ---------------------------------------------------------------------------------------- '
' --- Macro para ver cual es la última celda de cada hoja, por si se extiende mucho... --- '
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
Sub IMPRIMIR_PDF()
    Dim ws As Worksheet
    Dim numElementos As Integer
    Dim bSuccess As Boolean
    Dim strPDFs() As String
    Dim IMPRIMIR_TODO As Integer
    Dim Dato_Imprimir As String
    Dim Columna_Imprimir As ListColumn
    Dim Celda_Imprimir As Range
    Dim continuar As Integer
    
    Set ws = ActiveWorkbook.Sheets([N_Hoja_Pb])
    Set MiTabla = [Imagen].ListObject
    Set Columna_Imprimir = MiTabla.ListColumns("IMPRIMIR")
    
    [Salto] = ""
    continuar = MsgBox("¿Quieres imprimir los problemas AGRUPADOS en un solo archivo PDF?", vbYesNo)
    If [D_Fijos] = [_NO] Then
        numElementos = [N_Pb_Ale]
        If continuar = vbYes Then
            ReDim strPDFs(1 To 2) As String
            For i = 1 To numElementos
                strPDFs(2) = ActiveWorkbook.Path & "\TEMPORAL.pdf"
                ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(2), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                     IgnorePrintAreas:=False, OpenAfterPublish:=False
                If strPDFs(1) <> "" Then
                    bSuccess = MergePDFs(strPDFs, strPDFs(1))
                Else
                    strPDFs(1) = ActiveWorkbook.Path & "\Aleatorio - " & [Nombre_Archivo] & " - " & [Resultados] & " (" & [numElementos] & " pb).pdf"
                    ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(1), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                         IgnorePrintAreas:=False, OpenAfterPublish:=False
                End If
                Calculate
            Next i
            Kill strPDFs(2)
        Else
            ReDim strPDFs(1 To 1) As String
            For i = 1 To numElementos
                strPDFs(1) = ActiveWorkbook.Path & "\" & [Nombre_Pb] & " - " & [Resultados] & ".pdf"
                ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(1), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                     IgnorePrintAreas:=False, OpenAfterPublish:=False
                Calculate
            Next i
        End If
    Else
        numElementos = MiTabla.ListRows.Count
        IMPRIMIR_TODO = MsgBox("¿Imprimir todos los problemas de la tabla de Datos Fijos?", vbYesNo)
        If IMPRIMIR_TODO = vbYes Then [Imp_OK] = [_OK] Else [Imp_OK] = [_NO]

' Bucle desde la primera fila de la tabla de DATOS hasta la úlitma
' Recorrer la columna "NOMBRE" de la tabla y genera los PDFs
        If continuar = vbYes Then
            ReDim strPDFs(1 To 2) As String
            If IMPRIMIR_TODO = vbYes Then
                [Nombre_Datos] = MiTabla.ListRows(1).Range.Cells(1, 1)
                strPDFs(1) = ActiveWorkbook.Path & "\" & [Nombre_Archivo] & " - " & [Resultados] & " (" & numElementos & " pb).pdf"
                ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(1), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                     IgnorePrintAreas:=False, OpenAfterPublish:=False
                i = 1
            Else
                i = 0
                If [D_Imprimir] > 0 And [D_Fijos] = [_OK] Then continuar = MsgBox("Has marcado " & [D_Imprimir] & " problema. " & vbNewLine & vbNewLine _
                            & "¿Estás seguro que has marcado TODAS las filas de los datos que quieres imprimir en la tabla de Datos Fijos?", vbYesNo)
                If continuar = vbNo Or ([D_Imprimir] = 0 And [D_Fijos] = [_OK]) Then
                    MsgBox "Debes marcar con una X los datos que deseas imprimir, en la columna IMPRIMIR."
                    Exit Sub
                End If
            End If
        
            For Each Celda_Imprimir In Columna_Imprimir.DataBodyRange
                Dato_Imprimir = Celda_Imprimir
                i = i + 1
                If (IMPRIMIR_TODO = vbYes And i <= numElementos) Or (IMPRIMIR_TODO = vbNo And Dato_Imprimir <> "" And i <= numElementos) Then
                    [Nombre_Datos] = MiTabla.ListRows(i).Range.Cells(1, 1)
                    strPDFs(2) = ActiveWorkbook.Path & "\TEMPORAL.pdf"
                    ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(2), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                         IgnorePrintAreas:=False, OpenAfterPublish:=False
                    If strPDFs(1) <> "" Then
                        bSuccess = MergePDFs(strPDFs, strPDFs(1))
                    Else
                        strPDFs(1) = ActiveWorkbook.Path & "\" & [Nombre_Archivo] & " - " & [Resultados] & " (" & [D_Imprimir] & " pb).pdf"
                        ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(1), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                             IgnorePrintAreas:=False, OpenAfterPublish:=False
                    End If
                End If
            Next Celda_Imprimir
            Kill strPDFs(2)
        Else
            i = 0
            ReDim strPDFs(1 To 1) As String
            For Each Celda_Imprimir In Columna_Imprimir.DataBodyRange
                Dato_Imprimir = Celda_Imprimir
                i = i + 1
                If (IMPRIMIR_TODO = vbYes And i <= numElementos) Or (IMPRIMIR_TODO = vbNo And Dato_Imprimir <> "" And i <= numElementos) Then
                    [Nombre_Datos] = MiTabla.ListRows(i).Range.Cells(1, 1)
                    strPDFs(1) = ActiveWorkbook.Path & "\" & [Nombre_Pb] & " - " & [Resultados] & ".pdf"
                    ws.ExportAsFixedFormat Type:=xlTypePDF, FileName:=strPDFs(1), Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                         IgnorePrintAreas:=False, OpenAfterPublish:=False
                End If
            Next Celda_Imprimir
        End If
    End If
    MsgBox "Proceso terminado, los archivos están en la carpeta:" & vbNewLine & vbNewLine & ActiveWorkbook.Path, vbInformation
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
