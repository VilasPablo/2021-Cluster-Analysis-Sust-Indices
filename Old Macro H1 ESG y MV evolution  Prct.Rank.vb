Public i As Long
Sub macro_controladora()

    For pagina = 1 To Worksheets.Count
    If Worksheets(pagina).Name <> "Clasificación índices" Then
    Worksheets(pagina).Cells.ClearContents
    Worksheets(pagina).Cells.ColumnWidth = 14
    Worksheets(pagina).Rows("1:1").RowHeight = 45
    With Worksheets(pagina).Cells
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Worksheets(pagina).Rows("1:1")
        .Font.Bold = True
    End With
    End If
    Next
    
   Dim archivos As String
   Dim contador1 As Long
   Dim y   'mi matriz con nombre de archivos
   Dim tipo_archivo As String
   Dim ruta As String
   Dim columnas_y As Integer
   Dim tiempo_espera As String
   Dim wbk
   Dim copia As Integer

tiempo_espera = "00:00:00"
columnas_y = 2
tipo_archivo = "xlsx" 'xlsx
ruta = ThisWorkbook.Path 'ThisWorkbook.Path puedes poner otra cualquiera
        
    archivos = Dir(ruta & "\*" & tipo_archivo & "*")
    While archivos <> ""
        contador1 = contador1 + 1
        archivos = Dir ' pasar al siguiente xlsx
    Wend
    '' redimensionar la matriz en función de los acrchivos xlsx existentes
    ReDim y(contador1 - 1, columnas_y)
    archivos = Dir(ruta & "\*" & tipo_archivo & "*")
    For Z = 0 To contador1 - 1
    
        y(Z, 0) = archivos
        y(Z, 1) = FileDateTime(ruta & "\" & archivos)
        archivos = Dir
    Next

For i = 0 To contador1 - 1 'variable definida como public
    'El "if" unicamente en caso de que nuestro libro sea xlsx y no xlsm
    copia = i
    If y(i, 0) = ThisWorkbook.Name Then
    Else
        Workbooks.Open (ruta & "\" & y(i, 0))
        Application.Wait (Now + TimeValue("" & tiempo_espera & "")) 'esperar para darle tiempo a que se abra
        
        Call Macro1

        'Macros

    i = copia
    End If
    
Next

End Sub
    





Sub Macro1()
'
' Macro1 Macro
'

'
   Dim xx
   Dim wi As Double
   Dim N As Long
   Dim column_sum As Integer
   Dim column_restric As Integer
   Dim column_restric2 As Integer
   Dim contador1 As Integer
   Dim total As Long
   Dim criterio As Long
   Dim criterio2 As String
   Dim index_name As String
   Dim año_inicio As Long
   Dim año_final As Long
   Dim nombre_hoja As String
   Dim final_for As Long
   Dim columna_letra As String
   
   
   nombre_hoja = "Estadísticos Indice"
   año_inicio = 2007
   año_final = 2017
   column_sum = 4
   column_restric = 8
   column_restric2 = 12
   criterio2 = "<>No Dato"
   columna_lertra = "H"
   
   
   contador1 = 1
   index_name = Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5)
   Sheets("" & index_name & "").Select
   total = Range("B1").CurrentRegion.Rows.Count
   final_for = (-año_inicio + año_final + 1) * 12
   

   ReDim xx(final_for, 8)

   
   For año = año_inicio To año_final
   For mes = 1 To 12
   If mes < 10 Then
   criterio = año & 0 & mes
   Else
   criterio = año & mes
   End If
   xx(contador1, 0) = criterio
   
   xx(contador1, 1) = Application.WorksheetFunction.SumIfs(Range(Cells(2, column_sum), Cells(total, column_sum)), _
   Range(Cells(2, column_restric), Cells(total, column_restric)), criterio) 'Peso controlado
   
   xx(contador1, 2) = Application.WorksheetFunction.CountIfs(Range(Cells(2, column_restric), Cells(total, column_restric)), _
   criterio) 'empresas índice
   
   xx(contador1, 3) = Application.WorksheetFunction.SumIfs(Range(Cells(2, column_sum), Cells(total, column_sum)), _
   Range(Cells(2, column_restric), Cells(total, column_restric)), criterio, Range(Cells(2, column_restric2), _
   Cells(total, column_restric2)), criterio2) 'peso con score
   
   xx(contador1, 4) = Application.WorksheetFunction.CountIfs(Range(Cells(2, column_restric), Cells(total, _
   column_restric)), criterio, Range(Cells(2, column_restric2), Cells(total, column_restric2)), criterio2) 'Empresas con datos --> MV, ESG scores
   
   'If xx(contador1, 3) > 0 Then
   'a = 1 / xx(contador1, 3)
   'Else: a = 1
   'End If
   
   xx(contador1, 5) = ActiveSheet.Evaluate("=AVERAGE(IF((" & columna_lertra & "2:" & columna_lertra & "" & total & "=" & criterio & ")*(L2:L" & total & "<>""No Dato""),L2:L" & total & "))")  ' * a
   
   xx(contador1, 6) = ActiveSheet.Evaluate("=AVERAGE(IF((" & columna_lertra & "2:" & columna_lertra & "" & total & "=" & criterio & ")*(M2:M" & total & "<>""No Dato""),M2:M" & total & "))") ' * a
   
   xx(contador1, 7) = ActiveSheet.Evaluate("=AVERAGE(IF((" & columna_lertra & "2:" & columna_lertra & "" & total & "=" & criterio & ")*(N2:N" & total & "<>""No Dato""),N2:N" & total & "))") ' * a
   
   xx(contador1, 8) = ActiveSheet.Evaluate("=AVERAGE(IF((" & columna_lertra & "2:" & columna_lertra & "" & total & "=" & criterio & ")*(O2:O" & total & "<>""No Dato""),O2:O" & total & "))") ' * a
   contador1 = contador1 + 1
    Next
    Next
    
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("" & nombre_hoja & "").Delete
    Application.DisplayAlerts = True
    Sheets.Add.Name = nombre_hoja
    Sheets("" & nombre_hoja & "").Select
    
   'Debe coincidir con el nombre de las hojas de la macro principal
   xx(0, 0) = "Fecha"
   xx(0, 1) = "Peso con WTIDX"
   xx(0, 2) = "Empresas en índice"
   xx(0, 3) = "WTIDX controlado"
   xx(0, 4) = "Nº Empresas con Dato"
   xx(0, 5) = "RP MV$"
   xx(0, 6) = "RP ESG"
   xx(0, 7) = "RP CombinedS"
   xx(0, 8) = "RP ControverS"
   
    For aaa = 0 To 8 'columna
    For kkk = 0 To final_for ' fila
    Cells(kkk + 1, aaa + 1).Value = Application.WorksheetFunction.Index(xx, kkk + 1, aaa + 1)
    Next
    Next


Call dejar_bonito
Cells.Interior.ColorIndex = 2

    Set wbk = ActiveWorkbook

 Cells.Interior.ColorIndex = 2
 
    ThisWorkbook.Activate
    wbk.Close (True)



    For columna = 1 To 8
        For fila = 1 To final_for
        Sheets("" & xx(0, columna) & "").Cells(fila + 1, 2 + i).Value = Application.WorksheetFunction.Index(xx, fila + 1, columna + 1)
        Next
    Sheets("" & xx(0, columna) & "").Cells(1, 2 + i).Value = index_name
    Next
    
    If i = 0 Then ' 2 = El útilo índice
        For columna = 1 To 8
            For fila = 1 To final_for
            Sheets("" & xx(0, columna) & "").Cells(fila + 1, 1).Value = Application.WorksheetFunction.Index(xx, fila + 1, 1)
            Next
    Next
    End If

    
   
  
End Sub




Sub dejar_bonito()

    ActiveSheet.Cells.ColumnWidth = 14
    ActiveSheet.Cells.RowHeight = 14
    ActiveSheet.Rows("1:1").RowHeight = 45
    Columns(1).EntireColumn.AutoFit
    Columns(3).EntireColumn.AutoFit
    With ActiveSheet.Cells
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With ActiveSheet.Rows("1:1")
        .Font.Bold = True
    End With
    Cells.EntireColumn.AutoFit
End Sub

    
