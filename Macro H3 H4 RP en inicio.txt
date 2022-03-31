Public copia As Long
Public i As Long

Sub macro_controladora()

   Dim archivos As String
   Dim contador1 As Long
   Dim y   'mi matriz con nombre de archivos
   Dim tipo_archivo As String
   Dim ruta As String
   Dim columnas_y As Integer
   Dim tiempo_espera As String
   Dim wbk

tiempo_espera = "00:00:10"
columnas_y = 2
tipo_archivo = "xlsx" 'xlsx
ruta = ThisWorkbook.Path 'ThisWorkbook.Path puedes poner otra cualquiera

Worksheets("Listado Scores y MV").Range("Z1:XFD1048576").ClearContents
Worksheets("Hipótesis 3").Range("B1:XFD1048576").ClearContents
Worksheets("Hipótesis 4").Range("B1:XFD1048576").ClearContents
Worksheets("Hipótesis 3").Range("B1:XFD1048576").Interior.ColorIndex = 2
Worksheets("Hipótesis 4").Range("B1:XFD1048576").Interior.ColorIndex = 2

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
        
        Call identificador_empresa
        Call asignar_scores
        Call Entradas_Salidas_Mantenimientos
        Call Crear_hoja_datos_filtrados
        Call Rango_percentil_macro
        Call entradas_universo
        Call Rango_percentil_macro_universo_entradas
        'Macros
Set wbk = ActiveWorkbook
    ThisWorkbook.Activate
    wbk.Close (True)
    Application.Wait (Now + TimeValue("" & tiempo_espera & ""))
    i = copia
    End If
    
Next

End Sub

 
Sub identificador_empresa()
'Columns(11).Delete
'Columns(10).Delete
'Columns(8).Delete
'Columns(7).Delete
'Columns(6).Delete
   'Variables que se definen solas
    Dim nº_existencia As Long
    Dim Valor_buscado As Variant 'valor que sirvirá de coincidencia para buscar un valor
    Dim rango_donde_buscar As Range ' normalmente esta variable se define en función del resto (no siempre tiene pk ser así)
    Dim Nº_filas_donde_buscar As Long 'largura de la matriz dónde buscar coincidencia (filas)
    Dim todo As Long 'Largura de la matriz de los datos que vamos a buscar (filas)
    
    'Variables que necesitan la intervención de un humano
 Dim hoja_name As String 'Nombre de la hoja de los datos a buscar
 Dim libro_name As String ' Nombre del libro donde están los datos a buscar
 Dim columna_receptora As Integer 'Columna donde se pondrán los datos a buscar
 Dim nombre_columna_receptora As String 'Nombre de la columna donde se pondrán los datos que queremos traer
 Dim matriz_con_datos_buscar As Long 'Largura de la matriz de los datos que vamos a buscar (filas)
 Dim valor_no_encontrado As String ' que poner sino encontramos el valor buscado
 Dim columna_de_busqueda As Integer 'columna de los datos a buscar
 Dim columna_con_fecha As Integer   'columna_con_fecha en índice

 Dim Nº_filas_donde_buscar_coincidencia As Long 'largura de la matriz dónde buscar coincidencia (filas)
 Dim columna_datos_devueltos As Integer 'Nº columna donde están los datos que buscamos
 Dim columna_datos_buscados As Integer 'Nº de columna donde están los datos a llevar  nuestra columna receptora
 Dim columna_final_datos_buscados As Integer 'Nº de columna final donde están los datos a llevar  nuestra columna receptora
 Dim columna_inicial_rango_donde_buscar As Integer 'columna donde se va a buscar la coincidencia de valores buscados
 Dim columna_final_rango_donde_buscar As Integer
 Dim hoja_name_coincidencia As String 'Nombre de hoja donde está matriz de coincidencia
 
columna_receptora = 10                   'Número de la columna donde se pondrán los datos que queremos traer
columna_de_busqueda = 5                 'Nº columna donde están los datos buscar y que deberán coincidir con los valores de columna_inicial_rango_donde_buscar
matriz_con_datos_buscar = 0             'menor que 1 se cálcula automáticamente (Largura de la matriz de los datos que vamos a buscar)
nombre_columna_receptora = "Código identificador de empresa"
valor_no_encontrado = "No Empresa"      'Valor a devolver en caso de no existir el valor buscado
hoja_name = Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5) 'Nombre del libro donde se pondrán los valores que traemos
libro_name = Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5) 'Nombre de la hoja donde se pondrán los valores que traemos
columna_con_fecha = 8

Nº_filas_donde_buscar_coincidencia = 0 'menor que 1 se cálcula automáticamente (Largura de la matriz datos donde buscar coincidencia)
columna_inicial_rango_donde_buscar = 2 'colum inicio donde se busca la coincidencia (inicio del rango de busqueda)
columna_final_rango_donde_buscar = 8   'colum final del rango donde se encuentran los datos del rango de busqueda
columna_datos_devueltos = 8            'columna dentro del rango de busqueda donde están los datos a devolver(Nº FILA)
columna_datos_buscados = 7             'columna donde están los datos que queremos llevarnos (puede coincidir con la anterior)
columna_final_datos_buscados = 7       'columna final donde están los datos que queremos llevarnos
hoja_name_coincidencia = "Info Títulos" 'Nombre de hoja donde está matriz de coincidencia

    'YA ESTAN TODAS LAS VARIABLES DEFINIDAS
    
        Sheets("" & hoja_name & "").Select
        Cells.Interior.ColorIndex = 2
            If matriz_con_datos_buscar < 1 Then
            todo = WorksheetFunction.CountIf(Range("b1:b" & Range("B2").CurrentRegion.Rows.Count), "<>""")
            Else
            todo = matriz_con_datos_buscar
            End If
        Cells(1, columna_receptora).Value = nombre_columna_receptora
            Set wbk = ActiveWorkbook
            ThisWorkbook.Activate
            Sheets("" & hoja_name_coincidencia & "").Select
            If Nº_filas_donde_buscar_coincidencia < 1 Then
            Nº_filas_donde_buscar = Range("b1").CurrentRegion.Rows.Count
            Else
            Nº_filas_donde_buscar = Nº_filas_donde_buscar_coincidencia
            End If
        Sheets("" & hoja_name_coincidencia & "").Select
        Set rango_donde_buscar = Range(Cells(2, columna_inicial_rango_donde_buscar) _
        , Cells(Nº_filas_donde_buscar, columna_final_rango_donde_buscar))
        
            ' Se asigna registro único numérico a cada empresa
For ii = 2 To todo
    Valor_buscado = Workbooks("" & libro_name & "").Sheets("" & hoja_name & "") _
    .Cells(ii, columna_de_busqueda) 'columna con ISIN
    nº_existencia = (Application.VLookup(Valor_buscado, _
    rango_donde_buscar, columna_datos_devueltos - columna_inicial_rango_donde_buscar + 1, True))
    
        If Cells(nº_existencia, columna_inicial_rango_donde_buscar) = Valor_buscado Then
        For X1 = columna_datos_buscados To columna_final_datos_buscados
            Workbooks("" & libro_name & "").Sheets("" & hoja_name & "") _
            .Cells(ii, columna_receptora + X1 - columna_datos_buscados).Value = Cells(nº_existencia, X1)
            Workbooks("" & libro_name & "").Sheets("" & hoja_name & "") _
            .Cells(ii, columna_receptora + 1).Value = Cells(nº_existencia, X1) & _
            Workbooks("" & libro_name & "").Sheets("" & hoja_name & "").Cells(ii, columna_con_fecha)
            Next
        Else
        For x2 = columna_datos_buscados To columna_final_datos_buscados + 1
            Workbooks("" & libro_name & "").Sheets("" & hoja_name & "") _
            .Cells(ii, columna_receptora + x2 - columna_datos_buscados).Value = valor_no_encontrado
            Next
        End If
Next
wbk.Activate
Cells(1, 11).Value = "Código Identificador mensual"
Call dejar_bonito
End Sub



Sub asignar_scores()


   'Variables que se definen solas
    Dim nº_existencia As Long
    Dim Valor_buscado As Variant 'valor que sirvirá de coincidencia para buscar un valor
    Dim rango_donde_buscar As Range ' normalmente esta variable se define en función del resto (no siempre tiene pk ser así)
    Dim Nº_filas_donde_buscar As Long 'largura de la matriz dónde buscar coincidencia (filas)
    Dim todo As Long 'Largura de la matriz de los datos que vamos a buscar (filas)
    Dim criterio_filtrado As String
    
    'Variables que necesitan la intervención de un humano
 Dim hoja_name As String 'Nombre de la hoja de los datos a buscar
 Dim libro_name As String ' Nombre del libro donde están los datos a buscar
 Dim columna_receptora As Integer 'Columna donde se pondrán los datos a buscar
 Dim nombre_columna_receptora As String 'Nombre de la columna donde se pondrán los datos que queremos traer
 Dim matriz_con_datos_buscar As Long 'Largura de la matriz de los datos que vamos a buscar (filas)
 Dim valor_no_encontrado As String ' que poner sino encontramos el valor buscado
 Dim columna_de_busqueda As Integer 'columna de los datos a buscar

 Dim Nº_filas_donde_buscar_coincidencia As Long 'largura de la matriz dónde buscar coincidencia (filas)
 Dim columna_datos_devueltos As Integer 'Nº columna donde están los datos que buscamos
 Dim columna_datos_buscados As Integer 'Nº de columna donde están los datos a llevar  nuestra columna receptora
 Dim columna_final_datos_buscados As Integer 'Nº de columna final donde están los datos a llevar  nuestra columna receptora
 Dim columna_inicial_rango_donde_buscar As Integer 'columna donde se va a buscar la coincidencia de valores buscados
 Dim columna_final_rango_donde_buscar As Integer
 Dim hoja_name_coincidencia As String 'Nombre de hoja donde está matriz de coincidencia
 
 hoja_name = Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5)
 Set rng2 = ThisWorkbook.Worksheets("Clasificación índices").Range("F1:F" & 50).Find("" & hoja_name & "") 'zona geográfica índice si es global los datos se cogen de otras columnas
 criterio_filtrado = rng2.Offset(0, 1)
    If criterio_filtrado = "Global" Then
    colc = 4
    Else
    colc = 0
    End If
 
columna_receptora = 12                   'Número de la columna donde se pondrán los datos que queremos traer
columna_de_busqueda = 11                 'Nº columna donde están los datos buscar y que deberán coincidir con los valores de columna_inicial_rango_donde_buscar
matriz_con_datos_buscar = 0             'menor que 1 se cálcula automáticamente (Largura de la matriz de los datos que vamos a buscar)
nombre_columna_receptora = "Scores"
valor_no_encontrado = "No Dato"      'Valor a devolver en caso de no existir el valor buscado
hoja_name = Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5) 'Nombre del libro donde se pondrán los valores que traemos
libro_name = Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5) 'Nombre de la hoja donde se pondrán los valores que traemos

Nº_filas_donde_buscar_coincidencia = 0 'menor que 1 se cálcula automáticamente (Largura de la matriz datos donde buscar coincidencia)
columna_inicial_rango_donde_buscar = 7 'colum inicio donde se busca la coincidencia (inicio del rango de busqueda)
columna_final_rango_donde_buscar = 8   'colum final del rango donde se encuentran los datos del rango de busqueda
columna_datos_devueltos = 8            'columna dentro del rango de busqueda donde están los datos a devolver
columna_datos_buscados = 10 + colc     'columna donde están los datos que queremos llevarnos (puede coincidir con la anterior)
columna_final_datos_buscados = 13 + colc 'columna final donde están los datos que queremos llevarnos
hoja_name_coincidencia = "Listado Scores y MV" 'Nombre de hoja donde está matriz de coincidencia

    'YA ESTAN TODAS LAS VARIABLES DEFINIDAS
    
        Sheets("" & hoja_name & "").Select
            If matriz_con_datos_buscar < 1 Then
            todo = WorksheetFunction.CountIf(Range("b1:b" & Range("B2").CurrentRegion.Rows.Count), "<>""")
            Else
            todo = matriz_con_datos_buscar
            End If
        Columns(16).ClearContents
        Cells(1, columna_receptora).Value = nombre_columna_receptora
            Set wbk = ActiveWorkbook
            ThisWorkbook.Activate
            Sheets("" & hoja_name_coincidencia & "").Select
            If Nº_filas_donde_buscar_coincidencia < 1 Then
            Nº_filas_donde_buscar = Range("b1").CurrentRegion.Rows.Count
            Else
            Nº_filas_donde_buscar = Nº_filas_donde_buscar_coincidencia
            End If
        Sheets("" & hoja_name_coincidencia & "").Select
        Set rango_donde_buscar = Range(Cells(2, columna_inicial_rango_donde_buscar) _
        , Cells(Nº_filas_donde_buscar, columna_final_rango_donde_buscar))
        
            ' Se asigna registro único numérico a cada empresa
For ii = 2 To todo
    Valor_buscado = Workbooks("" & libro_name & "").Sheets("" & hoja_name & "") _
    .Cells(ii, columna_de_busqueda)
    If Valor_buscado = "No Empresa" Then GoTo line2
    nº_existencia = (Application.VLookup(Valor_buscado, _
    rango_donde_buscar, columna_datos_devueltos - columna_inicial_rango_donde_buscar + 1, True))
    
        If Cells(nº_existencia, columna_inicial_rango_donde_buscar) = Valor_buscado Then
        For X1 = columna_datos_buscados To columna_final_datos_buscados
        
        If Cells(nº_existencia, 25) <> criterio_filtrado And colc = 0 Then ' Descartamos la información de aquellos títulos en los que no coincide la zona geográfica con la del índice
           Workbooks("" & libro_name & "").Sheets("" & hoja_name & "") _
            .Cells(ii, columna_receptora + X1 - columna_datos_buscados).Value = "Zona Geo Inadecuada"
        Else
            Workbooks("" & libro_name & "").Sheets("" & hoja_name & "") _
            .Cells(ii, columna_receptora + X1 - columna_datos_buscados).Value = Cells(nº_existencia, X1)
        End If
            Next
        Workbooks("" & libro_name & "").Sheets("" & hoja_name & "").Cells(ii, 16).Value = nº_existencia
        Else
line2:
        For x2 = columna_datos_buscados To columna_final_datos_buscados
            Workbooks("" & libro_name & "").Sheets("" & hoja_name & "") _
            .Cells(ii, columna_receptora + x2 - columna_datos_buscados).Value = valor_no_encontrado
            Next
        Workbooks("" & libro_name & "").Sheets("" & hoja_name & "").Cells(ii, 16).Value = "No controlado"
        End If
 

Next
For X4 = columna_datos_buscados To columna_final_datos_buscados
    Workbooks("" & libro_name & "").Sheets("" & hoja_name & "") _
    .Cells(1, columna_receptora + X4 - columna_datos_buscados).Value = Cells(1, X4)
    Next
    
Workbooks("" & libro_name & "").Sheets("" & hoja_name & "").Cells(1, 16).Value = "Nº Fila (clave)"
wbk.Activate



End Sub


Sub Entradas_Salidas_Mantenimientos()
'
' Macro1 Macro
'

'

   Dim Valores_enT As Long
   Dim valores_enT1 As Long
   Dim todos_valores As Long
   Dim columna_isin As Integer
   Dim columna_fecha As Integer
   Dim columna_mantiene As Integer
   Dim columna_sale As Integer
   Dim columna_entra As Integer
   Dim columna_duplicado

columna_isin = 5
columna_fecha = 8
columna_mantiene = 17
columna_sale = 18
columna_entra = 19
columna_duplicado = 20

Columns(columna_mantiene).ClearContents
Columns(columna_sale).ClearContents
Columns(columna_entra).ClearContents
Columns(columna_duplicado).ClearContents

todos_valores = Range("B1").CurrentRegion.Rows.Count
   
For i = 2 To todos_valores
    Valores_enT = Application.WorksheetFunction.CountIfs(Range(Cells(i, columna_fecha), _
    Cells(todos_valores, columna_fecha)), Cells(i, columna_fecha))
    valores_enT1 = Application.WorksheetFunction.CountIfs(Range(Cells(i + Valores_enT _
    , columna_fecha), Cells(todos_valores, columna_fecha)), Cells(i + Valores_enT, _
    columna_fecha))
    
    'Si hay menos de dos empresas en índice que no haga nada y pase al siguiente mes
   If valores_enT1 < 2 Or Valores_enT < 2 Or Cells(i + Valores_enT, columna_fecha) = "" Then GoTo line2
   
    'Localizar los que se mantienen o salen
    For t1 = 0 To Valores_enT - 1
        If Application.WorksheetFunction.CountIf(Range(Cells(i + Valores_enT, columna_isin), _
            Cells(i + Valores_enT + valores_enT1 - 1, columna_isin)), Cells(i + t1, columna_isin)) = 1 Then
            Cells(i + t1, columna_mantiene).Value = "Se mantiene"
        Else
               If Application.WorksheetFunction.CountIf(Range(Cells(i + Valores_enT, columna_isin), _
                   Cells(i + Valores_enT + valores_enT1 - 1, columna_isin)), Cells(i + t1, columna_isin)) = 0 Then
                   Cells(i + t1, columna_sale).Value = "Sale"
               Else
            'Localizar duplicados
               If Application.WorksheetFunction.CountIf(Range(Cells(i, columna_isin), _
                   Cells(i + Valores_enT - 1, columna_isin)), Cells(i + t1, columna_isin)) > 1 And Cells(i + t1, columna_isin) <> "NA" Then
                       Cells(i + t1, columna_duplicado).Value = "Duplicado"
               End If
               End If
        End If

    Next
    'Localizar los que entran
    For t2 = 0 To valores_enT1 - 1
        If Application.WorksheetFunction.CountIf(Range(Cells(i, columna_isin), _
            Cells(i + Valores_enT - 1, columna_isin)), Cells(i + t2 + Valores_enT, columna_isin)) = 0 Then
            Cells(i + t2 + Valores_enT, columna_entra).Value = "Entra"
        Else
        End If
    Next
    'pasamos al siguiente mes
    i = Valores_enT + i + -1
line2:
Next
Cells(1, columna_entra).Value = "Entra"
Cells(1, columna_sale).Value = "Sale"
Cells(1, columna_mantiene).Value = "Se mantiene"
Cells(1, columna_duplicado).Value = "Duplicado"

Call dejar_bonito

End Sub


Sub Crear_hoja_datos_filtrados()
'
' Macro1 Macro
'

'
    Dim column As Integer
    Dim index_name As String
    Dim nombre_hoja As String
    
nombre_hoja = "Salidas vs mantenimientos"
criterio_filtrado = "<>No Dato"
index_name = Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5) ' nombre de la hoja que se desea filtrar
column = 12 'Nº de columna de la hoja a filtrar
    
    'Eliminar si ya existe una hoja con el mismo nombre
Application.DisplayAlerts = False
For i = 1 To Worksheets.Count
On Error Resume Next
    If Sheets(i).Name = nombre_hoja Then
        Sheets(i).Delete
        i = i + 1
    End If
Next
Application.DisplayAlerts = True
    
    'Añadir nueva hoja con los datos filtrados
Sheets.Add.Name = nombre_hoja
Sheets("" & index_name & "").Range("B" & column).AutoFilter Field:=column, Criteria1:=criterio_filtrado, Criteria2:="<>Zona Geo Inadecuada"
Sheets("" & index_name & "").Range("B" & column).AutoFilter Field:=19, Criteria1:="<>Entra"
Sheets("" & index_name & "").Range("B" & column).AutoFilter Field:=8, Criteria1:="> 200705"
Sheets("" & index_name & "").AutoFilter.Range.Copy
ActiveSheet.Paste Destination:=Range("A1") 'celda en la que se empiezan a pegar los datos
Cells.Interior.ColorIndex = 2

Sheets("" & index_name & "").ShowAllData ' no es necesario pero por si lo necesito alguna vez
Sheets("" & index_name & "").Range("B" & column).AutoFilter 'con esto basta para quitar el filtro
Sheets("" & nombre_hoja & "").Select


End Sub

    
Sub Rango_percentil_macro()
'
' Macro1 Macro
'

'
    'Calcular el rango percentil de un conjunto de datos delimitados por un criterio. Este criterio debe estar ordenado
    
    Dim contador As Long
    Dim Columna_inicio As Integer 'columna inicial para extraer datos de cálculo rango percentil
    Dim columna_final As Integer 'Columa final para extraer datos rango percentil
    Dim colocar_datos As Integer 'desde que columna se empiezan a colar los resultados de RP
    Dim columna_criterio As Integer
    Dim matriz_rango
    Dim region As Long
    Dim datos_matriz As Long
    Dim rango_percintil As Double
    Dim columna_criterio2 As Integer
    
Columna_inicio = 12
columna_final = 15
colocar_datos = 19
columna_criterio = 8
'columna_criterio2 = 16
region = Range("b2").CurrentRegion.Rows.Count

For ii = 2 To region
    datos_matriz = Application.WorksheetFunction.CountIfs(Range(Cells(ii, columna_criterio), Cells(region, columna_criterio)), _
    Cells(ii, columna_criterio))
    contador = contador + 1
    Range("A" & ii).Value = contador
    ReDim matriz_rango(datos_matriz - 1, 0)
        For datos_columnas = 0 To columna_final - Columna_inicio
            For datos_filas = 0 To datos_matriz - 1
                matriz_rango(datos_filas, 0) = Cells(datos_filas + ii, Columna_inicio + datos_columnas)
            Next
            For rango_dato = 0 To datos_matriz - 1
                Rango_percentil = Application.WorksheetFunction.PercentRank_Inc(matriz_rango, matriz_rango(rango_dato, 0), 5)
                Cells(ii + rango_dato, datos_columnas + colocar_datos).Value = Rango_percentil
            Next
        Next
    ii = ii + datos_matriz - 1
Next
    'colorines, centrar datos, (pijadas)
For datos_columnas2 = 0 To columna_final - Columna_inicio
    Columns(Columna_inicio + datos_columnas2).Interior.ColorIndex = 37
    Cells(1, colocar_datos + datos_columnas2).Interior.ColorIndex = 43
    Cells(1, colocar_datos + datos_columnas2).Value = "RP " & Cells(1, Columna_inicio + datos_columnas2)

Next
Rows(1).Select
    With Selection
        .Font.Bold = True
    End With
Cells(1, colocar_datos + datos_columnas2).Select

Call dejar_bonito
 
    'opcional
'Cells.EntireColumn.AutoFit

Call T_Test_Salida_vs_mantenimientos
  End Sub


Sub entradas_universo()

Dim hoja_name  As String
Dim libro_name As String
Dim clave As Long 'mi ref
Dim comienzo As Integer
Dim nombre_hoja As String
Dim universo As Long
Dim referencia_cruzada As Integer
Dim comienzo_índices_no_stbl As Integer

comienzo_índices_no_stbl = 7 ' Clafificación índices, fila donde empiezan los índices no sostenibles.
referencia_cruzada = 16
nombre_hoja = "Univers vs entra" ' nombre nueva hoja
list_column = 26 ' columna donde se empiezan a poner datos en list scores

Application.DisplayAlerts = False
On Error Resume Next
Sheets("" & nombre_hoja & "").Delete
Application.DisplayAlerts = True

Sheets.Add.Name = nombre_hoja
Sheets("" & nombre_hoja & "").Select
Set Rng = Range("A1")
Cells.Interior.ColorIndex = 2
Set wbk = ActiveWorkbook

libro_name = Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5)
hoja_name = Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5)
Sheets("" & libro_name & "").Select
universo = Range("B1").CurrentRegion.Rows.Count

ThisWorkbook.Activate
Columns(list_column + copia * 1).ClearContents

For aa = 2 To universo
'Referencia número de fila o no controlado
If Workbooks("" & libro_name & "").Sheets("" & hoja_name & "").Cells(aa, referencia_cruzada) = "No controlado" Then
Else
clave = Workbooks("" & libro_name & "").Sheets("" & hoja_name & "").Cells(aa, referencia_cruzada) 'REFERENCIA CRUZADA
Cells(clave, list_column + copia * 1) = "Dentro"
    If Workbooks("" & libro_name & "").Sheets("" & hoja_name & "").Cells(aa, 19) = "Entra" Then 'celda donde estan las entrada
        Cells(clave, list_column + copia * 1) = "1º mes dentro"
        End If
    If Workbooks("" & libro_name & "").Sheets("" & hoja_name & "").Cells(aa, 18) = "Sale" _
    And Workbooks("" & libro_name & "").Sheets("" & hoja_name & "").Cells(aa, 5) = Cells(clave + 1, 4) Then 'celda donde están las salidas y ISIN para comprobar
        Cells(clave + 1, list_column + copia * 1) = "1º mes fuera"
        End If
End If
Next
Cells(1, list_column + copia * 1).Value = hoja_name


Dim column As Integer
Dim index_name As String
Dim criterio_filtrado As String

Set rng2 = Worksheets("Clasificación índices").Range("F1:F" & 50).Find("" & hoja_name & "") 'zona geográfica índice
criterio_filtrado = rng2.Offset(0, 1)
Worksheets("Hipótesis 3").Cells(2, copia + 2).Value = rng2.Offset(0, -2) 'Ponemos el nombre del índice en la hipótesis correspondiente
Worksheets("Hipótesis 4").Cells(2, copia + 2).Value = rng2.Offset(0, -2) 'Ponemos el nombre del índice en la hipótesis correspondiente
Worksheets("Hipótesis 3").Cells(1, copia + 2).Value = criterio_filtrado
Worksheets("Hipótesis 4").Cells(1, copia + 2).Value = criterio_filtrado
If rng2.Offset(0, -3) = "S" Then
Colorcito = RGB(146, 208, 80) 'verde
Else
Colorcito = RGB(201, 201, 201) 'gris
End If
Worksheets("Hipótesis 3").Cells(2, copia + 2).Interior.Color = Colorcito 'Ponemos el nombre del índice en la hipótesis correspondiente el color correspondiente
Worksheets("Hipótesis 4").Cells(2, copia + 2).Interior.Color = Colorcito
If criterio_filtrado = "Global" Then
criterio_filtrado = "<>"
End If
 
index_name = "Listado Scores y MV" ' nombre de la hoja que se desea filtrar
column = list_column - 1 'Nº de columna de la hoja a filtrar
    


    'Añadir nueva hoja con los datos filtrados

Sheets("" & index_name & "").Range("B" & column).AutoFilter Field:=column, Criteria1:=criterio_filtrado
Sheets("" & index_name & "").Range("B" & column).AutoFilter Field:=list_column + 1 * copia, Criteria1:="1º mes dentro", _
Operator:=xlOr, Criteria2:="="
Sheets("" & index_name & "").Range("B" & column).AutoFilter Field:=2, Criteria1:=">05/31/2007"
Sheets("" & index_name & "").AutoFilter.Range.Copy

ActiveSheet.Paste Destination:=Rng 'celda en la que se empiezan a pegar los datos
Sheets("" & index_name & "").ShowAllData ' no es necesario pero por si lo necesito alguna vez
Sheets("" & index_name & "").Range("B" & column).AutoFilter 'con esto basta para quitar el filtro
wbk.Activate
Sheets("" & nombre_hoja & "").Select
For xx = 0 To copia - 1
Columns(list_column).Delete
Next
Columns("V:W").Delete
Columns("J:Q").Delete
Columns(9).Delete
Cells(1, 15).Value = "Universo vs Entrada"

Call ordenar_univers
Call dejar_bonito
 
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

    
Sub T_Test_Salida_vs_mantenimientos()

Dim columna_m1 As Integer
Dim columna_m2 As Integer
Dim m1 As Variant
Dim m2 As Variant
Dim N_salen As Long
Dim N_permanecen As Long
Dim Columna_inicio As Integer
Dim Columna_fin As Integer
Dim contador_m1 As Long
Dim contador_m2 As Long
Dim letras As Variant
letras = Array("S", "T", "U", "V", "W", "X", "Y", "Z", "AA") 'LETRAS COMO COLUMNAS ANALICEMOS


columna_m1 = 18 'matriz uno sale
columna_m2 = 17 'matriz dos se mantiene

Columna_inicio = 19 'donde empezar a coger datos
Columna_fin = 22 'donde terminar de cogerlos
columna_receptora = 24 'donde colocarlos
columna_fecha = 8

ThisWorkbook.Sheets("Hipótesis 3").Columns(2 + copia).ClearContents
'Una única restricción en cada caso
hasta = WorksheetFunction.CountIfs(Range(Cells(2, columna_fecha), Cells(Cells(2, columna_fecha).CurrentRegion.Rows.Count, columna_fecha)), _
 "< 201707") + 1

N_salen = WorksheetFunction.CountIfs(Range(Cells(2, columna_m1), Cells(hasta, columna_m1)), _
"Sale")
N_permanecen = WorksheetFunction.CountIfs(Range(Cells(2, columna_m2), Cells(hasta, columna_m2)), _
"Se mantiene")


ReDim m2(N_permanecen)
ReDim m1(N_salen)

For KTK = Columna_inicio To Columna_fin
contador_m1 = 0
contador_m2 = 0
For ase = 2 To hasta
If Cells(ase, columna_m2) = "Se mantiene" Then
m2(contador_m2) = Cells(ase, KTK)
contador_m2 = contador_m2 + 1
Else
If Cells(ase, columna_m1) = "Sale" Then
m1(contador_m1) = Cells(ase, KTK)
contador_m1 = contador_m1 + 1
End If
End If
Next

Media = ActiveSheet.Evaluate("=AVERAGE(IF((Q2:Q" & hasta & "=""Se mantiene"")," & letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "))")
Varianza = ActiveSheet.Evaluate("=VAR.P(IF((Q2:Q" & hasta & "=""Se mantiene"")," & letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "))")
PruebaF = ActiveSheet.Evaluate("=FTEST(IF((Q2:Q" & hasta & "=""Se mantiene"")," & letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "),IF((R2:R" & hasta & "=""Sale"")," & _
letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "))")
    If PruebaF < 0.05 Then
    a = 3
    Else
    a = 2
    End If
PruebaT = ActiveSheet.Evaluate("=TTEST(IF((Q2:Q" & hasta & "=""Se mantiene"")," & letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "),IF((R2:R" & hasta & "=""Sale"")," & _
letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "),2," & a & ")")




Cells(2, columna_receptora + KTK - Columna_inicio + 1).Value = N_permanecen
Cells(3, columna_receptora + KTK - Columna_inicio + 1).Value = N_salen

ThisWorkbook.Sheets("Hipótesis 3").Cells(5 + 10 * (KTK - Columna_inicio), 2 + copia).Value = N_permanecen
ThisWorkbook.Sheets("Hipótesis 3").Cells(6 + 10 * (KTK - Columna_inicio), 2 + copia).Value = N_salen

Cells(4, columna_receptora + KTK - Columna_inicio + 1).Value = Varianza
Cells(5, columna_receptora + KTK - Columna_inicio + 1).Value = WorksheetFunction.Var_P(m1)
Cells(6, columna_receptora + KTK - Columna_inicio + 1).Value = PruebaF

ThisWorkbook.Sheets("Hipótesis 3").Cells(7 + 10 * (KTK - Columna_inicio), 2 + copia).Value = Varianza
ThisWorkbook.Sheets("Hipótesis 3").Cells(8 + 10 * (KTK - Columna_inicio), 2 + copia).Value = WorksheetFunction.Var_P(m1)
ThisWorkbook.Sheets("Hipótesis 3").Cells(9 + 10 * (KTK - Columna_inicio), 2 + copia).Value = PruebaF

Cells(7, columna_receptora + KTK - Columna_inicio + 1).Value = Media
Cells(8, columna_receptora + KTK - Columna_inicio + 1).Value = WorksheetFunction.Average(m1)
Cells(9, columna_receptora + KTK - Columna_inicio + 1).Value = PruebaT

ThisWorkbook.Sheets("Hipótesis 3").Cells(10 + 10 * (KTK - Columna_inicio), 2 + copia).Value = Media
ThisWorkbook.Sheets("Hipótesis 3").Cells(11 + 10 * (KTK - Columna_inicio), 2 + copia).Value = WorksheetFunction.Average(m1)
ThisWorkbook.Sheets("Hipótesis 3").Cells(12 + 10 * (KTK - Columna_inicio), 2 + copia).Value = PruebaT

Cells(1, columna_receptora + KTK - Columna_inicio + 1).Value = Mid(Cells(1, KTK), 4)

ThisWorkbook.Sheets("Hipótesis 3").Cells(4 + 10 * (KTK - Columna_inicio), 2 + copia).Value = _
Mid(Cells(1, KTK), 4)

Next
ThisWorkbook.Sheets("Hipótesis 3").Cells(3, 2 + copia).Value = _
Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5)
'ThisWorkbook.Sheets("Hipótesis 3").Cells(1, 2 + copia).Value = _
'Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5)


'Dar nombres a la tabla
Cells(2, columna_receptora).Value = "Maintenances observations"
Cells(3, columna_receptora).Value = "Deletions observations"
Cells(4, columna_receptora).Value = "Maintenances Variance"
Cells(5, columna_receptora).Value = "Deletions Variance"
Cells(6, columna_receptora).Value = "P_value F-Test"
Cells(7, columna_receptora).Value = "Maintenances Average"
Cells(8, columna_receptora).Value = "Deletions Average"
Cells(9, columna_receptora).Value = "P_value T-Test"

End Sub



Sub Rango_percentil_macro_universo_entradas()
'
' Macro1 Macro
'

'
    'Calcular el rango percentil de un conjunto de datos delimitados por un criterio. Este criterio debe estar ordenado


    Dim contador As Long
    Dim Columna_inicio As Integer 'columna inicial para extraer datos de cálculo rango percentil
    Dim columna_final As Integer 'Columa final para extraer datos rango percentil
    Dim colocar_datos As Integer 'desde que columna se empiezan a colar los resultados de RP
    Dim columna_criterio As Integer
    Dim matriz_rango
    Dim region As Long
    Dim datos_matriz As Long
    Dim rango_percintil As Double

    
Columna_inicio = 10
columna_final = 13
colocar_datos = 17
columna_criterio = 2
region = Range("b2").CurrentRegion.Rows.Count

For ii = 2 To region
    datos_matriz = Application.WorksheetFunction.CountIfs(Range(Cells(ii, columna_criterio), Cells(region, columna_criterio)), _
    Cells(ii, columna_criterio))
    contador = contador + 1
    Range("A" & ii).Value = contador
    ReDim matriz_rango(datos_matriz - 1, 0)
        For datos_columnas = 0 To columna_final - Columna_inicio
            For datos_filas = 0 To datos_matriz - 1
                matriz_rango(datos_filas, 0) = Cells(datos_filas + ii, Columna_inicio + datos_columnas)
            Next
            For rango_dato = 0 To datos_matriz - 1
                Rango_percentil = Application.WorksheetFunction.PercentRank_Inc(matriz_rango, matriz_rango(rango_dato, 0), 5)
                Cells(ii + rango_dato, datos_columnas + colocar_datos).Value = Rango_percentil
            Next
        Next
    ii = ii + datos_matriz - 1
Next
    'colorines, centrar datos, (pijadas)
For datos_columnas2 = 0 To columna_final - Columna_inicio
    Columns(Columna_inicio + datos_columnas2).Interior.ColorIndex = 37
    Cells(1, colocar_datos + datos_columnas2).Interior.ColorIndex = 43
    Cells(1, colocar_datos + datos_columnas2).Value = "RP " & Cells(1, Columna_inicio + datos_columnas2)
Next
Rows(1).Select
    With Selection
        .Font.Bold = True
    End With
Cells(1, colocar_datos + datos_columnas2).Select

Call T_Test_Entradas_vs_Universo

End Sub


Sub ordenar_univers()
Dim Columna_ordenado As Integer
Dim letra As String

Columns(1).Insert
num_filas = Range("B1").CurrentRegion.Rows.Count
For com = 2 To num_filas
Cells(com, 2).Value = Mid(Cells(com, 3), 7, 4) & Mid(Cells(com, 3), 4, 2)
Next
Cells(1, 2).Value = "Fecha Numérica"
Columna_ordenado = 2
letra = Split(Cells(1, Columna_ordenado).Address, "$")(1)

ActiveWorkbook.Worksheets("Univers vs entra").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Univers vs entra").Sort.SortFields.Add2 Key:=Range _
        ("" & letra & "2:" & letra & Range("C1").CurrentRegion.Rows.Count), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Univers vs entra").Sort
        .SetRange Range("A1:ZL" & Range("C1").CurrentRegion.Rows.Count)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
cmb_mes = 1

Cells(1, 1).Value = "Cmb mes"
For com = 3 To num_filas
    If Cells(com, 2) = Cells(com - 1, 2) Then
    Else
    Cells(com, 1) = cmb_mes
    cmb_mes = cmb_mes + 1
    End If
Next

End Sub

Sub T_Test_Entradas_vs_Universo()

'borrar desde columna 2 hoja hipótesis4

Dim columna_m1 As Integer
Dim columna_m2 As Integer
Dim m1 As Variant
Dim m2 As Variant
Dim N_entran As Long
Dim N_universo As Long
Dim Columna_inicio As Integer
Dim Columna_fin As Integer
Dim contador_m1 As Long
Dim contador_m2 As Long
Dim letras As Variant
letras = Array("Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z") 'LETRAS COMO COLUMNAS ANALICEMOS


columna_m1 = 16 'matriz uno
columna_m2 = 16 'matriz dos

Columna_inicio = 17 'donde empezar a coger datos
Columna_fin = 20    'donde terminar de cogerlos
columna_receptora = 22 'donde colocarlos
columna_fecha = 2

'Una única restricción en cada caso
hasta = WorksheetFunction.CountIfs(Range(Cells(2, columna_fecha), Cells(Cells(2, columna_fecha).CurrentRegion.Rows.Count, columna_fecha)), _
 "< 201707") + 1 'fecha límite

N_entran = WorksheetFunction.CountIfs(Range(Cells(2, columna_m1), Cells(hasta, columna_m1)), _
"1º mes dentro") '
N_universo = WorksheetFunction.CountIfs(Range(Cells(2, columna_m2), Cells(hasta, columna_m2)), _
"") 'universo de posibilidades


ReDim m2(N_universo)
ReDim m1(N_entran)

For KTK = Columna_inicio To Columna_fin
contador_m1 = 0
contador_m2 = 0
For ase = 2 To hasta
If Cells(ase, columna_m2) = "" Then
m2(contador_m2) = Cells(ase, KTK)
contador_m2 = contador_m2 + 1
Else
If Cells(ase, columna_m1) = "1º mes dentro" Then
m1(contador_m1) = Cells(ase, KTK)
contador_m1 = contador_m1 + 1
End If
End If
Next
criterio = "1º mes dentro"
Media = ActiveSheet.Evaluate("=AVERAGE(IF((P2:P" & hasta & "="""")," & letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "))")
Varianza = ActiveSheet.Evaluate("=VAR.P(IF((P2:P" & hasta & "="""")," & letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "))")
PruebaF = ActiveSheet.Evaluate("=FTEST(IF((P2:P" & hasta & "="""")," & letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "),IF((P2:P" & hasta & "=""1º mes dentro"")," & _
letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "))")
    If PruebaF < 0.05 Then
    a = 3
    Else
    a = 2
    End If
PruebaT = ActiveSheet.Evaluate("=TTEST(IF((P2:P" & hasta & "="""")," & letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "),IF((P2:P" & hasta & "=""1º mes dentro"")," & _
letras(KTK - Columna_inicio) & "2:" & letras(KTK - Columna_inicio) & "" & hasta & "),2," & a & ")")



Cells(2, columna_receptora + KTK - Columna_inicio + 1).Value = N_universo
Cells(3, columna_receptora + KTK - Columna_inicio + 1).Value = N_entran

ThisWorkbook.Sheets("Hipótesis 4").Cells(5 + 10 * (KTK - Columna_inicio), 2 + copia).Value = N_universo
ThisWorkbook.Sheets("Hipótesis 4").Cells(6 + 10 * (KTK - Columna_inicio), 2 + copia).Value = N_entran

Cells(4, columna_receptora + KTK - Columna_inicio + 1).Value = Varianza
Cells(5, columna_receptora + KTK - Columna_inicio + 1).Value = WorksheetFunction.Var_P(m1)
Cells(6, columna_receptora + KTK - Columna_inicio + 1).Value = PruebaF

ThisWorkbook.Sheets("Hipótesis 4").Cells(7 + 10 * (KTK - Columna_inicio), 2 + copia).Value = Varianza
ThisWorkbook.Sheets("Hipótesis 4").Cells(8 + 10 * (KTK - Columna_inicio), 2 + copia).Value = WorksheetFunction.Var_P(m1)
ThisWorkbook.Sheets("Hipótesis 4").Cells(9 + 10 * (KTK - Columna_inicio), 2 + copia).Value = PruebaF

Cells(7, columna_receptora + KTK - Columna_inicio + 1).Value = Media
Cells(8, columna_receptora + KTK - Columna_inicio + 1).Value = WorksheetFunction.Average(m1)
Cells(9, columna_receptora + KTK - Columna_inicio + 1).Value = PruebaT

ThisWorkbook.Sheets("Hipótesis 4").Cells(10 + 10 * (KTK - Columna_inicio), 2 + copia).Value = Media
ThisWorkbook.Sheets("Hipótesis 4").Cells(11 + 10 * (KTK - Columna_inicio), 2 + copia).Value = WorksheetFunction.Average(m1)
ThisWorkbook.Sheets("Hipótesis 4").Cells(12 + 10 * (KTK - Columna_inicio), 2 + copia).Value = PruebaT


Cells(1, columna_receptora + KTK - Columna_inicio + 1).Value = Mid(Cells(1, KTK), 4)

ThisWorkbook.Sheets("Hipótesis 4").Cells(4 + 10 * (KTK - Columna_inicio), 2 + copia).Value = _
Mid(Cells(1, KTK), 4)


Next
ThisWorkbook.Sheets("Hipótesis 4").Cells(3, 2 + copia).Value = _
Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5)
'ThisWorkbook.Sheets("Hipótesis 4").Cells(1, 2 + copia).Value = _
'Mid(ActiveWorkbook.Name, 1, Len(ActiveWorkbook.Name) - 5)
'Dar nombres a la tabla
Cells(2, columna_receptora).Value = "Univers observations"
Cells(3, columna_receptora).Value = "Inclusions observations"
Cells(4, columna_receptora).Value = "Univers Variance"
Cells(5, columna_receptora).Value = "Inclusions Variance"
Cells(6, columna_receptora).Value = "P_value F-Test"
Cells(7, columna_receptora).Value = "Univers Average"
Cells(8, columna_receptora).Value = "Inclusions Average"
Cells(9, columna_receptora).Value = "P_value T-Test"

End Sub





