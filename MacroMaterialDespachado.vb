'Monta automaticamente o fluxo de materiais no padrão MRO com base no pedido ZPIN do GOMNET.
'Obs.: É feita verificação automática de material restante disponível em cada código SAP.

Sub materialDesp()
    Sheets.Add After:=ActiveSheet 'Acrescenta uma nova aba para começar a trabalhar com os dados obtidos.
    ActiveSheet.Paste
    Selection.ClearFormats
    Selection.Columns.AutoFit
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "s"
    Selection.AutoFilter
    Columns("C:C").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("C1").Select
    ActiveWorkbook.Worksheets("Planilha1").AutoFilter.Sort.SortFields.Clear 'Filtra a coluna "data" de forma descendente para remover materiais duplicados
    ActiveWorkbook.Worksheets("Planilha1").AutoFilter.Sort.SortFields.Add Key:= _
        Range("C1:C500"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption _
        :=xlSortNormal
    With ActiveWorkbook.Worksheets("Planilha1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ActiveSheet.Range("$A$1:$J$500").RemoveDuplicates Columns:=4, Header:=xlYes 'Remove os códigos de materiais duplicados
    
    Range("I1").Select
    ActiveSheet.Range("$A$1:$J$500").AutoFilter Field:=9, Criteria1:="=" 'Filtra apenas campos vazios
    
    'Copia e cola os códigos dos materiais
    Range("D1").Select
    ActiveCell.Offset(1, 0).Select 'Seleciona uma célula abaixo
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Ficha_Solicitação_Material").Select
    Range("H10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Copia e cola a quantidade de materiais
    Sheets("Planilha1").Select
    Range("H1").Select
    ActiveCell.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Ficha_Solicitação_Material").Select
    Range("L10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Copia e cola os pedidos (SAP) dos materiais
    Sheets("Planilha1").Select
    Range("B1").Select
    ActiveCell.Offset(1, 0).Select
    
    'Verifica, com base na coluna "A", a quantidade de linhas necessárias para selecionar as linhas de código SAP, pois podem haver celulas vazias
    '----------------------'
    Dim codigosap As Long
    codigosap = Range("A500").End(xlUp).Row
    Range("B2:B" & codigosap).Select
    '----------------------'
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Ficha_Solicitação_Material").Select
    Range("M10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Filtra os materiais já despachados, seja pedido parcial ou total
    '----------------------'
    Sheets("Planilha1").Select
    Range("I1").Select
    ActiveSheet.Range("$A$1:$J$500").AutoFilter Field:=9, Criteria1:="<>" 'Filtra apenas campos não vazios
    '----------------------'
    Range("J1").Select
    ActiveCell.Offset(1, 0).Select
    Application.CutCopyMode = False
    
    'Verifica, com base na coluna H, a quantidade de linhas necessárias para expandir a fórmula. Do contrário, o comando expandiria para até o fim da planilha, travando o Excel.
    'Aplica a fórmula de subtração para verificar se há algum material restante em cada pedido SAP.
    '----------------------'
    Dim lastrow As Long
    lastrow = Range("H500").End(xlUp).Row
    ActiveCell.FormulaR1C1 = "=IMSUB(RC[-2],RC[-1])"
    Selection.AutoFill Destination:=Range("J2:J" & lastrow)
    '----------------------'
    Range("J1").Select
    ActiveCell.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    'Define a variavel "selecao", com base nas células já selecionadas, para a execução dos próximos comandos.
    
    'Verifica, com base na coluna I, a quantidade necessária de linhas a serem selecionadas. Até aqui, sem alteração, por já ter sido executado no comando anterior.
    '----------------------'
    Dim selecao As Long
    selecao = Range("I500").End(xlUp).Row
    Range("J2:J" & selecao).Select
    '----------------------'
    
    'Converte, com base no resultado da fórmula de subtração, todas as células selecionadas e visíveis, para o formato "número". Do contrário, o Excel os reconhece como texto, gerando erro nos próximos comandos.
    'É feito um loop dentro de outro loop para verificar cada célula selecionada para convertê-la em numeral. A quantidade de linhas é baseada na variável "selecao" definida anteriormente.
    '----------------------'
    Dim converte As Range, bloco As Range
    Set bloco = Range("J2:J" & selecao)
    For Each converte In bloco.SpecialCells(xlCellTypeVisible)
    For Each r In converte
    If IsNumeric(r) Then
       r.Value = CSng(r.Value)
       r.NumberFormat = "0.0"
    End If
    Next
    Next converte
    '----------------------'
  
    'Define a variável booleana "found", instanciada como 0 (False). É feita a verificação linha a linha (loop), com base na variável "selecao", para saber se há algum número maior que 0.
    'Caso retorne 1 (True), significa que há códigos SAP com materiais ainda disponíveis a serem despachados.
    '----------------------'
    Dim found As Boolean
    found = False
    Dim cl As Range, rng As Range
    Set rng = Range("J2:J" & selecao)
    For Each cl In rng.SpecialCells(xlCellTypeVisible)
        If cl > 0 Then
            found = True
    End If
    Next cl
    '----------------------'

    'Caso "found" retorne 1 (True), será copiado todo código SAP com seus respectivos materiais ainda disponíveis para retirada.
    'Caso "found" retorne 0 (False), serão copiados apenas os materiais e suas quantidades (totais) já despachadas, sem o código SAP para retirada dos mesmos.
    If found = True Then
        ActiveSheet.Range("$A$1:$B$500").AutoFilter Field:=10, Criteria1:="<>0" 'Filtra apenas campos diferentes de 0
        Range("D1").Select
        ActiveCell.Offset(1, 0).Select
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        Sheets("Ficha_Solicitação_Material").Select
        Range("H10").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Sheets("Planilha1").Select
        Range("J1").Select
        ActiveCell.Offset(1, 0).Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Ficha_Solicitação_Material").Select
        Range("L10").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Sheets("Planilha1").Select
        Range("B1").Select
        ActiveCell.Offset(1, 0).Select
       
        Dim codigosap2 As Long
        codigosap2 = Range("A500").End(xlUp).Row
        Range("B2:B" & codigosap2).Select
    
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Ficha_Solicitação_Material").Select
        ActiveCell.Offset(0, 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Sheets("Planilha1").Select
        ActiveSheet.Range("$A$1:$J$500").AutoFilter Field:=10, Criteria1:="0,0"
        Range("D1").Select
        ActiveCell.Offset(1, 0).Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Ficha_Solicitação_Material").Select
        Range("H10").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Sheets("Planilha1").Select
        Range("H1").Select
        ActiveCell.Offset(1, 0).Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Ficha_Solicitação_Material").Select
        Range("L10").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("H10").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(0, -1).Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.CutCopyMode = False
        Selection.EntireRow.Delete
    Else
        Range("D1").Select
        ActiveCell.Offset(1, 0).Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.CutCopyMode = False
        Selection.Copy
        Sheets("Ficha_Solicitação_Material").Select
        Range("H10").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        Sheets("Planilha1").Select
        Range("I1").Select
        ActiveCell.Offset(1, 0).Select
    
        Dim material As Long
        material = Range("H500").End(xlUp).Row
        Range("I2:I" & material).Select
        
        Selection.Copy
        Sheets("Ficha_Solicitação_Material").Select
        Range("L10").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
        'Exclui as linhas em branco restantes na ficha de solicitação de material
        Range("H10").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        ActiveCell.Offset(0, -1).Select
        Range(Selection, Selection.End(xlDown)).Select
        Application.CutCopyMode = False
        Selection.EntireRow.Delete
    End If
End Sub