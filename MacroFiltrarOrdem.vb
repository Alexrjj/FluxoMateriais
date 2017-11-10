Sub Macro1()
'
' Macro1 Macro
'

'
    'Cria nova aba
    Sheets.Add After:=ActiveSheet
    'Digite o cabeçalho das colunas necessárias
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Código"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Material"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Qtd"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "CódigoSAP"
    'Seleciona e copia os códigos da ficha de materiais
    Sheets("Ficha_Solicitação_Material").Select
    Range("H9").Select
    ActiveCell.Offset(1, 0).Select 'Seleciona uma célula abaixo
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    'Cola o código de material na nova aba
    Sheets("Planilha2").Select
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Seleciona a aba ficha de material, copia a descrição dos materiais e cola na nova aba
    Sheets("Ficha_Solicitação_Material").Select
    Dim codigomat As Long
    codigomat = Range("H500").End(xlUp).Row
    Range("I10:I" & codigomat).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Planilha2").Select
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      :=False, Transpose:=False
    'Seleciona a aba ficha de material, copia a qtd dos materiais e cola na nova aba
    Sheets("Ficha_Solicitação_Material").Select
    Range("L9").Select
    ActiveCell.Offset(1, 0).Select 'Seleciona uma célula abaixo
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Planilha2").Select
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Seleciona a aba ficha de material, copia o código sap dos materiais e cola na nova aba
    Sheets("Ficha_Solicitação_Material").Select
    Dim codigosap2 As Long
    codigosap2 = Range("L500").End(xlUp).Row
    Range("M10:M" & codigosap2).Select
    Selection.Copy
    Sheets("Planilha2").Select
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    'Seleciona todas as células e ajusta a largura
    Cells.Select
    Selection.Columns.AutoFit
    'Seleciona a célula A1 e aplica o filtro
    Range("A1").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Planilha2").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Planilha2").AutoFilter.Sort.SortFields.Add Key:= _
        Range("B1:B500"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Planilha2").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Seleciona os códigos de materiais e cola na ficha de material
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Ficha_Solicitação_Material").Select
    Range("H10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Seleciona a qtd de materiais na nova aba e cola na ficha de material
    Sheets("Planilha2").Select
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Ficha_Solicitação_Material").Select
    Range("L10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Seleciona o código sap na nova aba e cola na ficha de material
    Sheets("Planilha2").Select
    Dim codigosap3 As Long
    codigosap3 = Range("C500").End(xlUp).Row
    Range("D2:D" & codigosap3).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Ficha_Solicitação_Material").Select
    Range("M10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("H9").Select
End Sub
