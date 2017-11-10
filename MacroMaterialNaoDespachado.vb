'Monta automaticamente o fluxo de materiais no padrão MRO com base no pedido ZPIN do GOMNET.
'Obs.: NÃO é feita verificação automática de material restante disponível em cada código SAP. Serve apenas para obras novas.

Sub materialNaoDesp()
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Selection.ClearFormats
    Selection.Columns.AutoFit
    
    Range("D1").Select
    ActiveCell.Offset(1, 0).Select 'Seleciona uma célula abaixo
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Ficha_Solicitação_Material").Select
    Range("H10").Select
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
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Planilha1").Select
    Range("B1").Select
    ActiveCell.Offset(1, 0).Select

    Dim codigosap As Long
    codigosap = Range("A500").End(xlUp).Row
    Range("B2:B" & codigosap).Select

    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Ficha_Solicitação_Material").Select
    Range("M10").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'Excluindo linhas em branco restantes
    Range("H10").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Offset(0, -1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Delete
End Sub
