'Procura células vazias com base na coluna "L" e deleta a linha inteira, facilitando a correção dos materiais desnecessários.

Sub apagaLinhas()
Dim ws As Excel.Worksheet
Dim LastRow As Long

Set ws = ActiveSheet
LastRow = ws.Range("K" & ws.Rows.Count).End(xlUp).Row
With ws.Range("L10:L" & LastRow)
    If WorksheetFunction.CountBlank(.Cells) > 0 Then
        .SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    End If
End With
End Sub
