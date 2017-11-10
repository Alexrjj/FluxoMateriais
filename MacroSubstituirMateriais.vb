'Esta macro substitui materiais fora padrão através de códigos pré definidos
'Basta preencher o código atual no array "fndList" e seu código substituto no array "rplcList", na ordem correta.
Sub Multi_FindReplace()

Dim fndList As Variant
Dim rplcList As Variant
Dim x As Long

fndList = Array("6810650", "6773388", "6792845", "6772224", "4589331", "6782623", "6782554")
rplcList = Array("6774159", "6773389", "6771968", "6772225", "6808985", "4679894", "6772186")

'1. Faz um loop em todos os itens no vetor fndList;
'2. Procura cada item e apaga seu respectivo código SAP.
For y = LBound(fndList) To UBound(fndList)
    Cells.Find(What:=fndList(y)).Select
    Selection.Offset(0, 5).Select
    Selection.ClearContents
Next y

'1. Faz um loop em todos os itens no vetor fndList e substitui de acordo com cada item no vetor rplcList.
For x = LBound(fndList) To UBound(fndList)
    Cells.Replace What:=fndList(x), Replacement:=rplcList(x), _
    LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
    SearchFormat:=False, ReplaceFormat:=False
Next x

End Sub
