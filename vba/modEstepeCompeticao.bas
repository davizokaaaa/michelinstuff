Attribute VB_Name = "modEstepeCompeticao"
Option Explicit

' ==========================================================================
' MODULO: modEstepeCompeticao
' Classifica ESTEPE, CAMPEONATO/COMPETICAO e alguns casos especificos de
' LCV a partir da DESCRICAODA MERCADORIA -- grava numa coluna propria, sem
' se misturar com LP/TIPO PRODUTO/MARCA/GAMA/DIMENSAO.
'
' ORIGEM: portado 1:1 da macro legada "Sub estepe()", que rodava sozinha,
' lendo a coluna H (hardcoded) e escrevendo na coluna T (hardcoded) da
' planilha ativa. Aqui a logica de decisao e IDENTICA (mesmas palavras-
' chave, mesma ordem de prioridade em cascata If/ElseIf, mesmo
' vbTextCompare) -- so isolada numa funcao pura que recebe o texto ja
' extraido pelo orquestrador (modMain), pra rodar dentro do mesmo laco
' principal em vez de exigir uma segunda passada pela planilha inteira.
'
' IMPORTANTE (equivalencia de resultado): a macro legada fazia
' UCase(ws.Cells(i, "H").Value) sem Trim. O orquestrador (modMain) ja
' calcula "descricao" como UCase(Trim(...)) da coluna DESCRICAODA
' MERCADORIA antes de chamar qualquer extracao -- Trim() so remove espacos
' nas PONTAS do texto, o que nunca muda o resultado de um InStr procurando
' uma palavra/trecho no MEIO do texto. Portanto passar "descricao" (em vez
' de reler a celula) preserva exatamente o mesmo resultado da macro legada,
' assumindo que a coluna H da planilha legada e a DESCRICAODA MERCADORIA
' sao a mesma coluna -- CONFIRME isso ao validar lado a lado (ver PREMISSA
' abaixo).
'
' PREMISSA A CONFIRMAR: a macro legada lia da coluna H. Se, na planilha
' onde a macro legada rodava, a coluna H NAO for a mesma coisa que
' "DESCRICAODA MERCADORIA", ajuste a chamada em modMain para passar o
' texto certo.
' ==========================================================================

' ==========================================================================
' Classifica uma linha em ESTEPE / CAMPEONATO / COMPETICAO / LCV (ou ""),
' na mesma ordem de prioridade em cascata da macro legada -- a primeira
' regra que bater decide, as de baixo nem sao checadas. Espera o texto ja
' em maiusculas (mesmo formato que "descricao" chega em modMain); vbTextCompare
' no InStr ja ignora caixa mesmo assim, entao funciona igual com texto misto.
' ==========================================================================
Function ClassificarEstepeCompeticao(texto As String) As String

    If InStr(1, texto, "ESTEPE", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "ESTEPE"
    ElseIf InStr(1, texto, "CAMPEONATO", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "CAMPEONATO"
    ElseIf InStr(1, texto, "COMPETICAO", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "COMPETICAO"
    ElseIf InStr(1, texto, "225/65 R16 C", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "LCV"
    ElseIf InStr(1, texto, "BANDA: 225, SERIE: 75R, ARO: 16C", vbTextCompare) > 0 _
        Or InStr(1, texto, "BANDA: 225, S" & ChrW(201) & "RIE: 75R, ARO: 16C", vbTextCompare) > 0 Then
        ' Checa as duas grafias (com e sem acento em "SERIE") porque a
        ' descricao real pode vir de qualquer um dos dois jeitos -- o E
        ' acentuado e montado via ChrW, nao como literal no arquivo-fonte
        ' (ver nota de encoding no topo do modulo).
        ClassificarEstepeCompeticao = "LCV"
    ElseIf InStr(1, texto, "BANDA: 205, SERIE: 75R, ARO: 16C", vbTextCompare) > 0 _
        Or InStr(1, texto, "BANDA: 205, S" & ChrW(201) & "RIE: 75R, ARO: 16C", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "LCV"
    ElseIf InStr(1, texto, "CODIFICACAO 16C", vbTextCompare) > 0 _
        Or InStr(1, texto, "CODIFICA" & ChrW(199) & ChrW(195) & "O 16C", vbTextCompare) > 0 Then
        ' "CODIFICACAO"/"CODIFICACAO" com C e A montados via ChrW.
        ClassificarEstepeCompeticao = "LCV"
    ElseIf InStr(1, texto, "225/75R16C", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "LCV"
    ElseIf InStr(1, texto, "205/75R16C", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "LCV"
    ElseIf InStr(1, texto, "195/75R16C", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "LCV"
    ElseIf InStr(1, texto, "215/65R16C", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "LCV"
    ElseIf InStr(1, texto, "195/65R16C", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "LCV"
    Else
        ClassificarEstepeCompeticao = ""
    End If

End Function
