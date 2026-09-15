Attribute VB_Name = "modEstepeCompeticao"
Option Explicit

' ==========================================================================
' MODULO: modEstepeCompeticao
' Classifica ESTEPE, CAMPEONATO/COMPETIÇÃO e alguns casos específicos de
' LCV a partir da DESCRIÇÃODA MERCADORIA — grava numa coluna própria, sem
' se misturar com LP/TIPO PRODUTO/MARCA/GAMA/DIMENSÃO.
'
' ORIGEM: portado 1:1 da macro legada "Sub estepe()", que rodava sozinha,
' lendo a coluna H (hardcoded) e escrevendo na coluna T (hardcoded) da
' planilha ativa. Aqui a lógica de decisão é IDÊNTICA (mesmas palavras-
' chave, mesma ordem de prioridade em cascata If/ElseIf, mesmo
' vbTextCompare) — só isolada numa função pura que recebe o texto já
' extraído pelo orquestrador (modMain), pra rodar dentro do mesmo laço
' principal em vez de exigir uma segunda passada pela planilha inteira.
'
' IMPORTANTE (equivalência de resultado): a macro legada fazia
' UCase(ws.Cells(i, "H").Value) sem Trim. O orquestrador (modMain) já
' calcula "descricao" como UCase(Trim(...)) da coluna DESCRIÇÃODA
' MERCADORIA antes de chamar qualquer extração — Trim() só remove espaços
' nas PONTAS do texto, o que nunca muda o resultado de um InStr procurando
' uma palavra/trecho no MEIO do texto. Portanto passar "descricao" (em vez
' de reler a célula) preserva exatamente o mesmo resultado da macro legada,
' assumindo que a coluna H da planilha legada e a DESCRIÇÃODA MERCADORIA
' são a mesma coluna — CONFIRME isso ao validar lado a lado (ver PREMISSA
' abaixo).
'
' PREMISSA A CONFIRMAR: a macro legada lia da coluna H. Se, na planilha
' onde a macro legada rodava, a coluna H NÃO for a mesma coisa que
' "DESCRIÇÃODA MERCADORIA", ajuste a chamada em modMain para passar o
' texto certo.
' ==========================================================================

' ==========================================================================
' Classifica uma linha em ESTEPE / CAMPEONATO / COMPETICAO / LCV (ou ""),
' na mesma ordem de prioridade em cascata da macro legada — a primeira
' regra que bater decide, as de baixo nem são checadas. Espera o texto já
' em maiúsculas (mesmo formato que "descricao" chega em modMain); vbTextCompare
' no InStr já ignora caixa mesmo assim, então funciona igual com texto misto.
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
    ElseIf InStr(1, texto, "BANDA: 225, SÉRIE: 75R, ARO: 16C", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "LCV"
    ElseIf InStr(1, texto, "BANDA: 205, SÉRIE: 75R, ARO: 16C", vbTextCompare) > 0 Then
        ClassificarEstepeCompeticao = "LCV"
    ElseIf InStr(1, texto, "CODIFICAÇÃO 16C", vbTextCompare) > 0 Then
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
