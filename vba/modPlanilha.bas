Attribute VB_Name = "modPlanilha"
Option Explicit

' ==========================================================================
' MODULO: modPlanilha
' Utilitarios genericos de planilha: localizar/criar colunas por cabecalho,
' normalizar texto de cabecalho para comparacao, achar ultima linha usada.
' ==========================================================================

' ==========================================================================
' Retorna a maior "última linha com dado" entre um conjunto de colunas,
' para não depender de uma única coluna (que pode estar vazia por acaso
' nas últimas linhas, ou por causa de reordenação de colunas).
' ==========================================================================
Function UltimaLinhaEntreColunas(ws As Worksheet, colunas As Variant) As Long
    Dim maiorLinha As Long, linhaAtual As Long, i As Long
    maiorLinha = 1

    For i = LBound(colunas) To UBound(colunas)
        If colunas(i) > 0 Then
            linhaAtual = ws.Cells(ws.Rows.Count, colunas(i)).End(xlUp).Row
            If linhaAtual > maiorLinha Then maiorLinha = linhaAtual
        End If
    Next i

    UltimaLinhaEntreColunas = maiorLinha
End Function

' ==========================================================================
' Como LocalizarColuna, mas aceita uma lista de nomes alternativos para o
' mesmo cabeçalho (ex: DIMENSÃO ou GEOBOX). Retorna a primeira que achar.
' ==========================================================================
Function LocalizarColunaAlternativas(ws As Worksheet, nomes As Variant) As Long
    Dim i As Long, resultado As Long

    For i = LBound(nomes) To UBound(nomes)
        resultado = LocalizarColuna(ws, CStr(nomes(i)))
        If resultado > 0 Then
            LocalizarColunaAlternativas = resultado
            Exit Function
        End If
    Next i

    LocalizarColunaAlternativas = 0
End Function

' ==========================================================================
' Como LocalizarColunaAlternativas, mas se nenhum dos nomes existir, CRIA
' a coluna usando o nomeParaCriar informado.
' ==========================================================================
Function LocalizarColunaAlternativasOuCriar(ws As Worksheet, nomes As Variant, nomeParaCriar As String) As Long
    Dim resultado As Long
    resultado = LocalizarColunaAlternativas(ws, nomes)

    If resultado > 0 Then
        LocalizarColunaAlternativasOuCriar = resultado
        Exit Function
    End If

    Dim novaCol As Long
    novaCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, novaCol).Value = nomeParaCriar

    LocalizarColunaAlternativasOuCriar = novaCol
End Function

' ==========================================================================
' Localiza a coluna pelo cabeçalho; se não existir, CRIA uma nova coluna
' na primeira posição vazia à direita dos dados e escreve o cabeçalho nela.
' Usada para as colunas de SAÍDA, que a macro tem autonomia para criar.
' ==========================================================================
Function LocalizarOuCriarColuna(ws As Worksheet, nomeColuna As String) As Long
    Dim resultado As Long
    resultado = LocalizarColuna(ws, nomeColuna)

    If resultado > 0 Then
        LocalizarOuCriarColuna = resultado
        Exit Function
    End If

    Dim novaCol As Long
    novaCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, novaCol).Value = nomeColuna

    LocalizarOuCriarColuna = novaCol
End Function

' ==========================================================================
' Normaliza texto para comparação de cabeçalhos: maiúsculas, sem espaços,
' sem acentos/cedilha. Evita falhas de comparação por causa de acentos
' que às vezes corrompem ao colar/salvar código VBA (á, é, í, ç etc.).
' ==========================================================================
Function NormalizarTexto(txt As String) As String
    Dim resultado As String
    resultado = UCase(Trim(txt))

    resultado = Replace(resultado, "Á", "A")
    resultado = Replace(resultado, "À", "A")
    resultado = Replace(resultado, "Ã", "A")
    resultado = Replace(resultado, "Â", "A")
    resultado = Replace(resultado, "É", "E")
    resultado = Replace(resultado, "Ê", "E")
    resultado = Replace(resultado, "Í", "I")
    resultado = Replace(resultado, "Ó", "O")
    resultado = Replace(resultado, "Ô", "O")
    resultado = Replace(resultado, "Õ", "O")
    resultado = Replace(resultado, "Ú", "U")
    resultado = Replace(resultado, "Ü", "U")
    resultado = Replace(resultado, "Ç", "C")

    resultado = Replace(resultado, " ", "")
    resultado = Replace(resultado, Chr(160), "") ' espaço não separável
    resultado = Replace(resultado, "/", "")
    resultado = Replace(resultado, "-", "")

    NormalizarTexto = resultado
End Function

' ==========================================================================
' Localiza a coluna cujo cabeçalho (linha 1) bate com o nome informado.
' Se houver MAIS DE UMA coluna com o mesmo nome de cabeçalho (ex: uma
' coluna vazia remanescente de reordenação), escolhe a que tiver mais
' células preenchidas abaixo do cabeçalho — evita pegar uma coluna "fantasma".
' ==========================================================================
Function LocalizarColuna(ws As Worksheet, nomeColuna As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Dim melhorCol As Long, melhorQtdPreenchida As Long
    Dim qtdColunasAchadas As Long
    melhorCol = 0
    melhorQtdPreenchida = -1
    qtdColunasAchadas = 0

    For c = 1 To lastCol
        If NormalizarTexto(CStr(ws.Cells(1, c).Value)) = NormalizarTexto(nomeColuna) Then
            qtdColunasAchadas = qtdColunasAchadas + 1

            Dim ultimaLinhaColuna As Long, qtdPreenchida As Long
            ultimaLinhaColuna = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
            qtdPreenchida = Application.WorksheetFunction.CountA( _
                ws.Range(ws.Cells(2, c), ws.Cells(Application.Max(ultimaLinhaColuna, 2), c)))

            If qtdPreenchida > melhorQtdPreenchida Then
                melhorQtdPreenchida = qtdPreenchida
                melhorCol = c
            End If
        End If
    Next c

    If qtdColunasAchadas > 1 Then
        Debug.Print "AVISO: coluna """ & nomeColuna & """ encontrada " & qtdColunasAchadas & _
                    " vezes no cabeçalho. Escolhida a coluna " & melhorCol & _
                    " (mais preenchida, " & melhorQtdPreenchida & " células com dado)."
    End If

    LocalizarColuna = melhorCol
End Function
