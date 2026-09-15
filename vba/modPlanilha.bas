Attribute VB_Name = "modPlanilha"
Option Explicit

' ==========================================================================
' MODULO: modPlanilha
' Utilitarios genericos de planilha: localizar/criar colunas por cabecalho,
' normalizar texto de cabecalho para comparacao, achar ultima linha usada.
' ==========================================================================

' ==========================================================================
' Retorna a maior "ultima linha com dado" entre um conjunto de colunas,
' para nao depender de uma unica coluna (que pode estar vazia por acaso
' nas ultimas linhas, ou por causa de reordenacao de colunas).
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
' mesmo cabecalho (ex: DIMENSAO ou GEOBOX). Retorna a primeira que achar.
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
' Localiza a coluna pelo cabecalho; se nao existir, CRIA uma nova coluna
' na primeira posicao vazia a direita dos dados e escreve o cabecalho nela.
' Usada para as colunas de SAIDA, que a macro tem autonomia para criar.
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
' Normaliza texto para comparacao de cabecalhos: maiusculas, so letras A-Z
' e digitos 0-9 -- QUALQUER outro caractere (espaco, acento, cedilha, barra,
' hifen, simbolo) e descartado.
'
' CORRECAO: a versao anterior tentava mapear cada acento conhecido (A, A,
' C...) para a letra sem acento. Isso quebra na pratica porque o texto
' colado no VBA Editor pode corromper acentos de formas diferentes conforme
' o caminho que o texto percorreu (copiar/colar, salvar, reabrir) -- os
' bytes que sobram nem sempre sao os mesmos caracteres que o mapa espera,
' entao a substituicao simplesmente nao acontecia e a coluna nao era
' encontrada (mesmo ela existindo, com o nome "certo", no cabecalho real).
' Descartar QUALQUER caractere fora de A-Z/0-9 (em vez de tentar adivinhar
' qual acento e qual) resolve isso na raiz: nao importa como "DESCRICAO"
' corrompeu, o "esqueleto" de letras simples (D-E-S-C-R-I...) sempre sobra
' igual dos dois lados da comparacao.
' ==========================================================================
Function NormalizarTexto(txt As String) As String
    Dim resultado As String
    resultado = UCase(Trim(txt))

    Dim saida As String, i As Long, c As String
    saida = ""
    For i = 1 To Len(resultado)
        c = Mid(resultado, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Then
            saida = saida & c
        End If
    Next i

    NormalizarTexto = saida
End Function

' ==========================================================================
' Localiza a coluna cujo cabecalho (linha 1) bate com o nome informado.
' Se houver MAIS DE UMA coluna com o mesmo nome de cabecalho (ex: uma
' coluna vazia remanescente de reordenacao), escolhe a que tiver mais
' celulas preenchidas abaixo do cabecalho -- evita pegar uma coluna "fantasma".
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
                    " vezes no cabecalho. Escolhida a coluna " & melhorCol & _
                    " (mais preenchida, " & melhorQtdPreenchida & " celulas com dado)."
    End If

    LocalizarColuna = melhorCol
End Function
