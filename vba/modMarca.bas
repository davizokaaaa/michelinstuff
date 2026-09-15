Attribute VB_Name = "modMarca"
Option Explicit

' ==========================================================================
' MODULO: modMarca
' Extracao da coluna MARCA a partir da DESCRICAODA MERCADORIA.
' ==========================================================================

' ==========================================================================
' Extrai a MARCA: 1) checa primeiro o dicionario de EXCECOES (erros de
' digitacao/apelidos conhecidos, ex: "MGNUM" -> "MAGNUM", "X WORKS" ->
' "MICHELIN") -- igual fazia a macro legada, antes de qualquer outra coisa.
' 2) Se nao bater nenhuma excecao, busca por marca conhecida (arrMarcas, que
' reune as marcas da Tabela de Referencia + da aba "MarcasExtras"/
' "MarcasExtra", da mais longa p/ mais curta -- ver
' modReferencia.CarregarTabelaReferencia).
'    - Em BR/MIN (somenteMinBr = True): busca de PALAVRA ISOLADA -- so aceita
'      se a marca nao estiver colada a outra letra (digito/pontuacao/espaco/
'      inicio-fim de texto contam como separador valido), com excecoes pra
'      quando vem colada ao rotulo "MARCA" a esquerda ou a um rotulo de
'      campo conhecido a direita (LARGURA, SERIE, PEDIDO etc -- ver
'      SeguidoPorRotuloCampo). Evita colisoes tipo "GRI" dentro de
'      "AGRICOLA".
'    - Fora de BR/MIN: busca cega por substring livre, igual sempre foi.
' 3) Se nao achar nada, retorna "".
' ==========================================================================
Function ExtrairMarca(descricao As String, arrMarcas() As String, dicExcecoesMarca As Object, _
                       Optional somenteMinBr As Boolean = False) As String

    Dim chaveExc As Variant
    For Each chaveExc In dicExcecoesMarca.Keys
        If InStr(1, descricao, CStr(chaveExc), vbTextCompare) > 0 Then
            ExtrairMarca = dicExcecoesMarca(chaveExc)
            Exit Function
        End If
    Next chaveExc

    If somenteMinBr Then
        ExtrairMarca = BuscarMarcaPalavraSeca(descricao, arrMarcas)
    Else
        ExtrairMarca = BuscarMarcaConhecida(descricao, arrMarcas)
    End If

End Function

' ==========================================================================
' Procura, dentro de um texto, qual marca conhecida (arrMarcas, ja ordenado
' da mais longa p/ mais curta) aparece como substring livre. Retorna "" se
' nenhuma for encontrada. (Comportamento legado, usado fora de BR/MIN.)
' ==========================================================================
Function BuscarMarcaConhecida(texto As String, arrMarcas() As String) As String
    Dim i As Long
    For i = LBound(arrMarcas) To UBound(arrMarcas)
        If Len(arrMarcas(i)) > 0 Then
            If InStr(1, texto, arrMarcas(i), vbTextCompare) > 0 Then
                BuscarMarcaConhecida = arrMarcas(i)
                Exit Function
            End If
        End If
    Next i
    BuscarMarcaConhecida = ""
End Function

' ==========================================================================
' True se o caractere for uma letra A-Z (maiuscula ou minuscula). String
' vazia (posicao fora do texto -- inicio/fim) conta como "nao e letra", ou
' seja, ja e fronteira valida.
' ==========================================================================
Private Function EhLetra(c As String) As Boolean
    If Len(c) = 0 Then
        EhLetra = False
        Exit Function
    End If
    Dim cu As String
    cu = UCase(c)
    EhLetra = (cu >= "A" And cu <= "Z")
End Function

' ==========================================================================
' True se o trecho imediatamente ANTES de posAchada for a palavra "MARCA"
' colada (sem espaco/pontuacao no meio) -- ex: "MODELO T510, MARCATRELLEBORG,
' INDICE...". Aqui "MARCA" e claramente um rotulo, nao uma palavra que
' colide por acaso (como "AGRI" em "AGRICOLA") -- entao esse caso especifico
' e aceito mesmo com letra colada, sem abrir mao da protecao geral.
' ==========================================================================
Private Function PrecedidoPorRotuloMarca(texto As String, posAchada As Long) As Boolean
    Const ROTULO As String = "MARCA"
    Dim inicioRotulo As Long
    inicioRotulo = posAchada - Len(ROTULO)
    If inicioRotulo < 1 Then
        PrecedidoPorRotuloMarca = False
        Exit Function
    End If
    PrecedidoPorRotuloMarca = (UCase(Mid(texto, inicioRotulo, Len(ROTULO))) = ROTULO)
End Function

' ==========================================================================
' True se o trecho imediatamente DEPOIS de posDepois for um dos rotulos de
' campo conhecidos (LARGURA, SERIE, TAMANHO, PEDIDO, ITEM, FATURA, DIMENSAO,
' ARO, REGISTRO, MODELO, COR, CARGA, INDICE, REF, NCM, DIAMETRO) colado
' direto, sem separador -- ex: "...MARCA: MICHELINLARGURA: 650",
' "...MARCA JUNGHEINRICHPEDIDO: 065034". Igual a ideia do rotulo "MARCA" a
' esquerda: esses nomes so aparecem colados como proximo campo do formulario,
' nao colidem por acaso, entao esse caso e aceito mesmo com letra colada a
' direita.
'
' PERFORMANCE/CORRECAO: a lista de rotulos e montada uma unica vez (Static +
' flag de inicializacao). Antes, a variavel era declarada Static mas
' reatribuida incondicionalmente a cada chamada (o Array(...) rodava de
' novo toda vez) -- ou seja, o cache nunca funcionava de fato. Essa funcao
' roda por candidato de marca testado, por linha, em bases BR/MIN, entao o
' desperdicio se multiplicava. Resultado final e identico: mesma lista,
' mesma ordem, mesma comparacao.
' ==========================================================================
Private Function SeguidoPorRotuloCampo(texto As String, posDepois As Long) As Boolean
    Static arrRotulos As Variant
    Static rotulosInicializados As Boolean
    If Not rotulosInicializados Then
        arrRotulos = Array("LARGURA", "SERIE", "TAMANHO", "PEDIDO", "ITEM", "FATURA", "DIMENSAO", "ARO", _
                            "REGISTRO", "MODELO", "COR", "CARGA", "INDICE", "REF", "NCM", "DIAMETRO", "DESIGNACAO")
        rotulosInicializados = True
    End If

    Dim j As Long
    For j = LBound(arrRotulos) To UBound(arrRotulos)
        Dim rotulo As String
        rotulo = arrRotulos(j)
        If posDepois + Len(rotulo) - 1 <= Len(texto) Then
            If UCase(Mid(texto, posDepois, Len(rotulo))) = rotulo Then
                SeguidoPorRotuloCampo = True
                Exit Function
            End If
        End If
    Next j
    SeguidoPorRotuloCampo = False
End Function

' ==========================================================================
' Procura, dentro do texto, a PRIMEIRA marca de arrMarcas (ja ordenado do
' nome mais longo pro mais curto) que aparece como PALAVRA ISOLADA -- sem
' regex: pra cada ocorrencia (InStr, varrendo TODAS, nao so a primeira),
' olha o caractere logo antes e logo depois e so aceita se nenhum dos dois
' for outra LETRA colada (digito, espaco, pontuacao, inicio/fim de texto
' contam como separador valido -- ex: "MRL" em "MRL7.50-16" e aceito porque
' "7" nao e letra; "GRI" dentro de "AGRICOLA" continua rejeitado porque "A"
' e "C" sao letras), EXCETO quando o lado esquerdo for a palavra "MARCA"
' colada (ver PrecedidoPorRotuloMarca) ou o lado direito for um rotulo de
' campo conhecido colado (ver SeguidoPorRotuloCampo). Retorna "" se nenhuma
' marca bater. (Usado so em BR/MIN.)
' ==========================================================================
Function BuscarMarcaPalavraSeca(texto As String, arrMarcas() As String) As String
    Dim i As Long
    For i = LBound(arrMarcas) To UBound(arrMarcas)
        Dim cand As String
        cand = arrMarcas(i)
        If Len(cand) > 0 Then
            Dim posBusca As Long, posAchada As Long
            posBusca = 1
            Do
                posAchada = InStr(posBusca, texto, cand, vbTextCompare)
                If posAchada = 0 Then Exit Do

                Dim charAntes As String, charDepois As String
                If posAchada > 1 Then
                    charAntes = Mid(texto, posAchada - 1, 1)
                Else
                    charAntes = ""
                End If

                Dim posDepois As Long
                posDepois = posAchada + Len(cand)
                If posDepois <= Len(texto) Then
                    charDepois = Mid(texto, posDepois, 1)
                Else
                    charDepois = ""
                End If

                Dim ladoEsquerdoOk As Boolean, ladoDireitoOk As Boolean
                ladoEsquerdoOk = (Not EhLetra(charAntes)) Or PrecedidoPorRotuloMarca(texto, posAchada)
                ladoDireitoOk = (Not EhLetra(charDepois)) Or SeguidoPorRotuloCampo(texto, posDepois)

                If ladoEsquerdoOk And ladoDireitoOk Then
                    BuscarMarcaPalavraSeca = cand
                    Exit Function
                End If

                posBusca = posAchada + 1 ' tenta a proxima ocorrencia dessa mesma marca
            Loop
        End If
    Next i
    BuscarMarcaPalavraSeca = ""
End Function
