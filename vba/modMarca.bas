Attribute VB_Name = "modMarca"
Option Explicit

' ==========================================================================
' MODULO: modMarca
' Extração da coluna MARCA a partir da DESCRIÇÃODA MERCADORIA.
' ==========================================================================

' ==========================================================================
' Extrai a MARCA: 1) checa primeiro o dicionário de EXCEÇÕES (erros de
' digitação/apelidos conhecidos, ex: "MGNUM" -> "MAGNUM", "X WORKS" ->
' "MICHELIN") — igual fazia a macro legada, antes de qualquer outra coisa.
' 2) Se não bater nenhuma exceção, busca por marca conhecida (arrMarcas, que
' reúne as marcas da Tabela de Referência + da aba "MarcasExtras"/
' "MarcasExtra", da mais longa p/ mais curta — ver
' modReferencia.CarregarTabelaReferencia).
'    - Em BR/MIN (somenteMinBr = True): busca de PALAVRA ISOLADA — só aceita
'      se a marca não estiver colada a outra letra (dígito/pontuação/espaço/
'      início-fim de texto contam como separador válido), com exceções pra
'      quando vem colada ao rótulo "MARCA" à esquerda ou a um rótulo de
'      campo conhecido à direita (LARGURA, SERIE, PEDIDO etc — ver
'      SeguidoPorRotuloCampo). Evita colisões tipo "GRI" dentro de
'      "AGRICOLA".
'    - Fora de BR/MIN: busca cega por substring livre, igual sempre foi.
' 3) Se não achar nada, retorna "".
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
' Procura, dentro de um texto, qual marca conhecida (arrMarcas, já ordenado
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
' True se o caractere for uma letra A-Z (maiúscula ou minúscula). String
' vazia (posição fora do texto — início/fim) conta como "não é letra", ou
' seja, já é fronteira válida.
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
' colada (sem espaço/pontuação no meio) — ex: "MODELO T510, MARCATRELLEBORG,
' INDICE...". Aqui "MARCA" é claramente um rótulo, não uma palavra que
' colide por acaso (como "AGRI" em "AGRICOLA") — então esse caso específico
' é aceito mesmo com letra colada, sem abrir mão da proteção geral.
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
' True se o trecho imediatamente DEPOIS de posDepois for um dos rótulos de
' campo conhecidos (LARGURA, SERIE, TAMANHO, PEDIDO, ITEM, FATURA, DIMENSAO,
' ARO, REGISTRO, MODELO, COR, CARGA, INDICE, REF, NCM, DIAMETRO) colado
' direto, sem separador — ex: "...MARCA: MICHELINLARGURA: 650",
' "...MARCA JUNGHEINRICHPEDIDO: 065034". Igual à ideia do rótulo "MARCA" à
' esquerda: esses nomes só aparecem colados como próximo campo do formulário,
' não colidem por acaso, então esse caso é aceito mesmo com letra colada à
' direita.
'
' PERFORMANCE/CORREÇÃO: a lista de rótulos é montada uma única vez (Static +
' flag de inicialização). Antes, a variável era declarada Static mas
' reatribuída incondicionalmente a cada chamada (o Array(...) rodava de
' novo toda vez) — ou seja, o cache nunca funcionava de fato. Essa função
' roda por candidato de marca testado, por linha, em bases BR/MIN, então o
' desperdício se multiplicava. Resultado final é idêntico: mesma lista,
' mesma ordem, mesma comparação.
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
' Procura, dentro do texto, a PRIMEIRA marca de arrMarcas (já ordenado do
' nome mais longo pro mais curto) que aparece como PALAVRA ISOLADA — sem
' regex: pra cada ocorrência (InStr, varrendo TODAS, não só a primeira),
' olha o caractere logo antes e logo depois e só aceita se nenhum dos dois
' for outra LETRA colada (dígito, espaço, pontuação, início/fim de texto
' contam como separador válido — ex: "MRL" em "MRL7.50-16" é aceito porque
' "7" não é letra; "GRI" dentro de "AGRICOLA" continua rejeitado porque "A"
' e "C" são letras), EXCETO quando o lado esquerdo for a palavra "MARCA"
' colada (ver PrecedidoPorRotuloMarca) ou o lado direito for um rótulo de
' campo conhecido colado (ver SeguidoPorRotuloCampo). Retorna "" se nenhuma
' marca bater. (Usado só em BR/MIN.)
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

                posBusca = posAchada + 1 ' tenta a próxima ocorrência dessa mesma marca
            Loop
        End If
    Next i
    BuscarMarcaPalavraSeca = ""
End Function
