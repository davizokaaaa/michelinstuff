Attribute VB_Name = "modGama"
Option Explicit

' ==========================================================================
' MODULO: modGama
' Extração da coluna GAMA a partir da DESCRIÇÃODA MERCADORIA.
' Mesmo modelo de modMarca.bas: exceções/abreviações primeiro (aba
' "ExcecoesGama" — ex: "PTNZ" -> "POTENZA", "TRNZ" -> "TURANZA"), depois
' busca pelo nome completo (ou por palavra-chave) da gama na descrição.
' ==========================================================================

' ==========================================================================
' Normaliza texto SÓ para efeito de comparação (nunca é gravado): remove
' espaço, barra, hífen e ponto, e coloca em maiúsculas.
' ==========================================================================
Private Function NormalizarParaComparacao(texto As String) As String
    Dim resultado As String
    resultado = UCase(texto)
    resultado = Replace(resultado, " ", "")
    resultado = Replace(resultado, "/", "")
    resultado = Replace(resultado, "-", "")
    resultado = Replace(resultado, ".", "")
    NormalizarParaComparacao = resultado
End Function

' ==========================================================================
' Mesma normalização de NormalizarParaComparacao, mas MANTÉM o espaço — só
' pra alimentar o teste de limite de palavra/número (\b) em candidatos
' "arriscados" (ver CandidatoGamaArriscado/CandidatoGamaBate). Remover o
' espaço (como a versão normal faz) grudaria palavras que estavam separadas
' no texto original, destruindo a fronteira que o \b precisa enxergar.
' ==========================================================================
Private Function NormalizarParaComparacaoComEspaco(texto As String) As String
    Dim resultado As String
    resultado = UCase(texto)
    resultado = Replace(resultado, "/", " ")
    resultado = Replace(resultado, "-", " ")
    resultado = Replace(resultado, ".", " ")
    NormalizarParaComparacaoComEspaco = resultado
End Function

' ==========================================================================
' Conta quantas PALAVRAS de um nome de gama (candidato) aparecem na
' descrição já normalizada. Quebra o candidato em palavras por espaço, "/",
' "-" e "." (ex: "FORZA H/T2" -> "FORZA", "H", "T2") e ignora palavras
' curtas demais (< 4 caracteres) pra não confiar num match tipo "AT" ou
' "H" sozinho, que apareceria em qualquer texto por acaso.
' ==========================================================================
Private Function ContarPalavrasBatendo(candidato As String, descNorm As String) As Long
    Dim texto As String
    texto = candidato
    texto = Replace(texto, "/", " ")
    texto = Replace(texto, "-", " ")
    texto = Replace(texto, ".", " ")

    Dim palavras() As String
    palavras = Split(texto, " ")

    Dim contagem As Long, p As Variant
    contagem = 0
    For Each p In palavras
        If Len(CStr(p)) >= 4 Then
            If InStr(1, descNorm, NormalizarParaComparacao(CStr(p)), vbTextCompare) > 0 Then
                contagem = contagem + 1
            End If
        End If
    Next p
    ContarPalavrasBatendo = contagem
End Function

' ==========================================================================
' Varre uma lista de candidatos de GAMA (Collection ou array de chaves de
' Dictionary) e devolve o que tiver MAIS palavras batendo na descrição
' (desempate: nome de gama mais longo). Devolve "" se nenhum candidato
' teve nenhuma palavra batendo.
' ==========================================================================
Private Function BuscarGamaPorPalavraChave(candidatos As Variant, descNorm As String) As String
    Dim melhorGama As String, melhorPontuacao As Long, melhorLen As Long
    melhorGama = ""
    melhorPontuacao = 0
    melhorLen = 0

    Dim item As Variant
    For Each item In candidatos
        Dim cand As String
        cand = CStr(item)
        If Len(cand) > 0 Then
            Dim pontos As Long
            pontos = ContarPalavrasBatendo(cand, descNorm)
            If pontos > 0 Then
                If pontos > melhorPontuacao Or (pontos = melhorPontuacao And Len(cand) > melhorLen) Then
                    melhorPontuacao = pontos
                    melhorGama = cand
                    melhorLen = Len(cand)
                End If
            End If
        End If
    Next item

    BuscarGamaPorPalavraChave = melhorGama
End Function

' ==========================================================================
' Verifica se um candidato de GAMA já normalizado (NormalizarParaComparacao)
' é "arriscado" o suficiente pra exigir limite de palavra/número na
' comparação: curto (<= 4 caracteres) ou puramente numérico. Nomes de gama
' assim colidem fácil com número solto no texto (medida em mm, capacidade
' de carga, código de peça...) — ex: gama "520" batendo dentro de
' "1.520MM". Só usado em modo BR/MIN (ver ExtrairGama).
' ==========================================================================
Private Function CandidatoGamaArriscado(candidatoNorm As String) As Boolean
    If Len(candidatoNorm) <= 4 Then
        CandidatoGamaArriscado = True
        Exit Function
    End If

    ' --- Regex cacheado (Static) — criar um objeto VBScript.RegExp novo a  ---
    ' --- cada chamada é caro, e essa função roda pra CADA candidato de     ---
    ' --- gama, em CADA linha da planilha. Sem cache, isso sozinho já       ---
    ' --- travava/deixava a macro lentíssima em bases grandes.              ---
    Static regexSoDigitos As Object
    If regexSoDigitos Is Nothing Then
        Set regexSoDigitos = CreateObject("VBScript.RegExp")
        regexSoDigitos.Pattern = "^\d+$"
    End If
    CandidatoGamaArriscado = regexSoDigitos.Test(candidatoNorm)
End Function

' ==========================================================================
' Escapa caracteres especiais de regex num texto literal (mesma lógica de
' modMarca.EscaparRegex, duplicada aqui pra não criar dependência cruzada
' entre módulos por uma função tão pequena).
' ==========================================================================
Private Function EscaparRegexGama(texto As String) As String
    Dim especiais As String
    especiais = "\^$.|?*+()[]{}"

    Dim resultado As String, i As Long, c As String
    resultado = ""
    For i = 1 To Len(texto)
        c = Mid(texto, i, 1)
        If InStr(especiais, c) > 0 Then
            resultado = resultado & "\" & c
        Else
            resultado = resultado & c
        End If
    Next i
    EscaparRegexGama = resultado
End Function

' ==========================================================================
' Testa se um candidato bate — substring livre no texto SEM espaço
' (descNorm, igual sempre foi), OU, se o candidato for "arriscado" (curto/
' numérico) E somenteMinBr, exigindo limite de palavra/número (\b) nos dois
' lados — testado contra descComEspaco (só maiúscula + remove barra/traço/
' ponto, MANTÉM espaço), porque descNorm já removeu os espaços e grudaria
' duas palavras que estavam separadas no texto original (ex: "RODAGEM
' VSDL" -> "RODAGEMVSDL"), destruindo a fronteira que o \b precisa achar.
' ==========================================================================
Private Function CandidatoGamaBate(candidatoNorm As String, descNorm As String, descComEspaco As String, somenteMinBr As Boolean) As Boolean
    If somenteMinBr And CandidatoGamaArriscado(candidatoNorm) Then
        Static regex As Object
        If regex Is Nothing Then
            Set regex = CreateObject("VBScript.RegExp")
            regex.Global = False
            regex.IgnoreCase = True
        End If
        regex.Pattern = "\b" & EscaparRegexGama(candidatoNorm) & "\b"
        CandidatoGamaBate = regex.Test(descComEspaco)
    Else
        CandidatoGamaBate = (InStr(1, descNorm, candidatoNorm, vbTextCompare) > 0)
    End If
End Function

' ==========================================================================
' Varre uma lista de candidatos de GAMA (Collection ou array de chaves de
' Dictionary) comparando o nome completo (normalizado) contra a descrição, e
' devolve o que BATER e for o MAIS LONGO entre todos — não o primeiro que
' bater (a ordem de iteração de um Dictionary é arbitrária, então "ficar no
' primeiro" deixava uma gama curta/coincidente vencer uma mais longa e
' correta só por sorte de ordenação). Em modo BR/MIN, candidatos curtos/
' numéricos (ver CandidatoGamaArriscado) só contam com limite de palavra.
' Devolve "" se nenhum candidato bateu.
' ==========================================================================
Private Function BuscarGamaPorNomeCompleto(candidatos As Variant, descNorm As String, descComEspaco As String, somenteMinBr As Boolean) As String
    Dim melhorGama As String, melhorLen As Long
    melhorGama = ""
    melhorLen = 0

    Dim item As Variant
    For Each item In candidatos
        Dim cand As String
        cand = CStr(item)
        If Len(cand) > 0 Then
            Dim candNorm As String
            candNorm = NormalizarParaComparacao(cand)
            If Len(candNorm) > melhorLen Then
                If CandidatoGamaBate(candNorm, descNorm, descComEspaco, somenteMinBr) Then
                    melhorLen = Len(candNorm)
                    melhorGama = cand
                End If
            End If
        End If
    Next item

    BuscarGamaPorNomeCompleto = melhorGama
End Function

' ==========================================================================
' Extrai a GAMA em camadas, da mais confiável pra mais permissiva:
'   1) Exceções/abreviações (dicExcecoesGama, aba "ExcecoesGama").
'   2) Nome completo da gama (normalizado contra espaço/barra/hífen/ponto),
'      primeiro só entre as gamas DAQUELA marca, depois busca ampla — fica
'      com o match MAIS LONGO entre as candidatas, não o primeiro que bater
'      (ver BuscarGamaPorNomeCompleto).
'   3) Se não achou nada assim, cai numa busca por PALAVRA-CHAVE: quebra
'      cada gama conhecida em palavras e vê se pelo menos uma (>= 4
'      caracteres) aparece na descrição — primeiro só nas gamas da marca já
'      identificada, depois busca ampla. Mais permissivo, mas escala melhor
'      que ficar caçando exceção de pontuação linha por linha; se overmatch
'      demais em algum caso real, dá pra restringir de volta esse passo.
' Se nada bateu em nenhuma camada, retorna "".
'
' GAMA é tratada como praticamente exclusiva de uma marca: sempre que ela é
' descoberta (em qualquer camada acima) e a MARCA da linha ainda está
' vazia, a marca é preenchida também (ByRef) com a dona daquela gama
' (dicMarcaPorGama, montado em modReferencia a partir da própria Tabela de
' Referência).
'
' somenteMinBr: em BR/MIN, a camada 2 exige limite de palavra/número pra
' candidatas curtas (<=4 caracteres) ou puramente numéricas — evita que uma
' gama tipo "520" bata só por estar dentro de "1.520MM". Fora de BR/MIN,
' mantém o comportamento antigo (substring livre) em toda a camada 2.
'
' Em BR/MIN, GAMA só é procurada se a MARCA da linha já tiver sido
' encontrada antes (marca não vazia) — uma gama "solta" sem marca conhecida
' tende a ser falso positivo (número da descrição batendo com gama numérica
' tipo "550", ou palavra-chave genérica tipo "FLORESTAL" sem ligação real
' com nenhuma marca confirmada). Fora de BR/MIN, mantém o comportamento
' antigo (procura mesmo com marca vazia).
' ==========================================================================
Function ExtrairGama(descricao As String, ByRef marca As String, _
                      dicGamasPorMarca As Object, dicGamaGlobalUnicos As Object, _
                      dicMarcaPorGama As Object, dicExcecoesGama As Object, _
                      Optional somenteMinBr As Boolean = False) As String

    If somenteMinBr And Len(marca) = 0 Then
        ExtrairGama = ""
        Exit Function
    End If

    Dim descNorm As String
    descNorm = NormalizarParaComparacao(descricao)

    Dim descComEspaco As String
    descComEspaco = NormalizarParaComparacaoComEspaco(descricao)

    ' --- Camada 1: exceções/abreviações ---
    Dim chaveExc As Variant
    For Each chaveExc In dicExcecoesGama.Keys
        If InStr(1, descricao, CStr(chaveExc), vbTextCompare) > 0 Then
            ExtrairGama = dicExcecoesGama(chaveExc)
            If Len(marca) = 0 And dicMarcaPorGama.Exists(ExtrairGama) Then
                marca = CStr(dicMarcaPorGama(ExtrairGama))
            End If
            Exit Function
        End If
    Next chaveExc

    ' --- Camada 2: nome completo (normalizado), marca conhecida primeiro ---
    Dim resultadoNomeCompleto As String
    If Len(marca) > 0 Then
        If dicGamasPorMarca.Exists(marca) Then
            resultadoNomeCompleto = BuscarGamaPorNomeCompleto(dicGamasPorMarca(marca), descNorm, descComEspaco, somenteMinBr)
            If Len(resultadoNomeCompleto) > 0 Then
                ExtrairGama = resultadoNomeCompleto
                Exit Function
            End If
        End If
    End If

    ' --- Camada 2: nome completo (normalizado), busca ampla ---
    resultadoNomeCompleto = BuscarGamaPorNomeCompleto(dicGamaGlobalUnicos.Keys, descNorm, descComEspaco, somenteMinBr)
    If Len(resultadoNomeCompleto) > 0 Then
        ExtrairGama = resultadoNomeCompleto
        If Len(marca) = 0 And dicMarcaPorGama.Exists(ExtrairGama) Then
            marca = CStr(dicMarcaPorGama(ExtrairGama))
        End If
        Exit Function
    End If

    ' --- Camada 3: palavra-chave, marca conhecida primeiro ---
    Dim resultadoPalavra As String
    If Len(marca) > 0 Then
        If dicGamasPorMarca.Exists(marca) Then
            resultadoPalavra = BuscarGamaPorPalavraChave(dicGamasPorMarca(marca), descNorm)
            If Len(resultadoPalavra) > 0 Then
                ExtrairGama = resultadoPalavra
                Exit Function
            End If
        End If
    End If

    ' --- Camada 3: palavra-chave, busca ampla ---
    resultadoPalavra = BuscarGamaPorPalavraChave(dicGamaGlobalUnicos.Keys, descNorm)
    If Len(resultadoPalavra) > 0 Then
        ExtrairGama = resultadoPalavra
        If Len(marca) = 0 And dicMarcaPorGama.Exists(ExtrairGama) Then
            marca = CStr(dicMarcaPorGama(ExtrairGama))
        End If
        Exit Function
    End If

    ExtrairGama = ""
End Function
