Attribute VB_Name = "modDimensaoExtenso"
Option Explicit

' ==========================================================================
' MODULO: modDimensaoExtenso
' Reconhecimento de medida de pneu ESCRITA POR EXTENSO em laudos/descrições
' (ex: "BANDA: 175, SÉRIE: 75R, ARO: 13" em vez de "175/75R13").
'
' Isolado de propósito em módulo próprio: a extração de DIMENSÃO por
' catálogo (modDimensao.ExtrairDimensao) já lê corretamente a maior parte
' da base. Esse módulo só ACRESCENTA candidatos ao texto de busca — nunca
' substitui nem decide sozinho o resultado final (quem decide se o
' candidato é válido continua sendo a comparação com o catálogo em
' modDimensao). Assim, se um padrão novo aqui vier errado, o pior caso é
' "não achou nada a mais" — não quebra o que já funciona.
'
' Para adicionar um padrão novo: crie uma função privada "TentarPadraoX"
' seguindo o mesmo formato (recebe o texto já em maiúsculas/normalizado
' básico, devolve "" se não bateu ou "LARGURA/PERFILRARO" se bateu) e
' registre a chamada dela em ExtrairDimensaoPorExtenso.
' ==========================================================================

' ==========================================================================
' Ponto de entrada: tenta cada padrão conhecido, em ordem, e devolve o
' primeiro candidato montado (ou "" se nenhum bateu). Não decide sozinho —
' quem valida contra o catálogo é modDimensao.ExtrairDimensao.
' ==========================================================================
' somenteMinBr: as 3 funções deste módulo assumem construção RADIAL e sempre
' montam o candidato com "/" e "R" (ex: "115/70R15"), não importa o que veio
' na descrição original. Isso é certo pra laudos radiais escritos por
' extenso, mas errado pra pneus diagonais BR/MIN (que usam "-", ex:
' "115-70-15") — forçaria um candidato num formato que nunca bate com o
' catálogo. Por isso, em modo BR/MIN, este módulo não entra em ação.
Function ExtrairDimensaoPorExtenso(descricao As String, Optional somenteMinBr As Boolean = False) As String
    If somenteMinBr Then
        ExtrairDimensaoPorExtenso = ""
        Exit Function
    End If

    Dim textoNorm As String
    textoNorm = NormalizarFormatacaoBasica(descricao)

    Dim resultado As String

    resultado = TentarBandaSerieAro(textoNorm)
    If Len(resultado) > 0 Then
        ExtrairDimensaoPorExtenso = resultado
        Exit Function
    End If

    resultado = TentarNomenclaturaEntreParenteses(textoNorm)
    If Len(resultado) > 0 Then
        ExtrairDimensaoPorExtenso = resultado
        Exit Function
    End If

    resultado = TentarBandaPerfilAroSemBarra(textoNorm)
    If Len(resultado) > 0 Then
        ExtrairDimensaoPorExtenso = resultado
        Exit Function
    End If

    ExtrairDimensaoPorExtenso = ""
End Function

' ==========================================================================
' Padrão "BANDA ... SÉRIE ... ARO" (o mais comum em laudos INMETRO), com
' rótulos e separadores variáveis. Cobre, por exemplo:
'   "BANDA: 175, SÉRIE: 75R, ARO: 13"
'   "BANDA 195, SERIE 55 E ARO 16"
'   "LARGURA(BANDA): 205MM - PERFIL: 60R - ARO: 16"
'   "BANDA 115/SERIE 70/ARO 15"
' O separador entre partes é "qualquer coisa que não seja dígito" (vírgula,
' "E", "MM", " - ", "/", ":"...), por isso "[^0-9]{1,40}" no meio em vez de
' exigir um separador exato. Rótulo de largura aceita "BANDA" ou
' "LARGURA(BANDA)"; rótulo de perfil aceita "SÉRIE"/"SERIE" ou "PERFIL".
' O "R" da série é opcional na origem e sempre recolocado no resultado,
' já que medida radial sempre usa R entre perfil e aro.
'
' PERFORMANCE: objeto RegExp cacheado com Static (criado uma única vez,
' reaproveitado a cada chamada) em vez de recriado toda vez — essa função
' roda uma vez por linha em toda a base fora de BR/MIN.
' ==========================================================================
Private Function TentarBandaSerieAro(textoNorm As String) As String
    Static regex As Object
    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = False
        regex.IgnoreCase = True
        regex.Pattern = "(?:LARGURA\(BANDA\)|BANDA)[^0-9]{0,10}(\d{2,3})[^0-9]{1,40}" & _
                        "(?:S[EÉ]RIE|SERIE|PERFIL)[^0-9]{0,10}(\d{1,3})R?[^0-9]{1,40}" & _
                        "ARO[^0-9]{0,10}(\d{1,2}(?:\.\d)?)"
    End If

    If regex.Test(textoNorm) Then
        Dim m As Object
        Set m = regex.Execute(textoNorm)(0)
        TentarBandaSerieAro = m.SubMatches(0) & "/" & m.SubMatches(1) & "R" & m.SubMatches(2)
    Else
        TentarBandaSerieAro = ""
    End If
End Function

' ==========================================================================
' Padrão "NOMENCLATURA (T125/80D15 95M)" — laudo descreve a medida por
' extenso (largura, diâmetros, carga etc.) e no fim cita, entre parênteses,
' a nomenclatura técnica compacta já pronta (que pode usar "D" de diagonal
' em vez de "R" de radial — não alteramos essa letra, ela é parte da
' construção do pneu, não um erro de digitação). Só extrai o token de
' medida de dentro dos parênteses, sem tentar interpretar o resto do laudo
' (largura em mm, diâmetro em polegadas etc.) — isso ficaria específico
' demais e frágil para valer a pena.
'
' PERFORMANCE: os dois objetos RegExp (bloco e medida) são cacheados com
' Static, mesma lógica de TentarBandaSerieAro acima.
' ==========================================================================
Private Function TentarNomenclaturaEntreParenteses(textoNorm As String) As String
    Static regexBloco As Object
    If regexBloco Is Nothing Then
        Set regexBloco = CreateObject("VBScript.RegExp")
        regexBloco.Global = False
        regexBloco.IgnoreCase = True
        regexBloco.Pattern = "NOMENCLATURA\s*\(([^)]+)\)"
    End If

    If Not regexBloco.Test(textoNorm) Then
        TentarNomenclaturaEntreParenteses = ""
        Exit Function
    End If

    Dim bloco As String
    bloco = regexBloco.Execute(textoNorm)(0).SubMatches(0)

    Static regexMedida As Object
    If regexMedida Is Nothing Then
        Set regexMedida = CreateObject("VBScript.RegExp")
        regexMedida.Global = False
        regexMedida.IgnoreCase = True
        regexMedida.Pattern = "\d{2,3}/\d{1,3}[A-Z]\d{1,2}(?:\.\d)?"
    End If

    If regexMedida.Test(bloco) Then
        TentarNomenclaturaEntreParenteses = regexMedida.Execute(bloco)(0).Value
    Else
        TentarNomenclaturaEntreParenteses = ""
    End If
End Function

' ==========================================================================
' Padrão "NNN NNLNN" — banda e perfil separados só por ESPAÇO (sem "/"),
' com a letra do tipo de construção/velocidade (R, Z, H...) colada direto
' no perfil, ex:
'   "DESCRICAO: 225 55Z19V" -> banda 225, perfil 55, aro 19 (o "Z" e o "V"
'   depois do aro são índice de velocidade, não fazem parte do GEOBOX)
'   "DESCRICAO 225 50R17 98W TL" -> banda 225, perfil 50, aro 17
' Sempre monta o resultado com "R" entre perfil e aro (medida radial usa
' R), independente da letra encontrada na origem — mesmo critério já usado
' na limpeza de Z/ZR/ZRF em NormalizarFormatacaoBasica.
' Isolado neste módulo (só ACRESCENTA candidato ao texto de busca) porque
' esse padrão já causou uma regressão na leitura de MARCA quando estava
' embutido direto no normalizador compartilhado (modDimensao) — aqui ele
' não tem esse risco, pois só é comparado contra o catálogo depois.
'
' PERFORMANCE: mesmo cache Static do objeto RegExp que as duas funções acima.
' ==========================================================================
Private Function TentarBandaPerfilAroSemBarra(textoNorm As String) As String
    Static regex As Object
    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = False
        regex.IgnoreCase = True
        regex.Pattern = "(\d{3})\s(\d{2})[A-Z]{1,2}(\d{2})"
    End If

    If regex.Test(textoNorm) Then
        Dim m As Object
        Set m = regex.Execute(textoNorm)(0)
        TentarBandaPerfilAroSemBarra = m.SubMatches(0) & "/" & m.SubMatches(1) & "R" & m.SubMatches(2)
    Else
        TentarBandaPerfilAroSemBarra = ""
    End If
End Function
