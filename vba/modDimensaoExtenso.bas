Attribute VB_Name = "modDimensaoExtenso"
Option Explicit

' ==========================================================================
' MODULO: modDimensaoExtenso
' Reconhecimento de medida de pneu ESCRITA POR EXTENSO em laudos/descricoes
' (ex: "BANDA: 175, SERIE: 75R, ARO: 13" em vez de "175/75R13").
'
' Isolado de proposito em modulo proprio: a extracao de DIMENSAO por
' catalogo (modDimensao.ExtrairDimensao) ja le corretamente a maior parte
' da base. Esse modulo so ACRESCENTA candidatos ao texto de busca -- nunca
' substitui nem decide sozinho o resultado final (quem decide se o
' candidato e valido continua sendo a comparacao com o catalogo em
' modDimensao). Assim, se um padrao novo aqui vier errado, o pior caso e
' "nao achou nada a mais" -- nao quebra o que ja funciona.
'
' Para adicionar um padrao novo: crie uma funcao privada "TentarPadraoX"
' seguindo o mesmo formato (recebe o texto ja em maiusculas/normalizado
' basico, devolve "" se nao bateu ou "LARGURA/PERFILRARO" se bateu) e
' registre a chamada dela em ExtrairDimensaoPorExtenso.
' ==========================================================================

' ==========================================================================
' Ponto de entrada: tenta cada padrao conhecido, em ordem, e devolve o
' primeiro candidato montado (ou "" se nenhum bateu). Nao decide sozinho --
' quem valida contra o catalogo e modDimensao.ExtrairDimensao.
' ==========================================================================
' somenteMinBr: as 3 funcoes deste modulo assumem construcao RADIAL e sempre
' montam o candidato com "/" e "R" (ex: "115/70R15"), nao importa o que veio
' na descricao original. Isso e certo pra laudos radiais escritos por
' extenso, mas errado pra pneus diagonais BR/MIN (que usam "-", ex:
' "115-70-15") -- forcaria um candidato num formato que nunca bate com o
' catalogo. Por isso, em modo BR/MIN, este modulo nao entra em acao.
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
' Padrao "BANDA ... SERIE ... ARO" (o mais comum em laudos INMETRO), com
' rotulos e separadores variaveis. Cobre, por exemplo:
'   "BANDA: 175, SERIE: 75R, ARO: 13"
'   "BANDA 195, SERIE 55 E ARO 16"
'   "LARGURA(BANDA): 205MM - PERFIL: 60R - ARO: 16"
'   "BANDA 115/SERIE 70/ARO 15"
' O separador entre partes e "qualquer coisa que nao seja digito" (virgula,
' "E", "MM", " - ", "/", ":"...), por isso "[^0-9]{1,40}" no meio em vez de
' exigir um separador exato. Rotulo de largura aceita "BANDA" ou
' "LARGURA(BANDA)"; rotulo de perfil aceita "SERIE"/"SERIE" ou "PERFIL".
' O "R" da serie e opcional na origem e sempre recolocado no resultado,
' ja que medida radial sempre usa R entre perfil e aro.
'
' PERFORMANCE: objeto RegExp cacheado com Static (criado uma unica vez,
' reaproveitado a cada chamada) em vez de recriado toda vez -- essa funcao
' roda uma vez por linha em toda a base fora de BR/MIN.
' ==========================================================================
Private Function TentarBandaSerieAro(textoNorm As String) As String
    Static regex As Object
    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = False
        regex.IgnoreCase = True
        ' "S[E" & ChrW(201) & "]RIE" cobre "SERIE" e "SERIE" (E montado via
        ' ChrW em vez de literal no arquivo-fonte -- ver nota de encoding
        ' no topo do modulo).
        regex.Pattern = "(?:LARGURA\(BANDA\)|BANDA)[^0-9]{0,10}(\d{2,3})[^0-9]{1,40}" & _
                        "(?:S[E" & ChrW(201) & "]RIE|SERIE|PERFIL)[^0-9]{0,10}(\d{1,3})R?[^0-9]{1,40}" & _
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
' Padrao "NOMENCLATURA (T125/80D15 95M)" -- laudo descreve a medida por
' extenso (largura, diametros, carga etc.) e no fim cita, entre parenteses,
' a nomenclatura tecnica compacta ja pronta (que pode usar "D" de diagonal
' em vez de "R" de radial -- nao alteramos essa letra, ela e parte da
' construcao do pneu, nao um erro de digitacao). So extrai o token de
' medida de dentro dos parenteses, sem tentar interpretar o resto do laudo
' (largura em mm, diametro em polegadas etc.) -- isso ficaria especifico
' demais e fragil para valer a pena.
'
' PERFORMANCE: os dois objetos RegExp (bloco e medida) sao cacheados com
' Static, mesma logica de TentarBandaSerieAro acima.
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
' Padrao "NNN NNLNN" -- banda e perfil separados so por ESPACO (sem "/"),
' com a letra do tipo de construcao/velocidade (R, Z, H...) colada direto
' no perfil, ex:
'   "DESCRICAO: 225 55Z19V" -> banda 225, perfil 55, aro 19 (o "Z" e o "V"
'   depois do aro sao indice de velocidade, nao fazem parte do GEOBOX)
'   "DESCRICAO 225 50R17 98W TL" -> banda 225, perfil 50, aro 17
' Sempre monta o resultado com "R" entre perfil e aro (medida radial usa
' R), independente da letra encontrada na origem -- mesmo criterio ja usado
' na limpeza de Z/ZR/ZRF em NormalizarFormatacaoBasica.
' Isolado neste modulo (so ACRESCENTA candidato ao texto de busca) porque
' esse padrao ja causou uma regressao na leitura de MARCA quando estava
' embutido direto no normalizador compartilhado (modDimensao) -- aqui ele
' nao tem esse risco, pois so e comparado contra o catalogo depois.
'
' PERFORMANCE: mesmo cache Static do objeto RegExp que as duas funcoes acima.
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
