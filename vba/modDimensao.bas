Attribute VB_Name = "modDimensao"
Option Explicit

' ==========================================================================
' MODULO: modDimensao
' Normalização de formatação de texto de medida, extração de DIMENSÃO
' (GEOBOX) e de ARO a partir da descrição.
' ==========================================================================

' ==========================================================================
' Extrai a DIMENSÃO (GEOBOX): 1) se a marca já foi identificada, procura
' primeiro só entre as medidas conhecidas DAQUELA marca (mais rápido e mais
' preciso). 2) Se não achar assim, faz busca ampla em todas as medidas
' conhecidas de todas as marcas. 3) Se ainda assim não achar, retorna "".
' Em caso de mais de uma medida batendo no texto, fica com a mais longa
' (evita pegar "R17" quando na verdade é "R17.5"). O texto é normalizado
' antes (vírgula->ponto, ×/*->X, espaços ao redor de R/barra/traço) para
' bater com mais variações de formatação. Também compara contra uma versão
' com "D" (construção diagonal, ex: "135/80D15") trocado por "R", pois a
' Tabela de Referência às vezes só tem a variante radial cadastrada — nesse
' caso grava o valor com "R" mesmo (como está na Referência), não o "D"
' original da descrição.
' ==========================================================================
' Tamanho mínimo aceito pra um candidato de DIMENSÃO ser considerado um
' match válido. Protege contra entradas curtas demais/corrompidas na
' Tabela de Referência (ex: uma célula com só "R") baterem por acidente
' em qualquer trecho da descrição e virarem resultado sem sentido — a
' menor medida real (tipo "9R20") ainda tem pelo menos 4 caracteres.
Const TAMANHO_MINIMO_DIMENSAO As Long = 4

' ==========================================================================
' DE-PARA CEGO de GEOBOX — usado SÓ no modo BR/MIN (somenteMinBr = True).
' Validado num teste isolado (modTesteMatchExato) contra a base real antes
' de virar padrão aqui: match LITERAL do valor exatamente como cadastrado na
' Tabela de Referência dentro da descrição, sem tentar reconhecer "formato"
' de medida (largura/perfil/aro) — só uma limpeza puramente de formatação,
' aplicada nos DOIS lados antes de comparar: vírgula -> ponto decimal.
' (A regra de tratar ".50" e ".5" como iguais foi removida daqui — estava
' gerando GEOBOX errado ao juntar medidas diferentes que só coincidem depois
' de arredondar o zero.) E o espaço é testado de duas formas (pode ser
' decorativo OU estar no lugar de uma barra que faltou): espaço removido, e
' espaço virando "/". Nunca inventa nem adivinha — só bate se existir
' exatamente (depois dessa limpeza) no dicionário da Tabela de Referência.
' ==========================================================================

Function ExtrairDimensao(descricao As String, marca As String, _
                          dicGeoboxPorMarca As Object, dicGeoboxGlobalUnicos As Object, _
                          dicPadroesLegado As Object, _
                          Optional somenteMinBr As Boolean = False) As String

    Dim textoNorm As String
    textoNorm = NormalizarTextoDimensao(descricao, somenteMinBr)

    ' --- Segunda versão do texto, com TODOS os espaços removidos. Cobre   ---
    ' --- casos como "245/35 ZR19" (espaço antes do Z, escapa da regra de  ---
    ' --- Z/ZR/ZRF que exige dígito colado) ou até números quebrados por   ---
    ' --- engano no meio ("22 5/55R17" -> "225/55R17" depois de juntar).   ---
    ' --- Reaplica a limpeza de Z/RF depois de juntar, pra pegar casos que ---
    ' --- só viram "dígito colado com Z/RF" depois da junção.             ---
    Dim textoSemEspaco As String
    textoSemEspaco = Replace(textoNorm, " ", "")
    textoSemEspaco = NormalizarFormatacaoBasica(textoSemEspaco)

    ' --- Medidas escritas por extenso (laudos INMETRO, "BANDA/SÉRIE/ARO" ---
    ' --- em vários formatos) — lógica isolada em modDimensaoExtenso pra   ---
    ' --- poder evoluir/reverter sem mexer na extração já validada. Só    ---
    ' --- ACRESCENTA o candidato montado ao texto de busca; a validação   ---
    ' --- (existe no catálogo? tamanho mínimo?) continua sendo feita nos  ---
    ' --- mesmos laços abaixo, como qualquer outro trecho da descrição.   ---
    Dim candidatoExtenso As String
    candidatoExtenso = ExtrairDimensaoPorExtenso(descricao, somenteMinBr)
    If Len(candidatoExtenso) > 0 Then
        textoNorm = textoNorm & " " & candidatoExtenso
        textoSemEspaco = textoSemEspaco & Replace(candidatoExtenso, " ", "")
    End If

    ' --- Versões com "D" (construção diagonal) trocado por "R", só para   ---
    ' --- efeito de COMPARAÇÃO com o catálogo. Existem aros com construção ---
    ' --- diagonal na descrição (ex: "135/80D15") cujo cadastro na Tabela  ---
    ' --- de Referência só tem a variante radial ("135/80R15") — sem isso, ---
    ' --- esses aros nunca seriam encontrados. Não reaplica em nenhum      ---
    ' --- outro lugar do código (marca, gama etc.), só nesta comparação.   ---
    Dim textoNormDparaR As String, textoSemEspacoDparaR As String
    textoNormDparaR = Replace(textoNorm, "D", "R")
    textoSemEspacoDparaR = Replace(textoSemEspaco, "D", "R")

    Dim melhor As String, melhorLen As Long
    melhor = ""
    melhorLen = 0

    ' --- Passo 0: DE-PARA cego — SÓ no modo BR/MIN --- Fora desse modo, a
    ' --- extração continua 100% igual à que já funciona (busca literal por ---
    ' --- substring abaixo), sem nenhuma mudança de comportamento.          ---
    If somenteMinBr Then
        Dim descComPonto As String
        descComPonto = Replace(descricao, ",", ".")

        Dim descBrSemEspaco As String, descBrEspacoBarra As String
        descBrSemEspaco = Replace(descComPonto, " ", "")
        descBrEspacoBarra = Replace(descComPonto, " ", "/")

        Dim chaveBr As Variant
        For Each chaveBr In dicGeoboxGlobalUnicos.Keys
            Dim geoBrComPonto As String
            geoBrComPonto = Replace(CStr(chaveBr), ",", ".")

            Dim geoBrSemEspaco As String
            geoBrSemEspaco = Replace(geoBrComPonto, " ", "")

            If Len(geoBrSemEspaco) > melhorLen Then
                If InStr(1, descBrSemEspaco, geoBrSemEspaco, vbTextCompare) > 0 _
                   Or InStr(1, descBrEspacoBarra, geoBrSemEspaco, vbTextCompare) > 0 Then
                    melhorLen = Len(geoBrSemEspaco)
                    melhor = CStr(chaveBr)
                End If
            End If
        Next chaveBr
    End If

    ' --- Medidas conhecidas DAQUELA marca E busca ampla (qualquer marca) ---
    ' IMPORTANTE: as duas buscas SEMPRE rodam, contra as DUAS versões do
    ' texto (normal e sem espaço), e fica com a mais longa entre todas —
    ' nunca aceita cegamente o que a busca por marca achar primeiro.
    ' Isso evita que uma medida catalogada pra aquela marca, mas SEM relação
    ' com o produto da linha, bata por coincidência em algum trecho solto da
    ' descrição (números de registro, certificado, código de família etc.)
    ' e "roube" a vaga de uma medida mais longa e mais correta que só
    ' apareceria na busca ampla (cadastrada sob outra marca).
    If Len(marca) > 0 Then
        If dicGeoboxPorMarca.Exists(marca) Then
            Dim candidatos As Collection
            Set candidatos = dicGeoboxPorMarca(marca)

            Dim geoCand As Variant
            For Each geoCand In candidatos
                If Len(CStr(geoCand)) >= TAMANHO_MINIMO_DIMENSAO Then
                    If Len(CStr(geoCand)) > melhorLen Then
                        If InStr(1, textoNorm, CStr(geoCand), vbTextCompare) > 0 _
                           Or InStr(1, textoSemEspaco, CStr(geoCand), vbTextCompare) > 0 _
                           Or InStr(1, textoNormDparaR, CStr(geoCand), vbTextCompare) > 0 _
                           Or InStr(1, textoSemEspacoDparaR, CStr(geoCand), vbTextCompare) > 0 Then
                            melhorLen = Len(CStr(geoCand))
                            melhor = CStr(geoCand)
                        End If
                    End If
                End If
            Next geoCand
        End If
    End If

    ' --- Busca ampla em qualquer medida conhecida (Tabela de Referência) ---
    ' --- Só fora de BR/MIN: em BR/MIN, o Passo 0 (DE-PARA cego, acima) já    ---
    ' --- varreu esse MESMO dicGeoboxGlobalUnicos inteiro pra achar o melhor ---
    ' --- match; repetir a varredura aqui (com 4 InStr por item) é trabalho  ---
    ' --- em dobro por linha — e com milhares de GEOBOX únicos numa base     ---
    ' --- BR/MIN, isso é o que deixava a macro lentíssima/travando.          ---
    If Not somenteMinBr Then
        Dim chaveG As Variant
        For Each chaveG In dicGeoboxGlobalUnicos.Keys
            If Len(CStr(chaveG)) >= TAMANHO_MINIMO_DIMENSAO Then
                If Len(CStr(chaveG)) > melhorLen Then
                    If InStr(1, textoNorm, CStr(chaveG), vbTextCompare) > 0 _
                       Or InStr(1, textoSemEspaco, CStr(chaveG), vbTextCompare) > 0 _
                       Or InStr(1, textoNormDparaR, CStr(chaveG), vbTextCompare) > 0 _
                       Or InStr(1, textoSemEspacoDparaR, CStr(chaveG), vbTextCompare) > 0 Then
                        melhorLen = Len(CStr(chaveG))
                        melhor = CStr(chaveG)
                    End If
                End If
            End If
        Next chaveG
    End If

    If melhorLen > 0 Then
        ExtrairDimensao = melhor
        Exit Function
    End If

    ' --- Passo 3: lista específica vinda da macro legada (padrão -> valor oficial) ---
    Dim chaveL As Variant
    For Each chaveL In dicPadroesLegado.Keys
        If Len(CStr(chaveL)) >= TAMANHO_MINIMO_DIMENSAO Then
            If Len(CStr(chaveL)) > melhorLen Then
                If InStr(1, textoNorm, CStr(chaveL), vbTextCompare) > 0 _
                   Or InStr(1, textoSemEspaco, CStr(chaveL), vbTextCompare) > 0 _
                   Or InStr(1, textoNormDparaR, CStr(chaveL), vbTextCompare) > 0 _
                   Or InStr(1, textoSemEspacoDparaR, CStr(chaveL), vbTextCompare) > 0 Then
                    melhorLen = Len(CStr(chaveL))
                    melhor = dicPadroesLegado(chaveL)
                End If
            End If
        End If
    Next chaveL

    ExtrairDimensao = melhor
End Function

' ==========================================================================
' Normalização BÁSICA de formatação (compartilhada): maiúsculas, vírgula
' decimal -> ponto, ×/* -> X, colapsa espaços soltos ao redor de R/barra/
' traço. Usada tanto no texto da descrição quanto nos valores de GEOBOX
' lidos da própria Tabela de Referência — os DOIS lados da comparação
' precisam estar no mesmo formato, senão "175/75 R13" (com espaço, na
' tabela) nunca bate com "175/75R13" (sem espaço, já limpo na descrição).
'
' PERFORMANCE: os dois objetos RegExp usados aqui são cacheados com Static
' (criados uma única vez por sessão do Excel, não a cada chamada). Essa
' função roda várias vezes por linha, em toda a base — recriar o objeto
' COM do RegExp a cada chamada (como era antes) é um dos maiores custos
' de desempenho do laço principal. O padrão/comportamento do regex não
' muda: só o objeto passa a ser reaproveitado.
' ==========================================================================
Function NormalizarFormatacaoBasica(texto As String) As String
    Dim resultado As String
    resultado = UCase(texto)

    resultado = Replace(resultado, ",", ".")
    resultado = Replace(resultado, ChrW(215), "X") ' × (sinal de multiplicação, U+00D7)
    resultado = Replace(resultado, Chr(215), "X")   ' fallback caso venha como ANSI
    resultado = Replace(resultado, "*", "X")

    ' Traços "parecidos" com hífen, mas que são caracteres Unicode diferentes
    ' (comuns em texto colado de Excel/Word/PDF) — todos viram o hífen comum
    ' "-" (U+002D). Sem isso, um GEOBOX como "33X12-20/7.50" registrado com
    ' um desses traços nunca bate com a mesma medida digitada com hífen
    ' normal na descrição (ou vice-versa), mesmo sendo visualmente idênticos.
    resultado = Replace(resultado, ChrW(8211), "-") ' en dash (U+2013)
    resultado = Replace(resultado, ChrW(8212), "-") ' em dash (U+2014)
    resultado = Replace(resultado, ChrW(8722), "-") ' sinal de menos matematico (U+2212)
    resultado = Replace(resultado, ChrW(8209), "-") ' hifen nao separavel (U+2011)
    ' Espaço não separável (U+00A0) -> espaço comum, pra não escapar das
    ' regras de colapso de espaço logo abaixo.
    resultado = Replace(resultado, ChrW(160), " ")

    Dim i As Long
    For i = 1 To 3 ' algumas passadas pra colapsar espaços múltiplos
        resultado = Replace(resultado, " R", "R")
        resultado = Replace(resultado, "R ", "R")
        resultado = Replace(resultado, " /", "/")
        resultado = Replace(resultado, "/ ", "/")
        resultado = Replace(resultado, " -", "-")
        resultado = Replace(resultado, "- ", "-")
    Next i

    ' Remove o "Z" de índices de velocidade embutidos no meio da medida
    ' (ZR, ZRF, Z sozinho) — "205/55ZR16" / "205/55ZRF16" / "205/55Z16"
    ' viram todos "205/55R16". Só troca quando está exatamente entre dois
    ' dígitos (perfil e aro), pra não mexer em "Z" de nome de marca/gama.
    On Error Resume Next
    Static regexZ As Object
    If regexZ Is Nothing Then
        Set regexZ = CreateObject("VBScript.RegExp")
        regexZ.Global = True
        regexZ.IgnoreCase = True
        regexZ.Pattern = "(\d)Z(?:RF|R)?(\d)"
    End If
    resultado = regexZ.Replace(resultado, "$1R$2")

    ' Remove o "F" de pneus Run Flat quando vem colado no R, SEM "Z" na
    ' frente — "235/50RF18" (RunFlat) vira "235/50R18". Mesma regra: só
    ' troca quando está exatamente entre dois dígitos.
    Static regexRF As Object
    If regexRF Is Nothing Then
        Set regexRF = CreateObject("VBScript.RegExp")
        regexRF.Global = True
        regexRF.IgnoreCase = True
        regexRF.Pattern = "(\d)RF(\d)"
    End If
    resultado = regexRF.Replace(resultado, "$1R$2")
    On Error GoTo 0

    NormalizarFormatacaoBasica = Trim(resultado)
End Function

' ==========================================================================
' Normaliza o texto para a busca de DIMENSÃO/GEOBOX, cobrindo variações de
' formatação vistas na macro legada:
'   - vírgula decimal -> ponto ("22,5" -> "22.5")
'   - "×" ou "*" no lugar de "X" ("31×10.50R15" / "31*10.5R15" -> "31X10.5R15")
'   - espaços ao redor de "R", "/" e "-" ("215/ 75R17.5" -> "215/75R17.5")
'   - padrão textual "NNN E ARO NN,N" -> "NNN/80RNN.N" (assume perfil 80,
'     igual fazia a macro legada nos dois casos hardcoded que ela tratava)
' NÃO sobrescreve a descrição original — usada só internamente na extração.
'
' PERFORMANCE: um único objeto RegExp cacheado com Static (reaproveitado
' entre chamadas, com o Pattern trocado antes de cada uso) no lugar de
' recriar 1 a 3 objetos COM a cada chamada — mesmo comportamento, chamada
' bem mais barata (roda uma vez por linha, em toda a base).
' ==========================================================================
Function NormalizarTextoDimensao(texto As String, Optional somenteMinBr As Boolean = False) As String
    Dim resultado As String
    resultado = NormalizarFormatacaoBasica(texto)

    On Error GoTo SemRegex
    Static regex As Object
    If regex Is Nothing Then
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = True
        regex.IgnoreCase = True
    End If

    ' Padrão "NNN E ARO NN,N" -> "NNN/80RNN.N" (assume perfil 80)
    regex.Pattern = "(\d{3})\s*E\s*ARO\s*(\d{2}(?:\.\d)?)"
    resultado = regex.Replace(resultado, "$1/80R$2")

    ' As duas regras de hífen abaixo assumem que "-" numa medida é sempre
    ' troca de digitação de "X" ou "/". Em bases BR/MIN existem GEOBOX com
    ' "-" que são válidos como estão (ex: "7.50-16", "9.00-20") — nelas essa
    ' suposição não vale, então as duas ficam desligadas nesse modo.
    If Not somenteMinBr Then
        ' Padrão "NN-NN.NNRNN" (hífen no lugar de "X", ex: "33-12.50R17") ->
        ' "NNXNN.NNRNN". Só aplica quando "R" vem logo depois do decimal, pra
        ' não confundir com o padrão "N.NN-NN" (ex: "9.00-20", onde o hífen faz
        ' o papel do "R" final, tratado depois pela troca hífen->R na gravação).
        regex.Pattern = "(\d{2,3})-(\d{1,2}\.\d{1,2})R"
        resultado = regex.Replace(resultado, "$1X$2R")

        ' Padrão "NNN-NNRNN" (hífen no lugar de "/", ex: "255-35R19", vindo de
        ' "255 - 35 R19" na descrição) -> "NNN/NNRNN". Só aplica quando o
        ' segundo número é inteiro (sem decimal — esse caso já foi tratado pela
        ' regra acima) e "R" vem logo em seguida — medida radial sempre usa "/"
        ' entre largura e perfil, então um hífen ali só pode ser troca de
        ' digitação. Não aplica em pneus diagonais tipo "9.00-20"/"7.50-16",
        ' que não têm "R" colado logo depois do segundo número.
        regex.Pattern = "(\d{2,3})-(\d{2,3})R"
        resultado = regex.Replace(resultado, "$1/$2R")
    End If

SemRegex:
    NormalizarTextoDimensao = resultado
End Function

' ==========================================================================
' Extrai o ARO (mesma lógica da macro ARO() original) — cheque formatos
' decimais (R17.5, R19.5, R22.5, R24.5) ANTES dos inteiros para não dar
' match parcial errado (ex: "R17" dentro de "R17.5").
' ==========================================================================
Function ExtrairAro(texto As String) As String

    Dim aros As Variant
    aros = Array("R17.5", "R19.5", "R22.5", "R24.5", _
                 "R12", "R13", "R14", "R15", "R16", "R17", "R18", "R19", _
                 "R20", "R21", "R22", "R23", "R24", "R26")

    Dim i As Long
    For i = LBound(aros) To UBound(aros)
        If InStr(1, texto, aros(i), vbTextCompare) > 0 Then
            ExtrairAro = aros(i)
            Exit Function
        End If
    Next i

    ExtrairAro = ""

End Function
