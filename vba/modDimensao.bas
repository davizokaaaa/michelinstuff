Attribute VB_Name = "modDimensao"
Option Explicit

' ==========================================================================
' MODULO: modDimensao
' Normalizacao de formatacao de texto de medida, extracao de DIMENSAO
' (GEOBOX) e de ARO a partir da descricao.
' ==========================================================================

' ==========================================================================
' Extrai a DIMENSAO (GEOBOX): 1) se a marca ja foi identificada, procura
' primeiro so entre as medidas conhecidas DAQUELA marca (mais rapido e mais
' preciso). 2) Se nao achar assim, faz busca ampla em todas as medidas
' conhecidas de todas as marcas. 3) Se ainda assim nao achar, retorna "".
' Em caso de mais de uma medida batendo no texto, fica com a mais longa
' (evita pegar "R17" quando na verdade e "R17.5"). O texto e normalizado
' antes (virgula->ponto, x/*->X, espacos ao redor de R/barra/traco) para
' bater com mais variacoes de formatacao. Tambem compara contra uma versao
' com "D" (construcao diagonal, ex: "135/80D15") trocado por "R", pois a
' Tabela de Referencia as vezes so tem a variante radial cadastrada -- nesse
' caso grava o valor com "R" mesmo (como esta na Referencia), nao o "D"
' original da descricao.
' ==========================================================================
' Tamanho minimo aceito pra um candidato de DIMENSAO ser considerado um
' match valido. Protege contra entradas curtas demais/corrompidas na
' Tabela de Referencia (ex: uma celula com so "R") baterem por acidente
' em qualquer trecho da descricao e virarem resultado sem sentido -- a
' menor medida real (tipo "9R20") ainda tem pelo menos 4 caracteres.
Const TAMANHO_MINIMO_DIMENSAO As Long = 4

' ==========================================================================
' DE-PARA CEGO de GEOBOX -- usado SO no modo BR/MIN (somenteMinBr = True).
' Validado num teste isolado (modTesteMatchExato) contra a base real antes
' de virar padrao aqui: match LITERAL do valor exatamente como cadastrado na
' Tabela de Referencia dentro da descricao, sem tentar reconhecer "formato"
' de medida (largura/perfil/aro) -- so uma limpeza puramente de formatacao,
' aplicada nos DOIS lados antes de comparar: virgula -> ponto decimal.
' (A regra de tratar ".50" e ".5" como iguais foi removida daqui -- estava
' gerando GEOBOX errado ao juntar medidas diferentes que so coincidem depois
' de arredondar o zero.) E o espaco e testado de duas formas (pode ser
' decorativo OU estar no lugar de uma barra que faltou): espaco removido, e
' espaco virando "/". Nunca inventa nem adivinha -- so bate se existir
' exatamente (depois dessa limpeza) no dicionario da Tabela de Referencia.
' ==========================================================================

Function ExtrairDimensao(descricao As String, marca As String, _
                          dicGeoboxPorMarca As Object, dicGeoboxGlobalUnicos As Object, _
                          dicPadroesLegado As Object, _
                          Optional somenteMinBr As Boolean = False) As String

    Dim textoNorm As String
    textoNorm = NormalizarTextoDimensao(descricao, somenteMinBr)

    ' --- Segunda versao do texto, com TODOS os espacos removidos. Cobre   ---
    ' --- casos como "245/35 ZR19" (espaco antes do Z, escapa da regra de  ---
    ' --- Z/ZR/ZRF que exige digito colado) ou ate numeros quebrados por   ---
    ' --- engano no meio ("22 5/55R17" -> "225/55R17" depois de juntar).   ---
    ' --- Reaplica a limpeza de Z/RF depois de juntar, pra pegar casos que ---
    ' --- so viram "digito colado com Z/RF" depois da juncao.             ---
    Dim textoSemEspaco As String
    textoSemEspaco = Replace(textoNorm, " ", "")
    textoSemEspaco = NormalizarFormatacaoBasica(textoSemEspaco)

    ' --- Medidas escritas por extenso (laudos INMETRO, "BANDA/SERIE/ARO" ---
    ' --- em varios formatos) -- logica isolada em modDimensaoExtenso pra   ---
    ' --- poder evoluir/reverter sem mexer na extracao ja validada. So    ---
    ' --- ACRESCENTA o candidato montado ao texto de busca; a validacao   ---
    ' --- (existe no catalogo? tamanho minimo?) continua sendo feita nos  ---
    ' --- mesmos lacos abaixo, como qualquer outro trecho da descricao.   ---
    Dim candidatoExtenso As String
    candidatoExtenso = ExtrairDimensaoPorExtenso(descricao, somenteMinBr)
    If Len(candidatoExtenso) > 0 Then
        textoNorm = textoNorm & " " & candidatoExtenso
        textoSemEspaco = textoSemEspaco & Replace(candidatoExtenso, " ", "")
    End If

    ' --- Versoes com "D" (construcao diagonal) trocado por "R", so para   ---
    ' --- efeito de COMPARACAO com o catalogo. Existem aros com construcao ---
    ' --- diagonal na descricao (ex: "135/80D15") cujo cadastro na Tabela  ---
    ' --- de Referencia so tem a variante radial ("135/80R15") -- sem isso, ---
    ' --- esses aros nunca seriam encontrados. Nao reaplica em nenhum      ---
    ' --- outro lugar do codigo (marca, gama etc.), so nesta comparacao.   ---
    Dim textoNormDparaR As String, textoSemEspacoDparaR As String
    textoNormDparaR = Replace(textoNorm, "D", "R")
    textoSemEspacoDparaR = Replace(textoSemEspaco, "D", "R")

    Dim melhor As String, melhorLen As Long
    melhor = ""
    melhorLen = 0

    ' --- Passo 0: DE-PARA cego -- SO no modo BR/MIN --- Fora desse modo, a
    ' --- extracao continua 100% igual a que ja funciona (busca literal por ---
    ' --- substring abaixo), sem nenhuma mudanca de comportamento.          ---
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
    ' IMPORTANTE: as duas buscas SEMPRE rodam, contra as DUAS versoes do
    ' texto (normal e sem espaco), e fica com a mais longa entre todas --
    ' nunca aceita cegamente o que a busca por marca achar primeiro.
    ' Isso evita que uma medida catalogada pra aquela marca, mas SEM relacao
    ' com o produto da linha, bata por coincidencia em algum trecho solto da
    ' descricao (numeros de registro, certificado, codigo de familia etc.)
    ' e "roube" a vaga de uma medida mais longa e mais correta que so
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

    ' --- Busca ampla em qualquer medida conhecida (Tabela de Referencia) ---
    ' --- So fora de BR/MIN: em BR/MIN, o Passo 0 (DE-PARA cego, acima) ja    ---
    ' --- varreu esse MESMO dicGeoboxGlobalUnicos inteiro pra achar o melhor ---
    ' --- match; repetir a varredura aqui (com 4 InStr por item) e trabalho  ---
    ' --- em dobro por linha -- e com milhares de GEOBOX unicos numa base     ---
    ' --- BR/MIN, isso e o que deixava a macro lentissima/travando.          ---
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

    ' --- Passo 3: lista especifica vinda da macro legada (padrao -> valor oficial) ---
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
' Normalizacao BASICA de formatacao (compartilhada): maiusculas, virgula
' decimal -> ponto, x/* -> X, colapsa espacos soltos ao redor de R/barra/
' traco. Usada tanto no texto da descricao quanto nos valores de GEOBOX
' lidos da propria Tabela de Referencia -- os DOIS lados da comparacao
' precisam estar no mesmo formato, senao "175/75 R13" (com espaco, na
' tabela) nunca bate com "175/75R13" (sem espaco, ja limpo na descricao).
'
' PERFORMANCE: os dois objetos RegExp usados aqui sao cacheados com Static
' (criados uma unica vez por sessao do Excel, nao a cada chamada). Essa
' funcao roda varias vezes por linha, em toda a base -- recriar o objeto
' COM do RegExp a cada chamada (como era antes) e um dos maiores custos
' de desempenho do laco principal. O padrao/comportamento do regex nao
' muda: so o objeto passa a ser reaproveitado.
' ==========================================================================
Function NormalizarFormatacaoBasica(texto As String) As String
    Dim resultado As String
    resultado = UCase(texto)

    resultado = Replace(resultado, ",", ".")
    resultado = Replace(resultado, ChrW(215), "X") ' x (sinal de multiplicacao, U+00D7)
    resultado = Replace(resultado, Chr(215), "X")   ' fallback caso venha como ANSI
    resultado = Replace(resultado, "*", "X")

    ' Tracos "parecidos" com hifen, mas que sao caracteres Unicode diferentes
    ' (comuns em texto colado de Excel/Word/PDF) -- todos viram o hifen comum
    ' "-" (U+002D). Sem isso, um GEOBOX como "33X12-20/7.50" registrado com
    ' um desses tracos nunca bate com a mesma medida digitada com hifen
    ' normal na descricao (ou vice-versa), mesmo sendo visualmente identicos.
    resultado = Replace(resultado, ChrW(8211), "-") ' en dash (U+2013)
    resultado = Replace(resultado, ChrW(8212), "-") ' em dash (U+2014)
    resultado = Replace(resultado, ChrW(8722), "-") ' sinal de menos matematico (U+2212)
    resultado = Replace(resultado, ChrW(8209), "-") ' hifen nao separavel (U+2011)
    ' Espaco nao separavel (U+00A0) -> espaco comum, pra nao escapar das
    ' regras de colapso de espaco logo abaixo.
    resultado = Replace(resultado, ChrW(160), " ")

    Dim i As Long
    For i = 1 To 3 ' algumas passadas pra colapsar espacos multiplos
        resultado = Replace(resultado, " R", "R")
        resultado = Replace(resultado, "R ", "R")
        resultado = Replace(resultado, " /", "/")
        resultado = Replace(resultado, "/ ", "/")
        resultado = Replace(resultado, " -", "-")
        resultado = Replace(resultado, "- ", "-")
    Next i

    ' Remove o "Z" de indices de velocidade embutidos no meio da medida
    ' (ZR, ZRF, Z sozinho) -- "205/55ZR16" / "205/55ZRF16" / "205/55Z16"
    ' viram todos "205/55R16". So troca quando esta exatamente entre dois
    ' digitos (perfil e aro), pra nao mexer em "Z" de nome de marca/gama.
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
    ' frente -- "235/50RF18" (RunFlat) vira "235/50R18". Mesma regra: so
    ' troca quando esta exatamente entre dois digitos.
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
' Normaliza o texto para a busca de DIMENSAO/GEOBOX, cobrindo variacoes de
' formatacao vistas na macro legada:
'   - virgula decimal -> ponto ("22,5" -> "22.5")
'   - "x" ou "*" no lugar de "X" ("31x10.50R15" / "31*10.5R15" -> "31X10.5R15")
'   - espacos ao redor de "R", "/" e "-" ("215/ 75R17.5" -> "215/75R17.5")
'   - padrao textual "NNN E ARO NN,N" -> "NNN/80RNN.N" (assume perfil 80,
'     igual fazia a macro legada nos dois casos hardcoded que ela tratava)
' NAO sobrescreve a descricao original -- usada so internamente na extracao.
'
' PERFORMANCE: um unico objeto RegExp cacheado com Static (reaproveitado
' entre chamadas, com o Pattern trocado antes de cada uso) no lugar de
' recriar 1 a 3 objetos COM a cada chamada -- mesmo comportamento, chamada
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

    ' Padrao "NNN E ARO NN,N" -> "NNN/80RNN.N" (assume perfil 80)
    regex.Pattern = "(\d{3})\s*E\s*ARO\s*(\d{2}(?:\.\d)?)"
    resultado = regex.Replace(resultado, "$1/80R$2")

    ' As duas regras de hifen abaixo assumem que "-" numa medida e sempre
    ' troca de digitacao de "X" ou "/". Em bases BR/MIN existem GEOBOX com
    ' "-" que sao validos como estao (ex: "7.50-16", "9.00-20") -- nelas essa
    ' suposicao nao vale, entao as duas ficam desligadas nesse modo.
    If Not somenteMinBr Then
        ' Padrao "NN-NN.NNRNN" (hifen no lugar de "X", ex: "33-12.50R17") ->
        ' "NNXNN.NNRNN". So aplica quando "R" vem logo depois do decimal, pra
        ' nao confundir com o padrao "N.NN-NN" (ex: "9.00-20", onde o hifen faz
        ' o papel do "R" final, tratado depois pela troca hifen->R na gravacao).
        regex.Pattern = "(\d{2,3})-(\d{1,2}\.\d{1,2})R"
        resultado = regex.Replace(resultado, "$1X$2R")

        ' Padrao "NNN-NNRNN" (hifen no lugar de "/", ex: "255-35R19", vindo de
        ' "255 - 35 R19" na descricao) -> "NNN/NNRNN". So aplica quando o
        ' segundo numero e inteiro (sem decimal -- esse caso ja foi tratado pela
        ' regra acima) e "R" vem logo em seguida -- medida radial sempre usa "/"
        ' entre largura e perfil, entao um hifen ali so pode ser troca de
        ' digitacao. Nao aplica em pneus diagonais tipo "9.00-20"/"7.50-16",
        ' que nao tem "R" colado logo depois do segundo numero.
        regex.Pattern = "(\d{2,3})-(\d{2,3})R"
        resultado = regex.Replace(resultado, "$1/$2R")
    End If

SemRegex:
    NormalizarTextoDimensao = resultado
End Function

' ==========================================================================
' Extrai o ARO (mesma logica da macro ARO() original) -- cheque formatos
' decimais (R17.5, R19.5, R22.5, R24.5) ANTES dos inteiros para nao dar
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
