Attribute VB_Name = "modReferencia"
Option Explicit

' ==========================================================================
' MODULO: modReferencia
' Carrega a Tabela de Referencia externa (arquivo REF_FILE_PATH, ver
' modConfig) com as abas "Referencia", "MarcasExtras"/"MarcasExtra" e
' "ExcecoesMarca", e monta os dicionarios usados nas demais etapas.
' ==========================================================================

' ==========================================================================
' Carrega a Tabela de Referencia (arquivo externo) e monta:
'   dicSegPorGeobox      : chave GEOBOX -> SEGMENTO mais frequente entre os
'                          valores NAO VAZIOS daquele GEOBOX na aba
'                          "Referencia" (linhas com Segmento vazio nao
'                          entram na votacao)
'   dicLpPorGeobox       : chave GEOBOX -> LP. Fora de BR/MIN: LP mais
'                          frequente entre os valores NAO VAZIOS daquele
'                          GEOBOX (maioria, votacao independente da de
'                          Segmento). Em BR/MIN: LP da PRIMEIRA linha da
'                          Referencia com aquele GEOBOX, sem votacao.
'   dicLpPorGeoboxMarca  : chave GEOBOX & "|@|" & MARCA -> LP mais frequente
'                          (maioria) entre as linhas da Referencia que tem
'                          ESSA combinacao exata de GEOBOX e MARCA. So
'                          montado em modo BR/MIN -- usado como 1a tentativa
'                          de LP em modMain (mais especifico que so GEOBOX);
'                          se a combinacao nao existir, cai no dicLpPorGeobox
'                          normal (so GEOBOX).
'   arrMarcas()          : array de marcas unicas, ordenado da mais longa p/ mais curta
'   dicGeoboxPorMarca    : chave MARCA -> Collection de GEOBOX unicos daquela marca
'   dicGeoboxGlobalUnicos: chave GEOBOX -> True (todos os geobox unicos, p/ busca ampla)
'   dicGamasPorMarca     : chave MARCA -> Collection de GAMA unicas daquela marca
'   dicGamaGlobalUnicos  : chave GAMA -> True (todas as gamas unicas, p/ busca ampla)
'   dicMarcaPorGama      : chave GAMA -> MARCA dona daquela gama (1a encontrada na
'                          Referencia) -- usado pra "descobrir" a marca a partir da
'                          gama quando a marca nao foi lida na descricao
'   dicExcecoesGama      : chave PADRAO ABREVIADO -> GAMA correta (aba "ExcecoesGama",
'                          ex: "PTNZ" -> "POTENZA"), checada antes da busca normal
'
' somenteMinBr: quando True, restringe SO os dicionarios de GEOBOX (dicSegPorGeobox,
' dicLpPorGeobox, dicGeoboxPorMarca, dicGeoboxGlobalUnicos) as linhas da aba
' "Referencia" cuja coluna F ("Base de referencia") seja uma das 4 fontes
' STORM de Beyond Road/Mineracao (40117090, 40118090, 40119090, 40129090),
' "Dicionario WW" ou "Input manual" (GEOBOX classificados a mao pelo
' usuario) -- essas LPs nao tem intersecao de GEOBOX entre si, entao
' restringir evita ambiguidade. MARCA e GAMA (arrMarcas, dicGamasPorMarca,
' dicGamaGlobalUnicos, dicMarcaPorGama) continuam sendo montados a partir da
' Referencia INTEIRA, sem esse filtro -- uma marca/gama pode estar cadastrada
' fora dessas fontes e mesmo assim ser a marca/gama certa do produto.
' ==========================================================================
Function CarregarTabelaReferencia(ByRef dicSegPorGeobox As Object, ByRef dicLpPorGeobox As Object, _
                                   ByRef dicLpPorGeoboxMarca As Object, _
                                   ByRef arrMarcas() As String, ByRef dicGeoboxPorMarca As Object, _
                                   ByRef dicGeoboxGlobalUnicos As Object, ByRef dicExcecoesMarca As Object, _
                                   ByRef dicGamasPorMarca As Object, ByRef dicGamaGlobalUnicos As Object, _
                                   ByRef dicMarcaPorGama As Object, ByRef dicExcecoesGama As Object, _
                                   ByVal somenteMinBr As Boolean, ByRef diagnostico As String) As Boolean

    On Error GoTo ErroAbrir

    Dim wbRef As Workbook
    Dim wasOpen As Boolean
    Dim wbName As String
    wbName = Mid(REF_FILE_PATH, InStrRev(REF_FILE_PATH, "\") + 1)

    ' Reaproveita se ja estiver aberto, senao abre
    On Error Resume Next
    Set wbRef = Workbooks(wbName)
    On Error GoTo ErroAbrir
    If wbRef Is Nothing Then
        Set wbRef = Workbooks.Open(REF_FILE_PATH, ReadOnly:=True)
        wasOpen = False
    Else
        wasOpen = True
    End If

    diagnostico = "Arquivo usado: " & wbRef.FullName & vbCrLf
    diagnostico = diagnostico & "Ja estava aberto antes de rodar a macro? " & IIf(wasOpen, "SIM (reaproveitou a instancia ja aberta)", "Nao (abriu agora)") & vbCrLf
    diagnostico = diagnostico & "Abas encontradas no arquivo: "
    Dim wsDiag As Worksheet
    For Each wsDiag In wbRef.Sheets
        diagnostico = diagnostico & """" & wsDiag.Name & """; "
    Next wsDiag
    diagnostico = diagnostico & vbCrLf

    Dim wsRef As Worksheet
    Set wsRef = wbRef.Sheets("Referencia")

    Set dicGeoboxPorMarca = CreateObject("Scripting.Dictionary")
    Set dicGeoboxGlobalUnicos = CreateObject("Scripting.Dictionary")
    Set dicGamasPorMarca = CreateObject("Scripting.Dictionary")
    Set dicGamaGlobalUnicos = CreateObject("Scripting.Dictionary")
    Set dicMarcaPorGama = CreateObject("Scripting.Dictionary")

    ' dicVotosSegPorGeo / dicVotosLpPorGeo: chave GEOBOX -> Dictionary(valor -> contagem).
    ' Votacoes independentes: uma conta SEGMENTO, a outra conta LP, cada
    ' uma ignorando linhas onde o proprio valor esta vazio.
    Dim dicVotosSegPorGeo As Object, dicVotosLpPorGeo As Object
    Set dicVotosSegPorGeo = CreateObject("Scripting.Dictionary")
    Set dicVotosLpPorGeo = CreateObject("Scripting.Dictionary")

    ' GEOBOX -> LP da PRIMEIRA linha encontrada com esse GEOBOX. So usado em
    ' modo BR/MIN (ver mais abaixo) -- substitui a votacao por maioria, que
    ' nesse modo estava juntando LP de linhas que nao deveriam contar.
    Dim dicLpPrimeiraOcorrenciaMinBr As Object
    Set dicLpPrimeiraOcorrenciaMinBr = CreateObject("Scripting.Dictionary")

    ' chave GEOBOX & "|@|" & MARCA -> Dictionary(LP -> contagem). So usado em
    ' modo BR/MIN -- vira dicLpPorGeoboxMarca (maioria) no fim da funcao.
    Dim dicVotosLpPorGeoMarca As Object
    Set dicVotosLpPorGeoMarca = CreateObject("Scripting.Dictionary")

    Dim dicMarcasUnicas As Object
    Set dicMarcasUnicas = CreateObject("Scripting.Dictionary")

    ' dicGeoboxSetPorMarca / dicGamaSetPorMarca: mesma informacao de
    ' dicGeoboxPorMarca/dicGamasPorMarca, mas como Dictionary (chave = valor
    ' ja visto) em vez de Collection. Usados so pra checar duplicata em O(1)
    ' ao montar as Collections abaixo -- varrer a Collection inteira (O(n))
    ' pra cada linha da Referencia vira O(n2) numa marca com muitos GEOBOX/
    ' GAMA unicos (caso comum em bases BR/MIN, com centenas por marca).
    Dim dicGeoboxSetPorMarca As Object, dicGamaSetPorMarca As Object
    Set dicGeoboxSetPorMarca = CreateObject("Scripting.Dictionary")
    Set dicGamaSetPorMarca = CreateObject("Scripting.Dictionary")

    ' --- Modo BR/MIN: filtra pela coluna F ("Base de referencia") em vez  ---
    ' --- de LP. Uma mesma LP (MIN/BR) podia estar espalhada em outras     ---
    ' --- fontes/linhas fora do escopo de Beyond Road -- filtrar direto     ---
    ' --- pela fonte de origem e mais preciso.                             ---
    Dim dicFontesMinBr As Object
    Set dicFontesMinBr = CreateObject("Scripting.Dictionary")
    dicFontesMinBr.Add "STORM 40117090", True
    dicFontesMinBr.Add "STORM 40118090", True
    dicFontesMinBr.Add "STORM 40119090", True
    dicFontesMinBr.Add "STORM 40129090", True
    dicFontesMinBr.Add UCase("Dicionario WW"), True ' baseRef e comparado em UCase mais abaixo
    dicFontesMinBr.Add UCase("Input manual"), True ' geobox classificados manualmente pelo usuario

    ' Usa a maior "ultima linha com dado" entre as 6 colunas (Geobox, Marca,
    ' Gama, LP, Segmento, Base de referencia) em vez de so a coluna A. Uma
    ' linha so com Marca preenchida (sem Geobox) precisa ser lida do mesmo
    ' jeito -- se olhassemos so a coluna A, o laco pararia antes de chegar nela.
    Dim lastRowRef As Long, i As Long
    lastRowRef = 1
    Dim colRef As Long, ultimaLinhaColRef As Long
    For colRef = 1 To 6
        ultimaLinhaColRef = wsRef.Cells(wsRef.Rows.Count, colRef).End(xlUp).Row
        If ultimaLinhaColRef > lastRowRef Then lastRowRef = ultimaLinhaColRef
    Next colRef

    Dim geo As String, gama As String, marcaRef As String, lp As String, seg As String, baseRef As String
    Dim qtdLinhasIgnoradasMinBr As Long
    qtdLinhasIgnoradasMinBr = 0

    ' --- Le as 6 colunas de uma vez so num array em memoria. Evita milhares de ---
    ' --- chamadas COM individuais (wsRef.Cells) numa base grande -- cada uma e ---
    ' --- uma exposicao a queda de conexao (ex: sync do OneDrive/SharePoint no ---
    ' --- meio da leitura, causando erro -2147417848 "_Default do Range falhou"). ---
    Dim arrRef As Variant
    arrRef = wsRef.Range(wsRef.Cells(2, 1), wsRef.Cells(lastRowRef, 6)).Value

    For i = 2 To lastRowRef
        geo = UCase(Trim(CStr(arrRef(i - 1, 1))))             ' GEOBOX
        geo = NormalizarFormatacaoBasica(geo)                 ' mesma normalizacao usada na busca
        marcaRef = UCase(Trim(CStr(arrRef(i - 1, 2))))        ' MARCA
        ' Placeholders de "marca desconhecida" na Referencia (erro de       ---
        ' digitacao/preenchimento provisorio) -- tratados como celula vazia, ---
        ' senao viram "marca" valida e batem em quase qualquer descricao    ---
        ' (ex: "-" acha hifen em qualquer texto, "0" acha o digito 0).      ---
        If marcaRef = "-" Or marcaRef = "0" Or marcaRef = "0000" Or marcaRef = "N/A" Or marcaRef = "NA" Then
            marcaRef = ""
        End If
        gama = UCase(Trim(CStr(arrRef(i - 1, 3))))            ' GAMA (nao usada na busca de segmento/LP)
        lp = UCase(Trim(CStr(arrRef(i - 1, 4))))              ' LP
        seg = UCase(Trim(CStr(arrRef(i - 1, 5))))             ' SEGMENTO
        baseRef = UCase(Trim(CStr(arrRef(i - 1, 6))))         ' BASE DE REFERENCIA

        ' --- Base BR/MIN: so restringe o que e ESPECIFICO DE GEOBOX (essas   ---
        ' --- LPs nao tem intersecao de geobox entre si, entao limitar as    ---
        ' --- 4 fontes STORM/WW/Input manual evita ambiguidade ali). MARCA e ---
        ' --- GAMA continuam sendo buscadas na Referencia INTEIRA, mesmo em  ---
        ' --- modo BR/MIN -- uma marca/gama pode estar cadastrada so fora     ---
        ' --- dessas fontes e ainda assim ser a marca/gama certa do produto. ---
        Dim geoBrOk As Boolean
        geoBrOk = (Not somenteMinBr) Or dicFontesMinBr.Exists(baseRef)
        If somenteMinBr And Not geoBrOk Then qtdLinhasIgnoradasMinBr = qtdLinhasIgnoradasMinBr + 1

        ' --- Vota SEGMENTO e LP separadamente para esse GEOBOX, cada um so ---
        ' --- quando o proprio valor nao esta vazio (linha em branco numa  ---
        ' --- delas nao conta como voto nem "suja" a outra votacao).       ---
        Dim subVotos As Object

        If geoBrOk Then
            If Len(geo) > 0 And Len(seg) > 0 Then
                If Not dicVotosSegPorGeo.Exists(geo) Then
                    dicVotosSegPorGeo.Add geo, CreateObject("Scripting.Dictionary")
                End If
                Set subVotos = dicVotosSegPorGeo(geo)
                If subVotos.Exists(seg) Then
                    subVotos(seg) = subVotos(seg) + 1
                Else
                    subVotos.Add seg, 1
                End If
            End If

            If Len(geo) > 0 And Len(lp) > 0 Then
                If somenteMinBr Then
                    ' --- Em BR/MIN, LP nao e por maioria de votos -- fica com a  ---
                    ' --- PRIMEIRA linha da Referencia que tiver esse GEOBOX     ---
                    ' --- (ordem da propria planilha), sem contar ocorrencias.   ---
                    If Not dicLpPrimeiraOcorrenciaMinBr.Exists(geo) Then
                        dicLpPrimeiraOcorrenciaMinBr.Add geo, lp
                    End If
                Else
                    If Not dicVotosLpPorGeo.Exists(geo) Then
                        dicVotosLpPorGeo.Add geo, CreateObject("Scripting.Dictionary")
                    End If
                    Set subVotos = dicVotosLpPorGeo(geo)
                    If subVotos.Exists(lp) Then
                        subVotos(lp) = subVotos(lp) + 1
                    Else
                        subVotos.Add lp, 1
                    End If
                End If

                ' --- Votacao por GEOBOX + MARCA combinados (so BR/MIN) --      ---
                ' --- mais especifica que a votacao so por GEOBOX acima; usada ---
                ' --- em modMain como 1a tentativa de LP, antes de cair na de  ---
                ' --- so GEOBOX.                                               ---
                If somenteMinBr And Len(marcaRef) > 0 Then
                    Dim chaveGeoMarca As String
                    chaveGeoMarca = geo & "|@|" & marcaRef
                    If Not dicVotosLpPorGeoMarca.Exists(chaveGeoMarca) Then
                        dicVotosLpPorGeoMarca.Add chaveGeoMarca, CreateObject("Scripting.Dictionary")
                    End If
                    Set subVotos = dicVotosLpPorGeoMarca(chaveGeoMarca)
                    If subVotos.Exists(lp) Then
                        subVotos(lp) = subVotos(lp) + 1
                    Else
                        subVotos.Add lp, 1
                    End If
                End If
            End If
        End If

        If Len(marcaRef) > 0 Then
            If Not dicMarcasUnicas.Exists(marcaRef) Then dicMarcasUnicas.Add marcaRef, True

            If geoBrOk Then
                If Not dicGeoboxPorMarca.Exists(marcaRef) Then
                    dicGeoboxPorMarca.Add marcaRef, New Collection
                    dicGeoboxSetPorMarca.Add marcaRef, CreateObject("Scripting.Dictionary")
                End If
                If Len(geo) > 0 Then
                    If Not dicGeoboxSetPorMarca(marcaRef).Exists(geo) Then
                        dicGeoboxSetPorMarca(marcaRef).Add geo, True
                        dicGeoboxPorMarca(marcaRef).Add geo
                    End If
                End If
            End If

            If Not dicGamasPorMarca.Exists(marcaRef) Then
                dicGamasPorMarca.Add marcaRef, New Collection
                dicGamaSetPorMarca.Add marcaRef, CreateObject("Scripting.Dictionary")
            End If
            If Len(gama) > 0 Then
                If Not dicGamaSetPorMarca(marcaRef).Exists(gama) Then
                    dicGamaSetPorMarca(marcaRef).Add gama, True
                    dicGamasPorMarca(marcaRef).Add gama
                End If
            End If
        End If

        If geoBrOk Then
            If Len(geo) > 0 Then
                If Not dicGeoboxGlobalUnicos.Exists(geo) Then dicGeoboxGlobalUnicos.Add geo, True
            End If
        End If

        If Len(gama) > 0 Then
            If Not dicGamaGlobalUnicos.Exists(gama) Then dicGamaGlobalUnicos.Add gama, True

            ' Gama e tratada como praticamente exclusiva de uma marca -- a
            ' primeira marca encontrada pra essa gama na Referencia "ganha"
            ' o dicionario reverso (ignora linhas com marca vazia).
            If Len(marcaRef) > 0 And Not dicMarcaPorGama.Exists(gama) Then
                dicMarcaPorGama.Add gama, marcaRef
            End If
        End If
    Next i

    ' --- Monta dicSegPorGeobox: fica com o valor NAO VAZIO mais votado       ---
    ' --- (maioria); em empate, o primeiro encontrado (ordem da Referencia). ---
    Set dicSegPorGeobox = MontarDicMaioriaPorGeobox(dicVotosSegPorGeo)

    ' --- dicLpPorGeobox: em modo BR/MIN, e a PRIMEIRA linha da Referencia   ---
    ' --- com aquele GEOBOX (dicLpPrimeiraOcorrenciaMinBr) -- nao e mais      ---
    ' --- maioria de votos, pra nao misturar LP de linhas que nao deveriam.  ---
    ' --- Fora de BR/MIN, continua sendo maioria, como sempre foi.          ---
    If somenteMinBr Then
        Set dicLpPorGeobox = dicLpPrimeiraOcorrenciaMinBr
        Set dicLpPorGeoboxMarca = MontarDicMaioriaPorGeobox(dicVotosLpPorGeoMarca)
    Else
        Set dicLpPorGeobox = MontarDicMaioriaPorGeobox(dicVotosLpPorGeo)
        Set dicLpPorGeoboxMarca = CreateObject("Scripting.Dictionary")
    End If

    diagnostico = diagnostico & "Ultima linha lida na aba ""Referencia"": " & lastRowRef & vbCrLf
    diagnostico = diagnostico & "Modo BR/MIN ativado? " & IIf(somenteMinBr, "SIM (GEOBOX restrito a ""Base de referencia"" = STORM 40117090/40118090/40119090/40129090, Dicionario WW ou Input manual; MARCA e GAMA continuam buscando na Referencia inteira)", "Nao") & vbCrLf
    If somenteMinBr Then
        diagnostico = diagnostico & "Linhas da Referencia fora dessas fontes (nao contam para GEOBOX, mas contam para MARCA/GAMA): " & qtdLinhasIgnoradasMinBr & vbCrLf
    End If
    diagnostico = diagnostico & "Total de GEOBOX unicos com SEGMENTO mapeado: " & dicSegPorGeobox.Count & vbCrLf
    diagnostico = diagnostico & "Total de GEOBOX unicos com LP mapeado: " & dicLpPorGeobox.Count & vbCrLf
    If wasOpen Then
        diagnostico = diagnostico & vbCrLf & "ATENCAO: o arquivo ja estava aberto e foi reaproveitado. " & _
                      "Se os numeros acima parecerem baixos demais, FECHE o arquivo Tabela_Referencia.xlsx " & _
                      "no Excel (sem salvar) e rode a macro de novo -- ele sera reaberto do zero, com a versao mais recente salva." & vbCrLf
    End If

    ' --- Soma marcas extras da aba "MarcasExtras" (aceita variacao sem "s"
    ' final: "MarcasExtra") ---
    Dim wsMarcasExtras As Worksheet
    Set wsMarcasExtras = Nothing
    On Error Resume Next
    Set wsMarcasExtras = wbRef.Sheets("MarcasExtras")
    If wsMarcasExtras Is Nothing Then Set wsMarcasExtras = wbRef.Sheets("MarcasExtra")
    On Error GoTo ErroAbrir

    Dim qtdMarcasExtrasLidas As Long
    qtdMarcasExtrasLidas = 0

    If Not wsMarcasExtras Is Nothing Then
        Dim lastRowExtras As Long, j As Long, mExtra As String
        lastRowExtras = wsMarcasExtras.Cells(wsMarcasExtras.Rows.Count, "A").End(xlUp).Row
        Dim arrExtras As Variant
        If lastRowExtras >= 2 Then arrExtras = wsMarcasExtras.Range(wsMarcasExtras.Cells(2, 1), wsMarcasExtras.Cells(lastRowExtras, 1)).Value
        For j = 2 To lastRowExtras
            If lastRowExtras = 2 Then
                mExtra = UCase(Trim(CStr(arrExtras(1, 1))))
            Else
                mExtra = UCase(Trim(CStr(arrExtras(j - 1, 1))))
            End If
            If Len(mExtra) > 0 Then
                If Not dicMarcasUnicas.Exists(mExtra) Then dicMarcasUnicas.Add mExtra, True
                qtdMarcasExtrasLidas = qtdMarcasExtrasLidas + 1
            End If
        Next j
        diagnostico = diagnostico & "Aba de marcas extras encontrada (""" & wsMarcasExtras.Name & """). Ultima linha: " & lastRowExtras & _
                      ". Marcas lidas (linha 2 ate " & lastRowExtras & "): " & qtdMarcasExtrasLidas & vbCrLf
    Else
        diagnostico = diagnostico & "Nenhuma aba ""MarcasExtras"" ou ""MarcasExtra"" encontrada nesse arquivo." & vbCrLf
    End If

    ' --- Carrega excecoes de marca da aba "ExcecoesMarca" (se existir) ---
    Set dicExcecoesMarca = CreateObject("Scripting.Dictionary")
    Dim wsExcecoes As Worksheet
    Set wsExcecoes = Nothing
    On Error Resume Next
    Set wsExcecoes = wbRef.Sheets("ExcecoesMarca")
    On Error GoTo ErroAbrir

    Dim qtdExcecoesLidas As Long
    qtdExcecoesLidas = 0

    If Not wsExcecoes Is Nothing Then
        Dim lastRowExc As Long, padrao As String, marcaCorreta As String
        lastRowExc = wsExcecoes.Cells(wsExcecoes.Rows.Count, "A").End(xlUp).Row
        Dim arrExcMarca As Variant
        If lastRowExc >= 2 Then arrExcMarca = wsExcecoes.Range(wsExcecoes.Cells(2, 1), wsExcecoes.Cells(lastRowExc, 2)).Value
        For j = 2 To lastRowExc
            If lastRowExc = 2 Then
                padrao = UCase(Trim(CStr(arrExcMarca(1, 1))))
                marcaCorreta = UCase(Trim(CStr(arrExcMarca(1, 2))))
            Else
                padrao = UCase(Trim(CStr(arrExcMarca(j - 1, 1))))
                marcaCorreta = UCase(Trim(CStr(arrExcMarca(j - 1, 2))))
            End If
            If Len(padrao) > 0 And Len(marcaCorreta) > 0 Then
                If Not dicExcecoesMarca.Exists(padrao) Then dicExcecoesMarca.Add padrao, marcaCorreta
                qtdExcecoesLidas = qtdExcecoesLidas + 1
            End If
        Next j
        diagnostico = diagnostico & "Aba ""ExcecoesMarca"" encontrada. Ultima linha: " & lastRowExc & _
                      ". Excecoes lidas (linha 2 ate " & lastRowExc & "): " & qtdExcecoesLidas & vbCrLf
    Else
        diagnostico = diagnostico & "Aba ""ExcecoesMarca"" NAO encontrada nesse arquivo." & vbCrLf
    End If

    ' --- Carrega excecoes/abreviacoes de GAMA da aba "ExcecoesGama" (se existir) ---
    ' Mesmo modelo da "ExcecoesMarca": coluna A = padrao abreviado como
    ' aparece na descricao (ex: "PTNZ"), coluna B = GAMA correta (ex:
    ' "POTENZA"). Checada antes da busca normal por GAMA em ExtrairGama.
    Set dicExcecoesGama = CreateObject("Scripting.Dictionary")
    Dim wsExcecoesGama As Worksheet
    Set wsExcecoesGama = Nothing
    On Error Resume Next
    Set wsExcecoesGama = wbRef.Sheets("ExcecoesGama")
    On Error GoTo ErroAbrir

    Dim qtdExcecoesGamaLidas As Long
    qtdExcecoesGamaLidas = 0

    If Not wsExcecoesGama Is Nothing Then
        Dim lastRowExcGama As Long, padraoGama As String, gamaCorreta As String
        lastRowExcGama = wsExcecoesGama.Cells(wsExcecoesGama.Rows.Count, "A").End(xlUp).Row
        Dim arrExcGama As Variant
        If lastRowExcGama >= 2 Then arrExcGama = wsExcecoesGama.Range(wsExcecoesGama.Cells(2, 1), wsExcecoesGama.Cells(lastRowExcGama, 2)).Value
        For j = 2 To lastRowExcGama
            If lastRowExcGama = 2 Then
                padraoGama = UCase(Trim(CStr(arrExcGama(1, 1))))
                gamaCorreta = UCase(Trim(CStr(arrExcGama(1, 2))))
            Else
                padraoGama = UCase(Trim(CStr(arrExcGama(j - 1, 1))))
                gamaCorreta = UCase(Trim(CStr(arrExcGama(j - 1, 2))))
            End If
            If Len(padraoGama) > 0 And Len(gamaCorreta) > 0 Then
                If Not dicExcecoesGama.Exists(padraoGama) Then dicExcecoesGama.Add padraoGama, gamaCorreta
                qtdExcecoesGamaLidas = qtdExcecoesGamaLidas + 1
            End If
        Next j
        diagnostico = diagnostico & "Aba ""ExcecoesGama"" encontrada. Ultima linha: " & lastRowExcGama & _
                      ". Excecoes lidas (linha 2 ate " & lastRowExcGama & "): " & qtdExcecoesGamaLidas & vbCrLf
    Else
        diagnostico = diagnostico & "Aba ""ExcecoesGama"" NAO encontrada nesse arquivo (abreviacoes de GAMA nao serao reconhecidas)." & vbCrLf
    End If

    diagnostico = diagnostico & "Total de marcas unicas na busca (Referencia + MarcasExtras): " & dicMarcasUnicas.Count

    ' --- Monta arrMarcas ordenado da mais longa para a mais curta ---
    ReDim arrMarcas(0 To dicMarcasUnicas.Count - 1)
    Dim k As Long
    k = 0
    Dim chaveMarca As Variant
    For Each chaveMarca In dicMarcasUnicas.Keys
        arrMarcas(k) = CStr(chaveMarca)
        k = k + 1
    Next chaveMarca

    Dim a As Long, b As Long, temp As String
    For a = LBound(arrMarcas) To UBound(arrMarcas) - 1
        For b = a + 1 To UBound(arrMarcas)
            If Len(arrMarcas(b)) > Len(arrMarcas(a)) Then
                temp = arrMarcas(a)
                arrMarcas(a) = arrMarcas(b)
                arrMarcas(b) = temp
            End If
        Next b
    Next a

    If Not wasOpen Then wbRef.Close SaveChanges:=False

    CarregarTabelaReferencia = True
    Exit Function

ErroAbrir:
    CarregarTabelaReferencia = False
End Function

' ==========================================================================
' Recebe um dicionario de votos (chave GEOBOX -> Dictionary(valor -> contagem),
' montado ignorando valores vazios) e devolve um dicionario simples
' chave GEOBOX -> valor mais votado (maioria). Em empate, fica com o
' primeiro valor encontrado (ordem de leitura da aba "Referencia").
' ==========================================================================
Function MontarDicMaioriaPorGeobox(dicVotosPorGeo As Object) As Object
    Dim dicResultado As Object
    Set dicResultado = CreateObject("Scripting.Dictionary")

    Dim geoKey As Variant, valorIter As Variant
    Dim subVotos As Object
    Dim melhorValor As String, melhorContagem As Long

    For Each geoKey In dicVotosPorGeo.Keys
        Set subVotos = dicVotosPorGeo(geoKey)
        melhorValor = ""
        melhorContagem = 0
        For Each valorIter In subVotos.Keys
            If subVotos(valorIter) > melhorContagem Then
                melhorContagem = subVotos(valorIter)
                melhorValor = CStr(valorIter)
            End If
        Next valorIter
        dicResultado.Add CStr(geoKey), melhorValor
    Next geoKey

    Set MontarDicMaioriaPorGeobox = dicResultado
End Function

' ==========================================================================
' Retorna, para uma DIMENSAO (GEOBOX) ja extraida, o SEGMENTO e a LP.
' Nao depende de MARCA nem de achar a GAMA no texto -- e busca direta por
' DIMENSAO. Regra de prioridade do LP:
'   1) Se achou SEGMENTO pra essa dimensao, LP = DeduzirLp(segmento)
'      (TC para PC/REC/COM, PL para TLD/PPL/BUS, BR para DM).
'   2) Se DeduzirLp nao souber mapear esse segmento (retornou ""), ou se
'      nao achou SEGMENTO nenhum, usa a votacao independente de LP
'      (dicLpPorGeobox) -- maioria dos valores de LP nao vazios daquele
'      GEOBOX na Tabela de Referencia.
' Se a DIMENSAO nao estiver cadastrada em nenhum dos dois dicionarios,
' retorna "" para os dois. DeduzirLp vem de modClassificacaoRegras.bas.
' ==========================================================================
Sub ObterSegmentoELpPorDimensao(dimensao As String, dicSegPorGeobox As Object, dicLpPorGeobox As Object, _
                                 ByRef segmentoOut As String, ByRef lpOut As String)
    segmentoOut = ""
    lpOut = ""

    If Len(dimensao) = 0 Then Exit Sub

    If dicSegPorGeobox.Exists(dimensao) Then segmentoOut = CStr(dicSegPorGeobox(dimensao))

    Dim lpVotado As String
    lpVotado = ""
    If dicLpPorGeobox.Exists(dimensao) Then lpVotado = CStr(dicLpPorGeobox(dimensao))

    If Len(segmentoOut) > 0 Then
        lpOut = DeduzirLp(segmentoOut)
        If Len(lpOut) = 0 Then lpOut = lpVotado
    Else
        lpOut = lpVotado
    End If
End Sub
