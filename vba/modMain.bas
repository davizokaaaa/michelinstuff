Attribute VB_Name = "modMain"
Option Explicit

' ==========================================================================
' MODULO: modMain
' Sub principal ClassificarTudo — orquestra a leitura das colunas de
' entrada, o carregamento da Tabela de Referência (modReferencia) e a
' aplicação das extrações/regras (modMarca, modDimensao,
' modClassificacaoRegras, modDicionarios, modEstepeCompeticao) linha a
' linha, gravando nas colunas de saída.
'
' Depende dos demais módulos deste projeto:
'   modConfig, modPlanilha, modDicionarios, modReferencia,
'   modMarca, modDimensao, modClassificacaoRegras, modEstepeCompeticao
'
' Colunas de ENTRADA (precisam já existir na planilha, vêm de fora):
'   - DESCRIÇÃODA MERCADORIA
'   - PROVÁVEL ADQUIRENTE (nome configurável em modConfig.COL_ADQUIRENTE_NOME)
'
' Colunas de SAÍDA (a macro cria sozinha se não existirem, nesta ordem
' quando precisar criar mais de uma): LP, TIPO PRODUTO, MARCA, GAMA, ANIP,
' RT/OE, DIMENSÃO (ou GEOBOX), ARO, ESTEPE/COMPETICAO
'
' Ver regras de extração detalhadas nos comentários de cada módulo
' correspondente (modMarca, modDimensao, modClassificacaoRegras,
' modReferencia, modEstepeCompeticao).
'
' PREMISSAS A CONFIRMAR:
'   1) RT/OE usa a coluna "PROVÁVEL ADQUIRENTE" como nome do comprador.
'      Se for outra coluna, troque modConfig.COL_ADQUIRENTE_NOME.
'   2) O arquivo de referência se chama "Tabela_Referencia.xlsx" e fica no
'      caminho definido em modConfig.REF_FILE_PATH. Dentro dele:
'        - aba "Referencia" (obrigatória): colunas GEOBOX, GAMA, MARCA, LP, SEGMENTO
'        - aba "MarcasExtras" (opcional): coluna MARCA — marcas extras somadas
'          à busca, sem precisar mexer no código
'        - aba "ExcecoesMarca" (opcional): colunas PADRAO_ENCONTRADO_NA_DESCRICAO
'          e MARCA_CORRETA — checadas antes de tudo na extração de MARCA
'        - aba "ExcecoesGama" (opcional): mesmo modelo da "ExcecoesMarca",
'          mas pra abreviações de GAMA (ex: "PTNZ" -> "POTENZA")
'   3) ESTEPE/COMPETIÇÃO (modEstepeCompeticao) foi portado da macro legada
'      "Sub estepe()", que lia a coluna H — aqui reusa a mesma
'      DESCRIÇÃODA MERCADORIA já lida pra MARCA/GAMA/DIMENSÃO. Confirme que
'      isso é equivalente à coluna H da planilha onde a legada rodava.
'
' PERFORMANCE (ver notas de cada bloco abaixo): esta versão substitui a
' leitura/escrita célula a célula (ws.Cells(i, col)) por leitura em bloco
' num array no início e escrita em bloco no final, e acrescenta um cache
' de resultado por (DESCRIÇÃO, ADQUIRENTE) — nenhuma das duas mudanças
' altera o valor calculado em nenhuma linha, só a forma de ler/escrever/
' evitar recálculo do que já foi calculado antes.
' ==========================================================================

Sub ClassificarTudo()

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim colDescricao As Long, colAdquirente As Long
    Dim colRtOe As Long, colTipoProduto As Long, colAro As Long, colAnip As Long
    Dim colLp As Long, colMarca As Long, colDimensao As Long, colGama As Long
    Dim colEstepeComp As Long

    Set ws = ActiveSheet

    ' --- Colunas de ENTRADA: precisam existir (a macro não pode inventar) ---
    colDescricao = LocalizarColuna(ws, "DESCRIÇÃODA MERCADORIA")
    colAdquirente = LocalizarColuna(ws, COL_ADQUIRENTE_NOME)

    If colDescricao = 0 Or colAdquirente = 0 Then

        Dim msgFaltantes As String
        msgFaltantes = "Não encontrei a(s) seguinte(s) coluna(s) de ENTRADA no cabeçalho (linha 1):" & vbCrLf & vbCrLf
        If colDescricao = 0 Then msgFaltantes = msgFaltantes & "- DESCRIÇÃODA MERCADORIA" & vbCrLf
        If colAdquirente = 0 Then msgFaltantes = msgFaltantes & "- " & COL_ADQUIRENTE_NOME & vbCrLf

        msgFaltantes = msgFaltantes & vbCrLf & "Cabeçalhos encontrados na linha 1:" & vbCrLf
        Dim lastColDiag As Long, cDiag As Long
        lastColDiag = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        For cDiag = 1 To lastColDiag
            msgFaltantes = msgFaltantes & "[" & cDiag & "] """ & CStr(ws.Cells(1, cDiag).Value) & """" & vbCrLf
        Next cDiag

        MsgBox msgFaltantes, vbCritical
        Exit Sub
    End If

    ' --- Colunas de SAÍDA: a macro cria sozinha se ainda não existirem.        ---
    ' --- Ordem de criação (quando faltar mais de uma): LP, TIPO PRODUTO,      ---
    ' --- MARCA, GAMA, ANIP, RT/OE, DIMENSÃO, ARO, ESTEPE/COMPETICAO — nessa   ---
    ' --- sequência.                                                          ---
    colLp = LocalizarOuCriarColuna(ws, "LP")
    colTipoProduto = LocalizarOuCriarColuna(ws, "TIPO PRODUTO")
    colMarca = LocalizarOuCriarColuna(ws, "MARCA")
    colGama = LocalizarOuCriarColuna(ws, "GAMA")
    colAnip = LocalizarOuCriarColuna(ws, "ANIP")
    colRtOe = LocalizarOuCriarColuna(ws, "RT/OE")
    colDimensao = LocalizarColunaAlternativasOuCriar(ws, Array("DIMENSÃO", "GEOBOX"), "DIMENSÃO")
    colAro = LocalizarOuCriarColuna(ws, "ARO")
    colEstepeComp = LocalizarOuCriarColuna(ws, COL_ESTEPE_COMPETICAO_NOME)

    ' --- Carrega dicionários auxiliares ---
    Dim dicMontadoras As Object, dicAnip As Object, dicExcecoesMarca As Object
    Set dicMontadoras = CarregarMontadoras()
    Set dicAnip = CarregarMarcasAnip()

    Dim dicPadroesLegado As Object
    Set dicPadroesLegado = CarregarPadroesGeoboxLegado()

    Dim dicSegPorGeobox As Object, dicLpPorGeobox As Object, dicLpPorGeoboxMarca As Object
    Dim arrMarcas() As String
    Dim dicGeoboxPorMarca As Object, dicGeoboxGlobalUnicos As Object
    Dim dicGamasPorMarca As Object, dicGamaGlobalUnicos As Object, dicExcecoesGama As Object
    Dim dicMarcaPorGama As Object
    Dim diagnosticoRef As String

    ' --- Pergunta se a base é de BR/MIN (Beyond Road/Mineração). Se SIM, a  ---
    ' --- Tabela de Referência é filtrada pela coluna F ("Base de           ---
    ' --- referência") — só as 4 fontes STORM de BR/MIN entram — em vez de  ---
    ' --- LP, pra não carregar o resto da Referência à toa nessa base.       ---
    Dim somenteMinBr As Boolean
    somenteMinBr = (MsgBox("Esta base é de BR/MIN (Beyond Road / Mineração)?" & vbCrLf & vbCrLf & _
                           "Se SIM, a busca de GEOBOX será restrita à coluna ""Base de referência""" & vbCrLf & _
                           "(STORM 40117090, 40118090, 40119090, 40129090, Dicionário WW e Input manual)." & vbCrLf & _
                           "A busca de MARCA e GAMA continua na Referência inteira.", _
                           vbYesNo + vbQuestion, "Tipo de base") = vbYes)

    ' --- A lista legada (modDicionarios.CarregarPadroesGeoboxLegado) é uma  ---
    ' --- lista fixa de medidas de pneu de carro/caminhão, sem relação com   ---
    ' --- nenhuma LP. Em modo BR/MIN ela NÃO deve entrar como fallback —     ---
    ' --- era ela que fazia a extração de GEOBOX "vazar" medidas de outras  ---
    ' --- LPs (ex: "9.00-20") que não têm nada a ver com a base BR/MIN.      ---
    If somenteMinBr Then Set dicPadroesLegado = CreateObject("Scripting.Dictionary")

    If Not CarregarTabelaReferencia(dicSegPorGeobox, dicLpPorGeobox, dicLpPorGeoboxMarca, arrMarcas, _
                                     dicGeoboxPorMarca, dicGeoboxGlobalUnicos, dicExcecoesMarca, _
                                     dicGamasPorMarca, dicGamaGlobalUnicos, dicMarcaPorGama, dicExcecoesGama, _
                                     somenteMinBr, diagnosticoRef) Then
        MsgBox "Não foi possível abrir a Tabela de Referência em:" & vbCrLf & REF_FILE_PATH, vbCritical
        Exit Sub
    End If

    ' --- Em BR/MIN, a busca de MARCA é só CEGA (ExtrairMarca): exceções +   ---
    ' --- procura direta da marca conhecida na descrição. A descoberta       ---
    ' --- "de trás pra frente" (achou GAMA -> assume a marca dona daquela    ---
    ' --- gama, dentro de ExtrairGama) fica DESLIGADA nesse modo — estava    ---
    ' --- causando confusão de GAMA que se propagava pra MARCA errada.       ---
    ' --- Passar um dicionário vazio neutraliza essa etapa sem mexer em      ---
    ' --- modGama.bas.                                                      ---
    If somenteMinBr Then Set dicMarcaPorGama = CreateObject("Scripting.Dictionary")

    If MOSTRAR_DIAGNOSTICO_REFERENCIA Then
        MsgBox "DIAGNÓSTICO DA TABELA DE REFERÊNCIA:" & vbCrLf & vbCrLf & diagnosticoRef, vbInformation
    End If

    lastRow = UltimaLinhaEntreColunas(ws, Array(colDescricao, colAdquirente))

    If lastRow < 2 Then
        MsgBox "Nenhuma linha de dados encontrada abaixo do cabeçalho." & vbCrLf & vbCrLf & _
               "Planilha ativa: """ & ws.Name & """" & vbCrLf & _
               "Última linha com dado em DESCRIÇÃODA MERCADORIA (coluna " & colDescricao & "): " & _
               ws.Cells(ws.Rows.Count, colDescricao).End(xlUp).Row & vbCrLf & vbCrLf & _
               "Confira se a aba ativa é a correta e se essa coluna tem dados.", vbExclamation
        Exit Sub
    End If

    ' --- Desliga recálculo automático/atualização de tela durante o        ---
    ' --- processamento. Com uma base grande, gravar célula a célula com    ---
    ' --- recálculo automático ligado (padrão do Excel) fica extremamente   ---
    ' --- lento — na prática parece um travamento. Restaurado em Finally.   ---
    Dim calcAnterior As XlCalculation
    calcAnterior = Application.Calculation
    On Error GoTo Finally
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim qtdLinhas As Long
    qtdLinhas = lastRow - 1

    ' --- PERFORMANCE: lê as duas colunas de ENTRADA inteiras num array em   ---
    ' --- memória de uma vez só, em vez de um ws.Cells(i, col) por linha (2  ---
    ' --- chamadas COM por linha). Isso e a escrita em bloco no final são,   ---
    ' --- de longe, a maior causa de lentidão numa base grande — cada        ---
    ' --- ws.Cells é uma chamada COM cara; um array em memória não é.        ---
    Dim arrEntradaDescricao As Variant, arrEntradaAdquirente As Variant
    arrEntradaDescricao = ws.Range(ws.Cells(2, colDescricao), ws.Cells(lastRow, colDescricao)).Value
    arrEntradaAdquirente = ws.Range(ws.Cells(2, colAdquirente), ws.Cells(lastRow, colAdquirente)).Value

    ' --- Arrays de SAÍDA (uma coluna cada, n linhas), preenchidos em memória ---
    ' --- durante o laço e gravados de uma vez só no final (ver Finally/final ---
    ' --- do laço). Mesmo valor que seria escrito célula a célula — só muda   ---
    ' --- COMO chega até a planilha.                                         ---
    Dim arrSaidaLp() As Variant, arrSaidaTipoProduto() As Variant
    Dim arrSaidaMarca() As Variant, arrSaidaGama() As Variant
    Dim arrSaidaAnip() As Variant, arrSaidaRtOe() As Variant
    Dim arrSaidaDimensao() As Variant, arrSaidaAro() As Variant
    Dim arrSaidaEstepeComp() As Variant
    ReDim arrSaidaLp(1 To qtdLinhas, 1 To 1)
    ReDim arrSaidaTipoProduto(1 To qtdLinhas, 1 To 1)
    ReDim arrSaidaMarca(1 To qtdLinhas, 1 To 1)
    ReDim arrSaidaGama(1 To qtdLinhas, 1 To 1)
    ReDim arrSaidaAnip(1 To qtdLinhas, 1 To 1)
    ReDim arrSaidaRtOe(1 To qtdLinhas, 1 To 1)
    ReDim arrSaidaDimensao(1 To qtdLinhas, 1 To 1)
    ReDim arrSaidaAro(1 To qtdLinhas, 1 To 1)
    ReDim arrSaidaEstepeComp(1 To qtdLinhas, 1 To 1)

    ' --- PERFORMANCE: cache de resultado por (DESCRIÇÃO, ADQUIRENTE). Bases  ---
    ' --- de importação costumam repetir a mesma descrição em várias linhas  ---
    ' --- (mesmo modelo de pneu, vários lotes/NCMs). Como toda a extração    ---
    ' --- (MARCA, GAMA, DIMENSÃO, SEGMENTO, LP, ARO, RT/OE, ANIP, ESTEPE/    ---
    ' --- COMPETIÇÃO) é uma função PURA de (descrição, adquirente) — sempre  ---
    ' --- dá o mesmo resultado pra mesma entrada, já que os dicionários      ---
    ' --- carregados no início não mudam durante o laço — a 2ª+ ocorrência   ---
    ' --- da mesma dupla (descrição, adquirente) não precisa refazer nenhuma ---
    ' --- busca: só copia o resultado já calculado. Isso NÃO muda nenhum     ---
    ' --- valor gravado, só evita recalcular o que já foi calculado.         ---
    Dim dicCacheLinhas As Object
    Set dicCacheLinhas = CreateObject("Scripting.Dictionary")

    Dim marca As String, dimensaoBruta As String, dimensaoFinal As String
    Dim descricao As String, adquirente As String
    Dim aro As String, segmento As String, rtOe As String, anip As String
    Dim estepeComp As String, lp As String

    Dim qtdSegmentoPreenchido As Long, qtdSegmentoVazio As Long
    Dim qtdDimensaoVazia As Long
    Dim amostraVazios As String
    qtdSegmentoPreenchido = 0
    qtdSegmentoVazio = 0
    qtdDimensaoVazia = 0
    amostraVazios = ""

    Dim idx As Long
    For i = 2 To lastRow
        idx = i - 1 ' índice 1-based dentro dos arrays de entrada/saída

        ' --- Progresso na barra de status a cada 100 linhas + DoEvents:    ---
        ' --- sem isso, o Windows mostra o Excel como "Não está respondendo" ---
        ' --- durante um laço longo (mesmo funcionando normalmente por      ---
        ' --- baixo), o que é indistinguível de um crash de verdade. Com    ---
        ' --- isso dá pra ver o andamento e confirmar que ainda está vivo.  ---
        If i Mod 100 = 0 Or i = lastRow Then
            Application.StatusBar = "Classificando linha " & i & " de " & lastRow & _
                                     " (" & Format((i - 1) / (lastRow - 1), "0%") & ")"
            DoEvents
        End If

        If qtdLinhas = 1 Then
            descricao = UCase(Trim(CStr(arrEntradaDescricao)))
            adquirente = UCase(Trim(CStr(arrEntradaAdquirente)))
        Else
            descricao = UCase(Trim(CStr(arrEntradaDescricao(idx, 1))))
            adquirente = UCase(Trim(CStr(arrEntradaAdquirente(idx, 1))))
        End If

        Dim chaveCache As String
        chaveCache = descricao & Chr(1) & adquirente

        If dicCacheLinhas.Exists(chaveCache) Then
            ' --- Já calculado numa linha anterior com a mesma DESCRIÇÃO e  ---
            ' --- ADQUIRENTE: reaproveita sem refazer nenhuma extração.     ---
            Dim cache As Variant
            cache = dicCacheLinhas(chaveCache)
            marca = cache(0)
            Dim gamaLinha As String
            gamaLinha = cache(1)
            dimensaoBruta = cache(2)
            dimensaoFinal = cache(3)
            rtOe = cache(4)
            aro = cache(5)
            anip = cache(6)
            segmento = cache(7)
            lp = cache(8)
            estepeComp = cache(9)
        Else
            ' --- MARCA: extraída da descrição (nome conhecido na Tabela de Referência) ---
            marca = ExtrairMarca(descricao, arrMarcas, dicExcecoesMarca, somenteMinBr)

            ' --- GAMA: extraída da descrição (exceções/abreviações primeiro,   ---
            ' --- depois nome completo conhecido na Tabela de Referência). Se   ---
            ' --- a MARCA não foi encontrada, ExtrairGama também tenta          ---
            ' --- descobrir a marca a partir da gama achada (gama é tratada    ---
            ' --- como praticamente exclusiva de uma marca) — por isso "marca" ---
            ' --- é passada ByRef e pode voltar preenchida daqui.              ---
            gamaLinha = ExtrairGama(descricao, marca, dicGamasPorMarca, dicGamaGlobalUnicos, dicMarcaPorGama, dicExcecoesGama, somenteMinBr)

            ' --- DIMENSÃO: extraída da descrição (medida conhecida na Tabela de Referência) ---
            dimensaoBruta = ExtrairDimensao(descricao, marca, dicGeoboxPorMarca, dicGeoboxGlobalUnicos, dicPadroesLegado, somenteMinBr)

            ' --- DIMENSÃO: substitui todo e qualquer hífen por "R" antes de gravar ---
            ' --- (pulado em bases BR/MIN: lá o "-" faz parte de GEOBOX válidos    ---
            ' --- — ex: "7.50-16", "9.00-20" — e a troca forçada corrompe o valor, ---
            ' --- fazendo comparações futuras darem vazio).                       ---
            If somenteMinBr Then
                dimensaoFinal = dimensaoBruta
            Else
                dimensaoFinal = Replace(dimensaoBruta, "-", "R")
            End If

            ' --- RT/OE ---
            rtOe = "RT"
            Dim chaveM As Variant
            For Each chaveM In dicMontadoras.Keys
                If InStr(1, adquirente, chaveM, vbTextCompare) > 0 Then
                    rtOe = "OE"
                    Exit For
                End If
            Next chaveM

            ' --- ARO (a partir da DIMENSÃO já com hífen convertido) ---
            aro = ExtrairAro(dimensaoFinal)

            ' --- ANIP ---
            anip = "Importado"
            If dicAnip.Exists(marca) Then anip = "ANIP"

            ' --- TIPO PRODUTO e LP: busca direto pela DIMENSÃO na Tabela de     ---
            ' --- Referência (aba "Referencia"). SEGMENTO e LP são votações     ---
            ' --- independentes (maioria dos valores não vazios daquele         ---
            ' --- GEOBOX); se achou SEGMENTO, o LP é deduzido dele (TC/PL/BR).  ---
            ' --- Usa a DIMENSÃO BRUTA (antes do hífen->R) porque é essa forma  ---
            ' --- que bate exatamente com o GEOBOX gravado na Tabela.           ---
            ObterSegmentoELpPorDimensao dimensaoBruta, dicSegPorGeobox, dicLpPorGeobox, segmento, lp

            ' --- BR/MIN: antes de aceitar o LP achado só por GEOBOX acima,       ---
            ' --- tenta primeiro a combinação GEOBOX + MARCA (mais específica —   ---
            ' --- LP dominante entre as linhas da Referência com essa marca       ---
            ' --- E esse geobox exatos). Só cai no LP por GEOBOX puro (já em lp)  ---
            ' --- se essa combinação não existir na Referência.                   ---
            If somenteMinBr And Len(marca) > 0 Then
                Dim chaveLpGeoMarca As String
                chaveLpGeoMarca = dimensaoBruta & "|@|" & marca
                If dicLpPorGeoboxMarca.Exists(chaveLpGeoMarca) Then
                    lp = CStr(dicLpPorGeoboxMarca(chaveLpGeoMarca))
                End If
            End If

            ' --- LP: se a Tabela não trouxe LP pra essa dimensão, cai no      ---
            ' --- fallback de deduzir pelo SEGMENTO.                           ---
            If Len(lp) = 0 Then lp = DeduzirLp(segmento)

            ' --- BR/MIN: o TIPO PRODUTO manda por cima de tudo o que veio      ---
            ' --- antes (combinação GEOBOX+MARCA, GEOBOX puro, dedução) — só     ---
            ' --- nesse modo. DeduzirLpBrMin cobre TODOS os TIPO PRODUTO         ---
            ' --- válidos em BR/MIN (PL, TC, BR e MIN) — um TIPO PRODUTO nunca   ---
            ' --- pode "vazar" pra uma LP que ele não representa.                ---
            If somenteMinBr Then
                Dim lpBrMin As String
                lpBrMin = DeduzirLpBrMin(segmento)
                If Len(lpBrMin) > 0 Then lp = lpBrMin
            End If

            ' --- ARO terminado em ".5" força "PL" — mas só como ÚLTIMO recurso, ---
            ' --- se depois de TODAS as etapas acima (GEOBOX+MARCA, GEOBOX puro, ---
            ' --- dedução por SEGMENTO, override de TIPO PRODUTO em BR/MIN)      ---
            ' --- nenhuma LP foi determinada.                                    ---
            If Len(lp) = 0 And AroTerminaEmMeio(aro) Then lp = "PL"

            ' --- ESTEPE/COMPETIÇÃO/LCV: mesma DESCRIÇÃODA MERCADORIA já lida    ---
            ' --- acima, sem depender de MARCA/GAMA/DIMENSÃO/LP (ver premissa    ---
            ' --- no cabeçalho do módulo modEstepeCompeticao).                   ---
            estepeComp = ClassificarEstepeCompeticao(descricao)

            dicCacheLinhas.Add chaveCache, Array(marca, gamaLinha, dimensaoBruta, dimensaoFinal, _
                                                  rtOe, aro, anip, segmento, lp, estepeComp)
        End If

        arrSaidaMarca(idx, 1) = marca
        arrSaidaGama(idx, 1) = gamaLinha
        arrSaidaDimensao(idx, 1) = dimensaoFinal
        arrSaidaRtOe(idx, 1) = rtOe
        arrSaidaAro(idx, 1) = aro
        arrSaidaAnip(idx, 1) = anip
        arrSaidaTipoProduto(idx, 1) = segmento
        arrSaidaLp(idx, 1) = lp
        arrSaidaEstepeComp(idx, 1) = estepeComp

        ' --- Diagnóstico: conta preenchidos/vazios e guarda uma amostra ---
        If Len(dimensaoBruta) = 0 Then
            qtdDimensaoVazia = qtdDimensaoVazia + 1
        End If
        If Len(segmento) > 0 Then
            qtdSegmentoPreenchido = qtdSegmentoPreenchido + 1
        Else
            qtdSegmentoVazio = qtdSegmentoVazio + 1
            Dim existeNaTabela As Boolean
            existeNaTabela = dicSegPorGeobox.Exists(dimensaoBruta)
            amostraVazios = amostraVazios & "Linha " & i & ": DIMENSÃO=""" & dimensaoBruta & _
                            """ (existe no mapa de SEGMENTO? " & IIf(existeNaTabela, "SIM", "NÃO")
            If dicLpPorGeobox.Exists(dimensaoBruta) Then
                amostraVazios = amostraVazios & " | LP votado: """ & CStr(dicLpPorGeobox(dimensaoBruta)) & """"
            End If
            amostraVazios = amostraVazios & ")" & vbCrLf
        End If

    Next i

    ' --- PERFORMANCE: grava cada coluna de SAÍDA de uma vez só (um Range.Value ---
    ' --- por coluna, em vez de um ws.Cells(i, col).Value por linha por coluna). ---
    ws.Range(ws.Cells(2, colLp), ws.Cells(lastRow, colLp)).Value = arrSaidaLp
    ws.Range(ws.Cells(2, colTipoProduto), ws.Cells(lastRow, colTipoProduto)).Value = arrSaidaTipoProduto
    ws.Range(ws.Cells(2, colMarca), ws.Cells(lastRow, colMarca)).Value = arrSaidaMarca
    ws.Range(ws.Cells(2, colGama), ws.Cells(lastRow, colGama)).Value = arrSaidaGama
    ws.Range(ws.Cells(2, colAnip), ws.Cells(lastRow, colAnip)).Value = arrSaidaAnip
    ws.Range(ws.Cells(2, colRtOe), ws.Cells(lastRow, colRtOe)).Value = arrSaidaRtOe
    ws.Range(ws.Cells(2, colDimensao), ws.Cells(lastRow, colDimensao)).Value = arrSaidaDimensao
    ws.Range(ws.Cells(2, colAro), ws.Cells(lastRow, colAro)).Value = arrSaidaAro
    ws.Range(ws.Cells(2, colEstepeComp), ws.Cells(lastRow, colEstepeComp)).Value = arrSaidaEstepeComp

    ' --- Restaura o estado do Excel assim que o processamento linha a     ---
    ' --- linha termina (o resto daqui pra baixo só mexe numa aba auxiliar  ---
    ' --- pequena, não precisa mais ficar com recálculo desligado).         ---
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = calcAnterior
    Application.StatusBar = False

    ' --- Joga a lista completa de linhas vazias numa aba auxiliar, mais    ---
    ' --- fácil de ler/copiar do que uma caixa de mensagem gigante.        ---
    Dim wsDiagVazios As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Diagnostico_Vazios").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsDiagVazios = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsDiagVazios.Name = "Diagnostico_Vazios"
    wsDiagVazios.Cells(1, 1).Value = "LINHA"
    wsDiagVazios.Cells(1, 2).Value = "DIMENSAO"
    wsDiagVazios.Cells(1, 3).Value = "EXISTE_NA_TABELA"
    wsDiagVazios.Cells(1, 4).Value = "SEGMENTO_GUARDADO"
    wsDiagVazios.Cells(1, 5).Value = "LP_GUARDADO"

    Dim linhaDiag As Long
    linhaDiag = 2
    Dim linhaTexto As Variant
    For Each linhaTexto In Split(amostraVazios, vbCrLf)
        If Len(Trim(CStr(linhaTexto))) > 0 Then
            wsDiagVazios.Cells(linhaDiag, 1).Value = linhaTexto
            linhaDiag = linhaDiag + 1
        End If
    Next linhaTexto

    MsgBox "Classificação concluída com sucesso!" & vbCrLf & _
           "Linhas processadas: " & (lastRow - 1) & vbCrLf & vbCrLf & _
           "--- DIAGNÓSTICO TIPO PRODUTO ---" & vbCrLf & _
           "GEOBOX únicos com SEGMENTO mapeado: " & dicSegPorGeobox.Count & vbCrLf & _
           "GEOBOX únicos com LP mapeado: " & dicLpPorGeobox.Count & vbCrLf & _
           "Linhas com DIMENSÃO vazia (não deu pra nem tentar buscar): " & qtdDimensaoVazia & vbCrLf & _
           "Linhas com TIPO PRODUTO preenchido: " & qtdSegmentoPreenchido & vbCrLf & _
           "Linhas com TIPO PRODUTO vazio: " & qtdSegmentoVazio & vbCrLf & vbCrLf & _
           "Lista completa das linhas vazias foi para a aba ""Diagnostico_Vazios"".", vbInformation

    Exit Sub

Finally:
    ' --- Garante que o Excel não fique "travado" em recálculo manual/tela ---
    ' --- parada se a macro cair aqui por causa de um erro no meio do laço. ---
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = calcAnterior
    Application.StatusBar = False
    MsgBox "A macro parou por causa de um erro: " & vbCrLf & vbCrLf & Err.Description, vbCritical

End Sub
