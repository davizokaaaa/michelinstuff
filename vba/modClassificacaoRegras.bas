Attribute VB_Name = "modClassificacaoRegras"
Option Explicit

' ==========================================================================
' MODULO: modClassificacaoRegras
' Regras de negocio "puras" de classificacao: deducao de LP a partir do
' SEGMENTO e a regra de ARO terminado em ".5" forcando LP="PL".
' (RT/OE e ANIP ficam direto no orquestrador por serem checagens simples
' de uma linha so contra um dicionario; ObterSegmentoELpPorDimensao esta
' em modReferencia por depender da Tabela de Referencia carregada.)
' ==========================================================================

' ==========================================================================
' Deduz a LP a partir do SEGMENTO (TIPO PRODUTO):
'   TLD, PPL -> PL
'   PC, REC, COM -> TC
'   Qualquer outro caso (incl. vazio) -> vazio
' ==========================================================================
Function DeduzirLp(segmento As String) As String

    Select Case UCase(Trim(segmento))
        Case "TLD", "PPL", "BUS"
            DeduzirLp = "PL"
        Case "PC", "REC", "COM"
            DeduzirLp = "TC"
        Case "DM"
            DeduzirLp = "BR"
        Case Else
            DeduzirLp = ""
    End Select

End Function

' ==========================================================================
' Verifica se o ARO termina em ".5" (ex: R17.5, R19.5, R22.5, R24.5) --
' regra que forca LP = "PL", independente do que o SEGMENTO indicaria.
' ==========================================================================
Function AroTerminaEmMeio(aro As String) As Boolean
    AroTerminaEmMeio = (Right(Trim(aro), 2) = ".5")
End Function

' ==========================================================================
' Override de LP a partir do TIPO PRODUTO (SEGMENTO) -- So pra bases BR/MIN
' (chamado de modMain so quando somenteMinBr = True). Cada TIPO PRODUTO so
' pode pertencer a LP que ele representa -- essa regra tem prioridade final
' sobre tudo o que veio antes (combinacao GEOBOX+MARCA, GEOBOX puro,
' DeduzirLp). A regra do ARO ".5" so e aplicada DEPOIS desta, e so se
' nenhuma LP tiver sido determinada em nenhuma etapa (ver modMain).
'   TLD, PPL, BUS -> "PL"
'   REC, COM, PC -> "TC"
'   AG, CL, CO, COLHEITADEIRA, COLHEITADEIRA/HPT, DM, HPT, IMPLEMENTO,
'   INDUS, INFRA, LPT, MH, MILITAR, MPT, OUTRO, PNEUMATICO, PORTOS, PORTS,
'   PULVERIZADOR, QUADRICICLO, SM, SOLIDO -> "BR"
'   MIN -> "MIN"
'   Qualquer outro TIPO PRODUTO (incl. vazio) -> "" (nao forca nada, mantem
'   o LP que ja tinha sido determinado antes desta regra).
' ==========================================================================
Function DeduzirLpBrMin(segmento As String) As String
    Static dicLpPorSegmentoBrMin As Object
    If dicLpPorSegmentoBrMin Is Nothing Then
        Set dicLpPorSegmentoBrMin = CreateObject("Scripting.Dictionary")

        dicLpPorSegmentoBrMin.Add "TLD", "PL"
        dicLpPorSegmentoBrMin.Add "PPL", "PL"
        dicLpPorSegmentoBrMin.Add "BUS", "PL"

        dicLpPorSegmentoBrMin.Add "REC", "TC"
        dicLpPorSegmentoBrMin.Add "COM", "TC"
        dicLpPorSegmentoBrMin.Add "PC", "TC"

        dicLpPorSegmentoBrMin.Add "AG", "BR"
        dicLpPorSegmentoBrMin.Add "CL", "BR"
        dicLpPorSegmentoBrMin.Add "CO", "BR"
        dicLpPorSegmentoBrMin.Add "COLHEITADEIRA", "BR"
        dicLpPorSegmentoBrMin.Add "COLHEITADEIRA/HPT", "BR"
        dicLpPorSegmentoBrMin.Add "DM", "BR"
        dicLpPorSegmentoBrMin.Add "HPT", "BR"
        dicLpPorSegmentoBrMin.Add "IMPLEMENTO", "BR"
        dicLpPorSegmentoBrMin.Add "INDUS", "BR"
        dicLpPorSegmentoBrMin.Add "INFRA", "BR"
        dicLpPorSegmentoBrMin.Add "LPT", "BR"
        dicLpPorSegmentoBrMin.Add "MH", "BR"
        dicLpPorSegmentoBrMin.Add "MILITAR", "BR"
        dicLpPorSegmentoBrMin.Add "MPT", "BR"
        dicLpPorSegmentoBrMin.Add "OUTRO", "BR"
        dicLpPorSegmentoBrMin.Add "PNEUMATICO", "BR"
        ' Chave extra com acento de verdade (montada via ChrW, nao como
        ' caractere literal no arquivo-fonte -- ver nota no topo do modulo
        ' sobre corrupcao de encoding), caso a coluna TIPO PRODUTO real
        ' venha escrita "PNEUMATICO" (com A acentuado).
        dicLpPorSegmentoBrMin.Add "PNEUM" & ChrW(193) & "TICO", "BR"
        dicLpPorSegmentoBrMin.Add "PORTOS", "BR"
        dicLpPorSegmentoBrMin.Add "PORTS", "BR"
        dicLpPorSegmentoBrMin.Add "PULVERIZADOR", "BR"
        dicLpPorSegmentoBrMin.Add "QUADRICICLO", "BR"
        dicLpPorSegmentoBrMin.Add "SM", "BR"
        dicLpPorSegmentoBrMin.Add "SOLIDO", "BR"
        ' Mesma logica: chave extra acentuada via ChrW para "SOLIDO".
        dicLpPorSegmentoBrMin.Add "S" & ChrW(211) & "LIDO", "BR"

        dicLpPorSegmentoBrMin.Add "MIN", "MIN"
    End If

    Dim segNorm As String
    segNorm = UCase(Trim(segmento))

    If dicLpPorSegmentoBrMin.Exists(segNorm) Then
        DeduzirLpBrMin = CStr(dicLpPorSegmentoBrMin(segNorm))
    Else
        DeduzirLpBrMin = ""
    End If
End Function
