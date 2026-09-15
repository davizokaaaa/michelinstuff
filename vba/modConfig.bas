Attribute VB_Name = "modConfig"
Option Explicit

' ==========================================================================
' MODULO: modConfig
' Constantes de configuracao global da macro de classificacao.
' ==========================================================================

' ==========================================================================
' Caminho do arquivo Tabela_Referencia.xlsx. Era uma Const de texto simples,
' mas o "Á" de "Área de Trabalho" corrompia ao colar/salvar no VBA Editor
' (mesmo bug de codificação dos nomes de coluna — ver modPlanilha.
' NormalizarTexto), fazendo Workbooks.Open procurar um caminho com bytes
' diferentes do real e falhar com "Não foi possível abrir". Por isso virou
' uma Function que monta a string em tempo de execução com ChrW(193) — o
' código-fonte deste arquivo não tem nenhum caractere acentuado literal
' nessa parte, então não tem o que corromper.
' ==========================================================================
Public Function REF_FILE_PATH() As String
    REF_FILE_PATH = "C:\Users\E125949\OneDrive - MFP Michelin\" & _
                     ChrW(193) & "rea de Trabalho\Teste importados\Tabela_Referencia.xlsx"
End Function

Public Const COL_ADQUIRENTE_NOME As String = "PROVÁVEL ADQUIRENTE"
Public Const MOSTRAR_DIAGNOSTICO_REFERENCIA As Boolean = True ' mude para False depois de confirmar que está lendo certo

' Nome da coluna de SAÍDA da feature ESTEPE/COMPETIÇÃO/LCV (ver modEstepeCompeticao).
' Sem cedilha/acento de propósito, mesmo padrão de nomes de coluna já usado
' no restante do projeto (ex: "DIMENSÃO" é a exceção, não a regra).
Public Const COL_ESTEPE_COMPETICAO_NOME As String = "ESTEPE/COMPETICAO"
