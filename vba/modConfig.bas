Attribute VB_Name = "modConfig"
Option Explicit

' ==========================================================================
' MODULO: modConfig
' Constantes de configuracao global da macro de classificacao.
' ==========================================================================

Public Const REF_FILE_PATH As String = "C:\Users\E125949\OneDrive - MFP Michelin\Área de Trabalho\Teste importados\Tabela_Referencia.xlsx"
Public Const COL_ADQUIRENTE_NOME As String = "PROVÁVEL ADQUIRENTE"
Public Const MOSTRAR_DIAGNOSTICO_REFERENCIA As Boolean = True ' mude para False depois de confirmar que está lendo certo

' Nome da coluna de SAÍDA da feature ESTEPE/COMPETIÇÃO/LCV (ver modEstepeCompeticao).
' Sem cedilha/acento de propósito, mesmo padrão de nomes de coluna já usado
' no restante do projeto (ex: "DIMENSÃO" é a exceção, não a regra).
Public Const COL_ESTEPE_COMPETICAO_NOME As String = "ESTEPE/COMPETICAO"
