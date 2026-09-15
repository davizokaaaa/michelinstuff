Attribute VB_Name = "modDicionarios"
Option Explicit

' ==========================================================================
' MODULO: modDicionarios
' Dicionarios fixos usados pelas regras de classificacao: montadoras (OE),
' marcas ANIP, e a lista legada de padroes de GEOBOX.
' ==========================================================================

' ==========================================================================
' Dicionário de montadoras (mesmo da macro OE() original)
' ==========================================================================
Function CarregarMontadoras() As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    dic.Add "VOLKSWAGEN DO BRASIL INDUSTRIA DE VEICULOS AUTOMOTORES LTDA", True
    dic.Add "FCA FIAT CHRYSLER AUTOMOVEIS BRASIL LTDA", True
    dic.Add "HYUNDAI MOTOR BRASIL MONTADORA DE AUTOMOVEIS LTDA", True
    dic.Add "STELLANTIS AUTOMOVEIS BRASIL LTDA.", True
    dic.Add "NISSAN DO BRASIL AUTOMOVEIS LTDA", True
    dic.Add "BMW DO BRASIL LTDA", True
    dic.Add "HONDA AUTOMOVEIS DO BRASIL LTDA", True
    dic.Add "FORD MOTOR COMPANY BRASIL LTDA", True
    dic.Add "GENERAL MOTORS DO BRASIL LTDA", True
    dic.Add "RENAULT DO BRASIL LTDA.", True
    dic.Add "SCANIA LATIN AMERICA LTDA", True
    dic.Add "VOLVO DO BRASIL VEICULOS LTDA", True

    Set CarregarMontadoras = dic
End Function

' ==========================================================================
' Dicionário de marcas ANIP (mesmo da macro ANIP() original)
' ==========================================================================
Function CarregarMarcasAnip() As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    dic.Add "MICHELIN", True
    dic.Add "PIRELLI", True
    dic.Add "BRIDGESTONE", True
    dic.Add "CONTINENTAL", True
    dic.Add "GENERAL TIRES", True
    dic.Add "GENERAL TIRE", True
    dic.Add "DUNLOP", True
    dic.Add "GOODYEAR", True
    dic.Add "FIRESTONE", True
    dic.Add "RINALDI", True
    dic.Add "TORTUGA", True
    dic.Add "PROMETEON", True

    Set CarregarMarcasAnip = dic
End Function

' ==========================================================================
' Lista específica de medidas (GEOBOX) vinda da macro legada Sub dimensão():
' cada padrão de texto observado mapeado para o valor "oficial" — inclui
' casos hardcoded que não seguem um formato genérico (ex: "275/80" sozinho
' assume "275/80R22.5"; "13R22.5" assume "13.00R22.5"; "205R14C" vira
' "205R14"). Usada como PASSO 3 (fonte extra) na extração de DIMENSÃO.
' ==========================================================================
Function CarregarPadroesGeoboxLegado() As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    dic.Add "295/80R22.5", "295/80R22.5"
    dic.Add "295/80 R22.5", "295/80R22.5"
    dic.Add "295 80R22.5", "295/80R22.5"
    dic.Add "295MM/PERFIL:80%/ARO: 22.5", "295/80R22.5"
    dic.Add "275/80R22.5", "275/80R22.5"
    dic.Add "275/80 R 22.5", "275/80R22.5"
    dic.Add "275MM/PERFIL:80%/ARO: 22.5", "275/80R22.5"
    dic.Add "275/80", "275/80R22.5"
    dic.Add "215/75R17.5", "215/75R17.5"
    dic.Add "235/75R17.5", "235/75R17.5"
    dic.Add "235/75 R17.5", "235/75R17.5"
    dic.Add "175/70R14", "175/70R14"
    dic.Add "7.50-16", "7.50-16"
    dic.Add "185R14", "185R14"
    dic.Add "12.00R24", "12.00R24"
    dic.Add "205/70R15", "205/70R15"
    dic.Add "225/75R16", "225/75R16"
    dic.Add "205/75R16", "205/75R16"
    dic.Add "175/65R14", "175/65R14"
    dic.Add "31X10.50R15", "31X10.50R15"
    dic.Add "10.00-20", "10.00-20"
    dic.Add "265/70R16", "265/70R16"
    dic.Add "265/65R17", "265/65R17"
    dic.Add "265/65 R17", "265/65R17"
    dic.Add "265/75R16", "265/75R16"
    dic.Add "9.00-20", "9.00-20"
    dic.Add "265/70R17", "265/70R17"
    dic.Add "265/70 R17", "265/70R17"
    dic.Add "225/65R16", "225/65R16"
    dic.Add "225/65R17", "225/65R17"
    dic.Add "325/95R24", "325/95R24"
    dic.Add "195/70R15", "195/70R15"
    dic.Add "285/75R16", "285/75R16"
    dic.Add "225/70R15", "225/70R15"
    dic.Add "265/60R18", "265/60R18"
    dic.Add "265/60 R18", "265/60R18"
    dic.Add "215/65R16", "215/65R16"
    dic.Add "215/65 R16", "215/65R16"
    dic.Add "265/65R18", "265/65R18"
    dic.Add "265/65 R18", "265/65R18"
    dic.Add "315/80R22.5", "315/80R22.5"
    dic.Add "195R14", "195R14"
    dic.Add "10.00R20", "10.00R20"
    dic.Add "275/65R18", "275/65R18"
    dic.Add "275/65 R18", "275/65R18"
    dic.Add "245/70R16", "245/70R16"
    dic.Add "275/70R22.5", "275/70R22.5"
    dic.Add "215/75R16", "215/75R16"
    dic.Add "225/60R18", "225/60R18"
    dic.Add "285/70R17", "285/70R17"
    dic.Add "285/70 R17", "285/70R17"
    dic.Add "195/75R16", "195/75R16"
    dic.Add "11.00-22", "11.00-22"
    dic.Add "7.50R16", "7.50R16"
    dic.Add "265/70R18", "265/70R18"
    dic.Add "235/75R15", "235/75R15"
    dic.Add "285/65R18", "285/65R18"
    dic.Add "235/85R16", "235/85R16"
    dic.Add "255/55R19", "255/55R19"
    dic.Add "245/75R16", "245/75R16"
    dic.Add "265/60R20", "265/60R20"
    dic.Add "265/60R19", "265/60R19"
    dic.Add "235/70R16", "235/70R16"
    dic.Add "245/65R17", "245/65R17"
    dic.Add "285/70R19.5", "285/70R19.5"
    dic.Add "315/75R16", "315/75R16"
    dic.Add "275/70R18", "275/70R18"
    dic.Add "235/65R16", "235/65R16"
    dic.Add "305/70R16", "305/70R16"
    dic.Add "195R15", "195R15"
    dic.Add "255/70R16", "255/70R16"
    dic.Add "275/60R20", "275/60R20"
    dic.Add "215/75R15", "215/75R15"
    dic.Add "215/ 75R17.5", "215/75R15"
    dic.Add "215/70R15", "215/70R15"
    dic.Add "32X11.50R15", "32X11.50R15"
    dic.Add "35X12.50R20", "35X12.50R20"
    dic.Add "225/70R17", "225/70R17"
    dic.Add "285/75R17", "285/75R17"
    dic.Add "245/70R17", "245/70R17"
    dic.Add "385/65R22.5", "385/65R22.5"
    dic.Add "385/65 R22.5", "385/65R22.5"
    dic.Add "35X12.50R17", "35X12.50R17"
    dic.Add "195/65R16", "195/65R16"
    dic.Add "35X12.50R18", "35X12.5R18"
    dic.Add "16.00R20", "16.00R20"
    dic.Add "33/12.50R18", "33X12.50R18"
    dic.Add "205/75 R14", "205/75R14"
    dic.Add "24R21", "24.00R21"
    dic.Add "12.00R20", "12.00R20"
    dic.Add "395/85R20", "395/85R20"
    dic.Add "13R22.5", "13.00R22.5"
    dic.Add "295/60R22.5", "295/60R22.5"
    dic.Add "255/55R20", "255/55R20"
    dic.Add "14.00R20", "14.00R20"
    dic.Add "11.00R22", "11.00R22"
    dic.Add "305/75R24.5", "305/75R24.5"
    dic.Add "7.00-16", "7.00-16"
    dic.Add "35-12.50R17", "35X12.50R17"
    dic.Add "245/60R18", "245/60R18"
    dic.Add "225/55R18", "225/55R18"
    dic.Add "275/65R17", "275/65R17"
    dic.Add "235/60R18", "235/60R18"
    dic.Add "205R14C", "205R14"
    dic.Add "31X10.5R15", "31X10.5R15"
    dic.Add "35X12.50R15", "35X12.50R15"
    dic.Add "33X12.50R18", "33X12.50R18"

    Set CarregarPadroesGeoboxLegado = dic
End Function
