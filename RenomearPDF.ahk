/************************************************************************
 * @description Renomeia arquivos PDF diversos para melhor organização
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/15
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0
#Include pdf2var.ahk

SetWorkingDir(A_LineFile '\..\..\.\data')

NomeNorm := Map()
Loop read 'obs_lista_ng.txt'
{
    col := StrSplit(A_LoopReadLine, '|')
    NomeNorm[col[2]] := Trim(col[1])
}

NomeMes := Map(), NomeMes.CaseSense := false
NomeMes.Set('jan','01','fev','02','mar','03','abr','04','mai','05','jun','06',
            'jul','07','ago','08','set','09','out','10','nov','11','dez','12')

NomeDarf := Map('8109', 'PIS', '6912', 'PIS', '2172', 'COFINS', '5856', 'COFINS',
    '2089', 'IRPJ', '2372', 'CSLL', '1708', 'Ret IRRF', '5952', 'Ret PIS-COFINS-CSLL', '4805', 'Requisicao de Selos',
    '1082', 'DARF Previdenciario', '1213', 'DARF Previdenciario', '1099', 'DARF Previdenciario', '4133', 'Div Ativa')

NomeDAE := Map('0215-4', 'ICMS Frete', '0209-7', 'ICMS-ST Entradas', '0211-3', 'ICMS-ST Bebidas',
    '0221-2', 'ICMS-ST Industria', '0326-9', 'Antecipação ICMS Comercio', '0327-7', 'Antecipação ICMS Industria',
    '0317-8', 'DIFAL', '0120-6', 'ICMS', '0112-3', 'ICMS', '0121-4', 'ICMS',
    '0115-6', 'ICMS', '0806', 'ICMS', '0791', 'DIFAL', '2036', 'FECP', '0313-7', 'ICMS-ST Antecipado')

/**
 * O script pode ser usado como biblioteca ou usado de forma stand-alone
 * Se for executado sozinho, abrirá caixa de diálogo para seleção dos arquivos
 */
if A_LineFile = A_ScriptFullPath
{
    dir_open := A_Args.Length ? A_Args[1] : ''
    select_files := FileSelect('M', dir_open, 'Selecione o arquivos', 'Arquivos PDF (*.pdf)')
    if not select_files.Length
        ExitApp()

    for filename in select_files
    {
        if namepdf := ExtrairNomePDF(filename)
            RenomearArquivoPDF(filename, namepdf)
        else
            MsgBox('Não foi possível extrair nome', 'Renomear PDF', '16 T2')
    }
    MsgBox('Concluído', 'Renomear PDF', '64 T2')
}

RenomearArquivoPDF(filepath, namepdf)
{
    SplitPath(filepath,, &Dir), inc := 0
    namepdf_inc := (namepdf := Dir '\' namepdf) '.pdf'
    while FileExist(namepdf_inc)
        namepdf_inc := Format('{} ({}).pdf', namepdf, ++inc)

    FileMove(filepath, namepdf_inc)
}

ExtrairNomePDF(filepath)
{
    global texto := pdf2var(filepath)

    if InStr(texto, 'Documento de Arrecadação de Receitas Federais')
        return ExtrairNomeDARF()

    else if InStr(texto, 'Documento de Arrecadação do Simples Nacional')
        return ExtrairNomeDAS()

    else if InStr(texto, 'Imposto Sobre Serviços de Qualquer Natureza')
        return ExtrairNomeISS()

    else if InStr(texto, 'DOCUMENTO DE ARRECADAÇÃO ESTADUAL - DAE')
        return ExtrairNomeDAE()

    else if InStr(texto, 'MUNICÍPIO`r`nJUAZEIRO')
        return ExtrairNomeDAEBahia()

    else if InStr(texto, 'SINTEGRA')
        return ExtrairNomeSintegra()

    else if InStr(texto, 'DAPI Modelo 1')
        return ExtrairNomeDAPI()

    else if InStr(texto, 'Versão Sped Fiscal')
        return ExtrairNomeSpedFiscal()

    else if InStr(texto, 'Versão EFD-Contribuições')
        return ExtrairNomeSpedContribuicoes()

    else if InStr(texto, 'DeSTDA')
        return ExtrairNomeDeSTDA()

    else if InStr(texto, 'Município de Papagaios - MG')
        return ExtrairLivroISSPrefeitura()

    else if InStr(texto, 'SECRETARIA DE FINANÇAS - SEMFIN Regime Especial')
        return ExtrairLivroISS()

    else if InStr(texto, 'Extrato do Simples Nacional')
        return ExtrairExtratoPGDAS()

    else if InStr(texto, 'LIVRO REGISTRO DE ENTRADAS')
        return ExtrairLivroEntradas()

    else if Instr(texto, 'LIVRO REGISTRO DE SAÍDAS')
        return ExtrairLivroSaidas()

    else if Instr(texto, 'Registro de Apuração do ICMS')
        return ExtrairLivroApuracaoICMS()

    else if Instr(texto, 'Registro de Apuração do IPI')
        return ExtrairLivroApuracaoIPI()

    else
        return false
}

RegExCNPJ(pos := 1)
{
    if RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj, pos)
        return RegExReplace(cnpj[0], '\D')
}

RegExData(pos := 1)
{
    if RegExMatch(texto, '(\d\d\/)(\d\d)\/(\d\d\d\d)', &data, pos)
        return data[2] '-' data[3]
}

ExtrairNomeSintegra()
{
    cnpj := RegExCNPJ(), data := RegExData()
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Rec Sintegra' A_Space data
}

ExtrairNomeDAPI()
{
    cnpj := RegExCNPJ(), data := RegExData(InStr(texto, 'PERÍODO DE REFERÊNCIA'))
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Rec DAPI' A_Space data
}

ExtrairNomeSpedFiscal()
{
    cnpj := RegExCNPJ(), data := RegExData()
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Rec Sped Fiscal' A_Space data
}

ExtrairNomeSpedContribuicoes()
{
    cnpj := RegExCNPJ(), data := RegExData()
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Rec Sped Contribuições' A_Space data
}

ExtrairNomeDeSTDA()
{
    cnpj := RegExCNPJ()
    RegExMatch(texto, '([a-z]{3})\/(\d{4})', &data)
    data := NomeMes[data[1]] '-' data[2]
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Rec DeSTDA' A_Space data
}

ExtrairNomeDARF()
{
    cnpj := RegExCNPJ(), data := RegExData()
    RegExMatch(texto, 'm)((^\d{4}$)|Código.* (\d{4}))', &cod)
    cod := cod[2] ? cod[2] : cod[3]
    return NomeNorm.Get(cnpj, cnpj) A_Space NomeDarf.Get(cod, cod) A_Space data
}

ExtrairNomeDAS()
{
    cnpj := RegExCNPJ()
    RegExMatch(texto, '([A-Za-z]{3}).*\/(\d{4})', &data)
    data := NomeMes[data[1]] '-' data[2], nome := 'DAS'
    if Instr(texto, 'DAS de PARCSN')
    {
        data := RegExData(Instr(texto, 'Data de Vencimento'))
        nome := 'Parcelamento DAS'
    }
    return NomeNorm.Get(cnpj, cnpj) A_Space nome A_Space data
}

ExtrairNomeISS()
{
    cnpj := RegExCNPJ(InStr(texto, 'Dados do Contribuinte'))
    data := RegExData(InStr(texto, 'MÊS/ANO REFERÊNCIA'))
    return NomeNorm.Get(cnpj, cnpj) A_Space 'ISSQN' A_Space data
}

ExtrairNomeDAE()
{
    if RegExMatch(texto, 'PARCELA: (\d{3})\/(\d{3})', &ref)
    {
        ref := ref[1] '-' ref[2]
        if RegExMatch(texto, '\d{9}\.\d{2}\-?\d{2}', &insc, Instr(texto, 'IDENTIFICAÇÃO', 'Locale'))
            cnpj := BuscarCNPJ(insc[0])

        else if RegExMatch(texto, '\d{2}\.\d{6}\.\d{4}\/\d{2}', &insc, Instr(texto, 'IDENTIFICAÇÃO', 'Locale'))
            cnpj := RegExReplace(cnpj[0], '\D')

        return NomeNorm.Get(cnpj, cnpj) A_Space 'Parcelamento ICMS' A_Space ref
    }

    RegExMatch(texto, '(\d{2}).{1,3}(\d{4})', &data, Instr(texto, 'Período'))
    RegExMatch(texto, '\d{9}\.\d{2}\-?\d{2}',  &insc, Instr(texto, 'IDENTIFICAÇÃO', 'Locale'))
    RegExMatch(texto, '\b\d{3,4}-\d\b', &cod, InStr(texto, 'Receita'))
    data := data[1] '-' data[2]
    cod := StrLen(cod[0]) = 5 ? '0' cod[0] : cod[0]
    cnpj := BuscarCNPJ(insc[0])
    return NomeNorm.Get(cnpj, cnpj) A_Space NomeDAE.get(cod, cod) A_Space data

    BuscarCNPJ(insc)
    {
        Loop read 'Conferencia_Empresas_NG.csv'
        {
            col := StrSplit(A_LoopReadLine, ';')
            if RegExReplace(insc, '\D') = RegExReplace(col[20], '\D')
                return RegExReplace(col[3], '\D')
        }
    }
}

ExtrairNomeDAEBahia()
{
    RegExMatch(texto, 'm)REFERÊNCIA\s{1,2}.*(\d{2})\/(\d{4})', &data)
    data := data[1] '-' data[2], cnpj := RegExCNPJ()
    RegExMatch(texto, '\d{4}', &cod, InStr(texto, 'CÓDIGO DA RECEITA'))
    return NomeNorm.Get(cnpj, cnpj) A_Space NomeDAE.get(cod[0], cod[0]) A_Space data
}

ExtrairLivroISSPrefeitura()
{
    pos := InStr(texto, 'Período'), RegExMatch(texto, '\d{2}\/\d{4}', &data, pos)
    cnpj := RegExCNPJ(pos)
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Livro ISS Prefeitura' A_Space StrReplace(data[0], '/', '-')
}

ExtrairLivroISS()
{
    cnpj := RegExCNPJ(), data := RegExData()
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Livro ISS' A_Space data
}

ExtrairLivroEntradas()
{
    cnpj := RegExCNPJ(), data := RegExData()
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Livro Reg Entradas' A_Space data
}

ExtrairLivroSaidas()
{
    cnpj := RegExCNPJ(), data := RegExData()
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Livro Reg Saidas' A_Space data
}

ExtrairLivroApuracaoICMS()
{
    cnpj := RegExCNPJ(), data := RegExData()
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Livro Reg ICMS' A_Space data
}

ExtrairLivroApuracaoIPI()
{
    cnpj := RegExCNPJ(), data := RegExData()
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Livro Reg IPI' A_Space data
}

ExtrairExtratoPGDAS()
{
    cnpj := RegExCNPJ(), RegExMatch(texto, 'Período de Apuração \(PA\): (\d{2}\/\d{4})', &data)
    return NomeNorm.Get(cnpj, cnpj) A_Space 'Extrato Simples Nacional' A_Space StrReplace(data[1], '/', '-')
}