/************************************************************************
 * @description Retorna um objeto com as principais propriedades da guia
 * @author Pedro Henrique C. Xavier
 * @date 2023-09-20
 * @version 2.0.9
 ***********************************************************************/

#Requires AutoHotkey v2.0
#Include %A_LineFile%\..\pdf2var.ahk

/**
 * Encontra os padrões das guias txt e retorna um objeto com as seguintes propriedades:
 * @param {string} caminho do arquivo pdf
 * @returns {object}
 * @prop {string} valorGuia valor da Guia
 * @prop {timestamp} competencia data da competencia em formato AAAAMM para facilitar a formatação com FormatTime
 * @prop {string} vencimento em formato DD/MM/AAAA com o vencimento da guia
 * @prop {string} cnpj o CNPJ. No caso de DAEs, é feito a correspondência com EmpresasNG.csv
 * @prop {string} nomeGuia o nome da guia de acordo com o dicionário
 */
ExtrairDadosGuias(filename)
{
    texto := pdf2var(filename)
    /** Dicionários/Tabelas com conversão de códigos em nomes
     ********************************************************/
    NomeDarf := Map('8109', 'PIS', '6912', 'PIS', '2172', 'COFINS', '5856', 'COFINS',
        '2089', 'IRPJ', '2372', 'CSLL', '1708', 'Ret IRRF', '5952', 'Ret PIS-COFINS-CSLL', '4805', 'Requisicao de Selos',
        '1082', 'DARF PREVIDENCIARIO', '1213', 'DARF PREVIDENCIARIO', '1099', 'DARF PREVIDENCIARIO')

    NomeDAE := Map('0215-4', 'ICMS sob Frete', '0209-7', 'ICMS-ST Entradas', '0211-3', 'ICMS-ST Bebidas',
        '0221-2', 'ICMS-ST Industria', '0326-9', 'Antecipação ICMS Comercio', '0327-7', 'Antecipação ICMS Industria',
        '0317-8', 'Diferenca de Aliquota ICMS', '0120-6', 'ICMS Normal', '0112-3', 'ICMS Normal', '0121-4', 'ICMS Normal',
        '0115-6', 'ICMS Normal', '0806', 'ICMS Normal', '0791', 'Diferença de Aliquota ICMS', '2036', 'FECP', '0313-7', 'ICMS-ST Rec Antecipado')

    NomeMes := Map('jan', '01', 'fev', '02', 'mar', '03', 'abr', '04', 'mai', '05', 'jun', '06',
        'jul', '07', 'ago', '08', 'set', '09', 'out', '10', 'nov', '11', 'dez', '12')

    pos := 1
    /** DARF gerado pelo NG Tributos e Sicalc Web sem código de barras */
    if InStr(texto, 'Documento de Arrecadação de Receitas Federais') and !(texto ~= '(\d{11}\s\d\s*){4}')
    {
        pos := RegExMatch(texto, '\d{2}\/\d{2}\/\d{4}', &dataApuracao, pos)
        pos := RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj, pos)
        pos := RegExMatch(texto, '\d{4}', &coddarf, pos + cnpj.len)
        pos := RegExMatch(texto, '\d{2}\/\d{2}\/\d{4}', &vencimento, pos)
        pos := RegExMatch(texto, 'i)10 VALOR TOTAL(\s*)((\d{1,3}\.)*\d{1,3}\,\d{2})', &valor)

        try
        {
            return {
                competencia: SubStr(dataApuracao[0], 7, 4) . SubStr(dataApuracao[0], 4, 2),
                vencimento: vencimento[0],
                cnpj: cnpj[0],
                valorGuia: valor[2],
                nomeGuia: NomeDarf.Get(coddarf[0], FindNomeDarf(coddarf[0]))
            }
        }
        catch
            return false
    }
    /** DARF gerado pelo Sicalc Web (com código de barras) */
    else if InStr(texto, 'Documento de Arrecadação de Receitas Federais') and (texto ~= '(\d{11}\s\d\s*){4}')
    {
        pos := RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj, pos)
        pos := RegExMatch(texto, '\d{2}\/\d{2}\/\d{4}', &dataApuracao, pos)
        pos := RegExMatch(texto, 'i)Pagar este documento até(\s*)(\d{2}\/\d{2}\/\d{4})', &vencimento)
        pos := RegExMatch(texto, 'i)Valor Total do Documento(\s*)((\d{1,3}\.)*\d{1,3}\,\d{2})', &valor)
        pos := RegExMatch(texto, 'i)Código.*(\d{4})', &coddarf)

        try
        {
            return {
                competencia: SubStr(dataApuracao[0], 7, 4) . SubStr(dataApuracao[0], 4, 2),
                vencimento: vencimento[2],
                cnpj: cnpj[0],
                valorGuia: valor[2],
                nomeGuia: NomeDarf.Get(coddarf[1], FindNomeDarf(coddarf[1]))
            }
        }
        catch
            return false
    }
    /** Guia do Simples Nacional */
    else if InStr(texto, 'Documento de Arrecadação do Simples Nacional') and !InStr(texto, 'Parcelamento')
    {
        pos := RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj)
        pos := RegExMatch(texto, 'i)Pagar este documento até(\s*)(\d{2}\/\d{2}\/\d{4})', &vencimento)
        pos := RegExMatch(texto, 'i)Valor Total do Documento(\s*)((\d{1,3}\.)*\d{1,3}\,\d{2})', &valor)
        pos := RegExMatch(texto, '([A-Za-z]{3}).*\/(\d{4})', &dataApuracao)

        try
        {
            return {
                competencia: pos ? dataApuracao[2] . NomeMes.Get(StrLower(dataApuracao[1])) : 'Diversos',
                vencimento: vencimento[2],
                cnpj: cnpj[0],
                valorGuia: valor[2],
                nomeGuia: 'Simples Nacional'
            }
        }
        catch
            return false
    }
    /** Guia de Parcelamento do Simples Nacional */
    else if InStr(texto, 'Documento de Arrecadação do Simples Nacional') and InStr(texto, 'Parcelamento')
    {
        pos := RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj, pos)
        pos := RegExMatch(texto, '\d{2}\/\d{2}\/\d{4}', &vencimento, pos)
        pos := RegExMatch(texto, 'i)Valor Total do Documento(\s*)((\d{1,3}\.)*\d{1,3}\,\d{2})', &valor)
        pos := RegExMatch(texto, 'i)Número\sda\sParcela\:\s\d+\/\d+', &parcela)

        try
        {
            return {
                competencia: StrLower(parcela[0]),
                vencimento: vencimento[0],
                cnpj: cnpj[0],
                valorGuia: valor[2],
                nomeGuia: 'Parcelamento Simples Nacional'
            }
        }
        catch
            return false
    }
    /** DAE emitido pelo DAPI */
    else if InStr(texto, 'DOCUMENTO DE ARRECADAÇÃO ESTADUAL') and InStr(texto, 'DAPI')
    {
        pos := RegExMatch(texto, '\d{2}\/\d{2}\/\d{4}', &vencimento, pos)
        pos := RegExMatch(texto, '\d{2}\/\d{2}\/\d{4}', &dataApuracao, pos + vencimento.len)
        pos := RegExMatch(texto, '\d{9}\.\d{2}(\-)*\d{2}', &InsEst, pos)
        pos := RegExMatch(texto, '\d{3}\-\d', &coddae, pos)
        pos := RegExMatch(texto, 'i)Valor Total(\s*)((\d{1,3}\.)*\d{1,3}\,\d{2})', &valor)

        try
        {
            return {
                competencia: SubStr(dataApuracao[0], 7, 4) . SubStr(dataApuracao[0], 4, 2),
                vencimento: vencimento[0],
                cnpj: FindIECNPJ(InsEst[0]),
                valorGuia: valor[2],
                nomeGuia: NomeDAE.Get('0' coddae[0], FindNomeDae('0' coddae[0]))
            }
        }
        catch
            return false
    }
    /** DAE emitido pelo Siare (exceto parcelamentos) */
    else if InStr(texto, 'DOCUMENTO DE ARRECADAÇÃO ESTADUAL') and !InStr(texto, 'DAPI') and !InStr(texto, 'PARCELAMENTO:')
    {
        pos := RegExMatch(texto, '\s(\d{4}\-\d)\s', &coddae, pos)
        pos := RegExMatch(texto, '\d{2}\/\d{2}\/\d{4}', &dataApuracao, pos)
        pos := RegExMatch(texto, 'i)Validade(\s*)(\d{2}\/\d{2}\/\d{4})', &vencimento)
        pos := RegExMatch(texto, 'i)R\$(\s*)((\d{1,3}\.)*\d{1,3}\,\d{2})', &valor)
        pos := RegExMatch(texto, '\d{9}\.\d{2}(\-)*\d{2}', &InsEst)

        try
        {
            return {
                competencia: SubStr(dataApuracao[0], 7, 4) . SubStr(dataApuracao[0], 4, 2),
                vencimento: vencimento[2],
                cnpj: FindIECNPJ(InsEst[0]),
                valorGuia: valor[2],
                nomeGuia: NomeDAE.Get(coddae[1], FindNomeDae(coddae[1]))
            }
        }
        catch
            return false
    }
    /** DAE emitido pelo Siare (Parcelamentos) */
    else if InStr(texto, 'DOCUMENTO DE ARRECADAÇÃO ESTADUAL') and !InStr(texto, 'DAPI') and InStr(texto, 'PARCELAMENTO:')
    {
        pos := RegExMatch(texto, 'i)PARCELA\:\s\d{3}\/\d{3}', &dataApuracao)
        pos := RegExMatch(texto, 'i)Vencimento\s(\d{2}\.\d{2}\.\d{4})', &vencimento)
        pos := RegExMatch(texto, 'i)TOTAL(\s*)R\$(\s*)((\d{1,3}\.)*\d{1,3}\,\d{2})', &valor)
        pos := RegExMatch(texto, '\d{9}\.\d{2}(\-)*\d{2}', &InsEst)

        try
        {
            return {
                competencia: StrTitle(dataApuracao[0]),
                vencimento: StrReplace(vencimento[1], '.', '/'),
                cnpj: FindIECNPJ(InsEst[0]),
                valorGuia: valor[3],
                nomeGuia: 'Parcelamento ICMS'
            }
        }
        catch
            return false
    }
    /** DAE emitido pela Bahia */
    else if InStr(texto, 'Nº DE SÉRIE / NOSSO NÚMERO')
    {
        pos := RegExMatch(texto, 'CÓDIGO DA RECEITA\s+(\d{4})', &coddae, pos)
        pos := RegExMatch(texto, '\d{2}\/\d{2}\/\d{4}', &vencimento, pos)
        pos := RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj, pos)
        pos := RegExMatch(texto, 'REFERÊNCIA\s+(\d{2})\/(\d{4})', &dataApuracao, pos)
        pos := RegExMatch(texto, 'TOTAL A RECOLHER(\s*)R\$\s((\d{1,3}\.)*\d{1,3}\,\d{2})', &valor)

        try
        {
            return {
                competencia: dataApuracao[2] . dataApuracao[1],
                vencimento: vencimento[0],
                cnpj: cnpj[0],
                valorGuia: valor[2],
                nomeGuia: NomeDAE.Get(coddae[1], '')
            }
        }
        catch
            return false
    }

    /** ISSQN emitido pela Prefeitura de Maravilhas */
    else if InStr(texto, 'PREFEITURA MUNCIPAL DE MARAVILHAS Imposto Sobre Serviços de Qualquer Natureza')
    {
        pos := RegExMatch(texto, 'CNPJ\:\s(\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})', &cnpj)
        pos := RegExMatch(texto, 'i)Vencimento\s(\d{2}\/\d{2}\/\d{4})', &vencimento)
        pos := RegExMatch(texto, 'i)REFERÊNCIA\:\s(\d{2})\/(\d{4})', &dataApuracao)
        pos := RegExMatch(texto, 'i)VALOR\sCOBRADO\:\s((\d{1,3}\.)*\d{1,3}\,\d{2})', &valor)

        try
        {
            return {
                competencia: dataApuracao[2] . dataApuracao[1],
                vencimento: vencimento[1],
                cnpj: cnpj[1],
                valorGuia: valor[1],
                nomeGuia: 'ISSQN Maravilhas'
            }
        }
        catch
            return false
    }

    /** ISSQN emitido pela Prefeitura de Papagaios */
    else if InStr(texto, 'PREFEITURA MUNCIPAL DE PAPAGAIOS Imposto Sobre Serviços de Qualquer Natureza')
    {
        pos := RegExMatch(texto, 'CNPJ\:\s(\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2})', &cnpj)
        pos := RegExMatch(texto, 'i)Vencimento\s(\d{2}\/\d{2}\/\d{4})', &vencimento)
        pos := RegExMatch(texto, 'i)REFERÊNCIA\:\s(\d{2})\/(\d{4})', &dataApuracao)
        pos := RegExMatch(texto, 'i)VALOR\sCOBRADO\:\s((\d{1,3}\.)*\d{1,3}\,\d{2})', &valor)

        try
        {
            return {
                competencia: dataApuracao[2] . dataApuracao[1],
                vencimento: vencimento[1],
                cnpj: cnpj[1],
                valorGuia: valor[1],
                nomeGuia: 'ISSQN Papagaios'
            }
        }
        catch
            return false
    }
    else
        return false
}

/**
 * Lê o arquivo em busca da descrição do código
 * Essa função só será acionada caso não seja encontrado código no Map NomeDae
 * @param {string} coddae
 * @returns {string}
 */
FindNomeDae(coddae)
{
    Loop read A_LineFile '\..\..\data\tab_cod_dae.txt'
        if SubStr(A_LoopReadLine, 1, 6) = coddae
            return SubStr(A_LoopReadLine, 8)
}

/**
 * Lê o arquivo em busca da descrição do código
 * Essa função só será acionada caso não seja encontrado código no Map NomeDarf
 * @param {string} coddarf
 * @returns {string}
 */
FindNomeDarf(coddarf)
{
    Loop read A_LineFile '\..\..\data\tab_cod_darf.txt'
        if SubStr(A_LoopReadLine, 1, 4) = coddarf
            return SubStr(A_LoopReadLine, 6)
}

/**
 * Lê o arquivo em busca da descrição da Inscrição Estdual e retorna o CNPJ
 * @param {string} strIE
 * @returns {any}
 */
FindIECNPJ(strIE)
{
    strIE := RegExReplace(strIE, '\D')
    Loop read A_LineFile '\..\..\data\EmpresasNG.csv'
    {
        str := RegexReplace(A_LoopReadLine, '\D')
        if InStr(str, strIE)
            return StrSplit(A_LoopReadLine, ';')[3]
    }
}
