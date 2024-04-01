/************************************************************************
 * @description Extrai as principais informações do Extrato do Simples Nacional
 * @author Pedro Henrique C. Xavier
 * @date 2023-09-20
 * @version 2.0.9
 ***********************************************************************/

#Requires AutoHotkey v2
#Include %A_LineFile%\..\pdf2var.ahk

/**
 * Extrai as principais informações do Extrato do Simples Nacional
 * @param {string} caminho do arquivo pdf
 * @returns {object} Objeto com as seguintes propriedades
 * @prop {string} cnpj CNPJ da empresa
 * @prop {string} apuracao Período de apuração
 * @prop {array} array com os faturamentos anteriores (mês/ano - faturamento)
 */
ExtrairDadosExtrato(fileName)
{
    texto := pdf2var(fileName, '-raw')
    if not InStr(texto, 'Extrato do Simples Nacional')
        return false

    RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj)
    RegExMatch(texto, 'Período de Apuração \(PA\): (\d{2}\/\d{4})', &comp)
    
    /**
     * Marca a parada do Regex
     * só irá pegar o faturamento do mercado interno por enquanto
     */
    faturamentos := [] 
    mercado_externo := InStr(texto, '2.2.2) Mercado Externo')
    while pos := RegExMatch(texto, '(\d{2}\/\d{4})\s(((\.)?\d{1,3})+\,\d{2})', &fat, pos ?? 1)
    {
        if pos > mercado_externo
            break
        faturamentos.push( [ fat[1], fat[2] ] )
        pos += fat.len
    }
    return {
        cnpj: RegExReplace(cnpj[0], '\D'),
        apuracao: comp[1],
        faturamentos: faturamentos
    }
}

