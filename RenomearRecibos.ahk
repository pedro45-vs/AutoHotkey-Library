/************************************************************************
 * @description Gera o nome do arquivo com base nas informações do recibo
 * @author Pedro Henrique C. Xavier
 * @date 2023/12/22
 * @version 2.0.10
 ***********************************************************************/

#Requires AutoHotkey v2.0
#Include %A_LineFile%\..\pdf2var.ahk

/**
 * Função para Renomear os Recibos DeSTDA
 * @param FilePath Arquivo a ser analisado
 * @returns {String} Nome do arquivo com base no conteúdo
 */
RenomearDeSTDA(FilePath)
{
    NomeMes := Map('jan','01','fev','02','mar','03','abr','04','mai','05','jun','06',
				   'jul','07','ago','08','set','09','out','10','nov','11','dez','12')

    texto := pdf2var(FilePath)

    try
    {
        RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj)
        RegExMatch(texto, '([a-z]{3})\/(\d{4})', &data)
        RegExMatch(texto, 'trm (\d{2})\/(\d{2})\/(\d{4}) \* (\d{2})\:(\d{2})\:(\d{2})', &tram)

        cnpj := RegExReplace(cnpj[0], '\D')
        tram := Format('{4:}{3:}{2:}{5:}{6:}{7:}', tram*)
        data := data[2] '-' NomeMes[data[1]]
    }
    catch
        return false

    return 'Rec_DeSTDA_' cnpj '_' data '_' tram '.pdf'
}

/**
 * Função para Renomear os Recibos Sintegra
 * @param FilePath Arquivo a ser analisado
 * @returns {String} Nome do arquivo com base no conteúdo
 */
RenomearSintegra(FilePath)
{
    texto := pdf2var(FilePath)

    try
    {
        RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj)
        RegExMatch(texto, '\d\d\/(\d\d)\/(\d\d\d\d)', &data)
        RegExMatch(texto, '(\d{2})\/(\d{2})\/(\d{4}) às (\d{2})\:(\d{2})\:(\d{2})', &tram)

        cnpj := RegExReplace(cnpj[0], '\D')
        tram := Format('{4:}{3:}{2:}{5:}{6:}{7:}', tram*)
        data := data[2] '-' data[1]
    }
    catch
        return false

    return 'Rec_Sintegra_' cnpj '_' data '_' tram '.pdf'
}

/**
 * Função para Renomear os Recibos DAPI
 * @param FilePath Arquivo a ser analisado
 * @returns {String} Nome do arquivo com base no conteúdo
 */
RenomearDAPI(FilePath)
{
    texto := pdf2var(FilePath)

    try
    {
        RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj)
        RegExMatch(texto, 'REFERÊNCIA: 01 a \d{2}\/(\d{2})\/(\d{4})', &data)
        RegExMatch(texto, 'TRANSMISSÃO: (\d{2})\/(\d{2})\/(\d{4})', &trad)
        RegExMatch(texto, 'HORA: \d\d:\d\d:\d\d', &trah)

        cnpj := RegExReplace(cnpj[0], '\D')
        data := data[2] '-' data[1]
        trad := Format('{4:}{3:}{2:}', trad*)
        trah := RegExReplace(trah[0], '\D')
        tram := trad . trah
    }
    catch
        return false

    return 'Rec_DAPI_' cnpj '_' data '_' tram '.pdf'
}

/**
 * Função para Renomear os Recibos Sped
 * @param FilePath Arquivo a ser analisado
 * @returns {String} Nome do arquivo com base no conteúdo
 */
RenomearReciboSped(FilePath)
{
    texto := pdf2var(FilePath)
    try
    {
        RegExMatch(texto, '\d{2}\.\d{3}\.\d{3}\/\d{4}\-\d{2}', &cnpj)
        RegExMatch(texto, '\d{2}\/(\d{2})\/(\d{4})', &data)
        RegExMatch(texto, '(\d{2})\/(\d{2})\/(\d{4}) às (\d{2})\:(\d{2})\:(\d{2})', &tram)

        cnpj := RegExReplace(cnpj[0], '\D')
        data := data[2] '-' data[1]
        tram := Format('{4:}{3:}{2:}{5:}{6:}{7:}', tram*)
    }
    catch
        return false

    if InStr(texto, 'Versão Sped Fiscal:')
        filename := 'Rec_Sped_Fiscal_' cnpj '_' data '_' tram '.pdf'
    
    else if InStr(texto, 'Versão EFD-Contribuições:')
        filename := 'Rec_EFD_Contribuicoes_' cnpj '_' data '_' tram '.pdf'
    
    else 
        filename := false
    
    return filename
}