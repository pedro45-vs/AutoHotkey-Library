/************************************************************************
 * @description 
 * @author Pedro Henrique C. Xavier
 * @date 2024-03-28
 * @version 2.1-alpha.9
 ***********************************************************************/

#Requires AutoHotkey v2.0

PessoasNG()
{
    Pessoas := Map('codigo', [], 'nome', [], 'cnpj', [], 'insc', [])
    Loop read A_LineFile '\..\..\data\PessoasNG.psv'
    {
        col := StrSplit(A_LoopReadLine, '|')
        Pessoas['codigo'].push(col[1])
        Pessoas['nome'].push(col[2])
        Pessoas['cnpj'].push(col[3])
        Pessoas['insc'].push(col[4])
    }

    Pessoas.DefineProp('Search', {call: Search})
    
    return Pessoas

    Search(mapObj, campo, pesquisa)
    {
        for index, value in mapObj[campo]
        {
            if value = pesquisa
                return {
                    codigo: mapObj['codigo'][index], 
                    nome: mapObj['nome'][index],
                    cnpj: mapObj['cnpj'][index],
                    insc: mapObj['insc'][index],
                }
        }
        return {codigo: false, nome: false, cnpj: false, insc: false}
    }
}