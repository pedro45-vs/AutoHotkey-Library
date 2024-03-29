﻿/************************************************************************
 * @description Retorna um MAP com as informações relevantes do arquivo
 * para ser usado por outros scripts
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/24
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0

/**
 * Retorna um Map com as informações relevantes das empresas do NG
 * @returns {map} cnpj
 * @prop nome
 * @prop insc
 * @prop regime
 * @prop status
 */
ConferenciaEmpresaNG()
{
    dados := Map()
    FileEncoding('CP0')
    Loop read A_LineFile '\..\..\data\Conferencia_Empresas_NG.csv'
    {
        col := StrSplit(A_LoopReadLine, ';')
        if A_Index = 1
        {
            cab := col.Length
            continue
        }
        if col.Length > cab
            throw ValueError('quantidade de colunas inválida', -1, col[5])

        cnpj := RegExReplace(col[3], '\D'), insc := RegExReplace(col[20], '[\.\-]')
        dados[cnpj] := { nome: trim(col[5]), insc: insc, regime: col[46], status: col[1] }
        if dados[cnpj].status = 'Ativa' and not dados[cnpj].regime
        {
            if MsgBox('A empresa ' dados[cnpj].nome ' não tem regime definido`nDeseja continuar?.', 'ConferenciaEmpresaNG()', 'OC Icon!') = 'OK'
                continue
            else
                Exit()                
        }
    }
    ; Insere os dados da Associação para poder sair impresso na lista mesmo não estando cadastrado no NG
    dados['02950258000184'] := { nome: 'ASSOCIACAO COMUNITARIA BENEFICENTE E CULTURAL DE PAPAGAIOS', insc: 'ISENTO', regime: 'Lucro Presumido', status: 'Ativa' }
    return dados
}
