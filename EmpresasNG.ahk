/************************************************************************
 * @description Retorna um MAP com as informações relevantes do arquivo
 * para ser usado por outros scripts
 * @author Pedro Henrique C. Xavier
 * @date 2024-03-26
 * @version 2.1-alpha.9
 ***********************************************************************/

#Requires AutoHotkey v2.0

/**
 * Retorna um Map com as informações relevantes das empresas do NG
 * @returns {map} cnpj
 * @prop nome
 * @prop insc
 * @prop regime
 * @prop uf
 * @prop grupo
 */
EmpresasNG()
{
    dados := Map()
    FileEncoding('CP0')
    Loop read A_LineFile '\..\..\data\EmpresasNG.csv'
    {
        col := StrSplit(A_LoopReadLine, ';')
        if A_Index = 1
        {
            cab := col.Length
            continue
        }
        if col.Length > cab
            throw ValueError('quantidade de colunas inválida', -1, col[5])

        if col[1] = 'Inativa'
            continue

        cnpj := RegExReplace(col[3], '\D'), insc := RegExReplace(col[20], '[\.\-]')
        grupo := col[50] ? StrSplit(col[50], A_Space)[-1] : ''

        dados[cnpj] := { nome: trim(col[5]), insc: insc, regime: col[46], uf: col[39], grupo: grupo }

        if not dados[cnpj].regime
        {
            if MsgBox('A empresa ' dados[cnpj].nome ' não tem regime definido`nDeseja continuar?.', 'EmpresasNG()', 'OC Icon!') = 'OK'
                continue
            else
                Exit()
        }
        if not dados[cnpj].grupo
        {
            if MsgBox('A empresa ' dados[cnpj].nome ' não tem grupo definido`nDeseja continuar?.', 'EmpresasNG()', 'OC Icon!') = 'OK'
                continue
            else
                Exit()
        }
    }
    ; Insere os dados da Associação para poder sair impresso na lista mesmo não estando cadastrado no NG
    dados['02950258000184'] := { nome: 'ASSOCIACAO COMUNITARIA BENEFICENTE E CULTURAL DE PAPAGAIOS', insc: 'ISENTO', regime: 'Lucro Presumido', uf: 'MG', grupo: 'Serviços' }
    return dados
}

AdicionarInativas(EmpNG)
{
    FileEncoding('CP0')
    Loop read A_LineFile '\..\..\data\EmpresasNG.csv'
    {
        col := StrSplit(A_LoopReadLine, ';')
        if A_Index > 1 and col[1] = 'Inativa'
        {
            cnpj := RegExReplace(col[3], '\D')
            EmpNG[cnpj] := { nome: trim(col[5]), insc: false, regime: false, uf: false, grupo: false }
        }
    }
}
