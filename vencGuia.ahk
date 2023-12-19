/************************************************************************
 * @description Calcula o vencimento de guias levando em conta os feriados
 * @file vencGuia.ahk
 * @author Pedro Henrique C Xavier
 * @date 2023/08/21
 * @version 2.0.5
 ***********************************************************************/

#Include %A_LineFile%\..\dataLib.ahk

/**
 * Calcula o vencimento de guias levando em conta os feriados
 * No ultimo dia útil do ano não tem expediente bancário e não considerei como dia útil
 * @param {string} guia Nome da Guia
 * @param {string} formato Formato de saída
 * @returns {string} 
 */
VencGuia(guia, formato := 'yyyyMMdd')
{
    if guia = 'Simples Nacional' or guia = 'DAS'
        venc := postergaData(A_YYYY A_MM '20')
        
    else if guia = 'ISSQN'
        venc := postergaData(A_YYYY A_MM '15')
        
    else if guia = 'ICMS Normal'
        venc := postergaData(A_YYYY A_MM '08')
        
    else if guia = 'DIFAL' or guia = 'Diferença de Alíquota ICMS' or guia = 'ICMS-ST Entradas' or guia = 'ICMS S/Frete'
        venc := postergaData(mesFuturo('yyyyMM') . '02')
        
    else if guia = 'Antecipação de ICMS'
        venc := postergaData(mesFuturo('yyyyMM') . '20')
        
    else if guia = 'PIS' or guia = 'COFINS'
        venc := antecipaData(A_YYYY A_MM '25')
        
    else if guia = 'RET. IRRF' or guia = 'RET. PIS/COFINS/CSLL' or guia = 'GPS Produtor Rural'
        venc := antecipaData(A_YYYY A_MM '20')
        
    else if guia = 'FGTS'
        venc := antecipaData(A_YYYY A_MM '07')
        
    else if guia = 'IRPJ' or guia = 'CSLL' or guia = 'Parcelamento ICMS'
    {
        venc := antecipaData(ultimoDiaMes())
        if SubStr(venc, 5, 2) = '12'
            venc := DateAdd(venc, -1, 'days')
    }
    return FormatTime(venc, formato)
}
