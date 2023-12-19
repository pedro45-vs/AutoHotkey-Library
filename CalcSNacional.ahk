/************************************************************************
 * @description Class para calcular as alíquotas do Simples Nacional
 * @file CalcSNacional.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/09/04
 * @version 2.0.7
 ***********************************************************************/

/**
 * Classe para cálculo das alíquotas do Simples Nacional.
 *
 * Permite definir exclusões de tributos para cálculo do ST, monofásico e retenções.
 *
 * Compatível até a 5ª faixa de tributação (a 6ª faixa requer regras especiais)
 *
 * @prop {number} aliqEfetiva retorna a alíquota devida
 * @prop {number} pis retorna a parcela correspondente ao PIS
 * @prop {number} cofins retorna a parcela correspondente ao COFINS
 * @prop {number} irpj retorna a parcela correspondente ao IRPJ
 * @prop {number} csll retorna a parcela correspondente ao CSLL
 * @prop {number} inss retorna a parcela correspondente ao INSS
 * @prop {number} ipi retorna a parcela correspondente ao IPI
 * @prop {number} icms retorna a parcela correspondente ao ICMS
 * @prop {number} iss retorna a parcela correspondente ao ISS
 */
class CalcSNacional
{
    /**
     * Inicia a classe e define as propriedades de acordo com a alíquota
     * @param {string} anexo O anexo a ser considerado. Opções válidas são: I, II, III, IV e V
     * @param {number} RB12 A Receita Bruta acumulada dos 12 últimos meses
     */
    __New(anexo, RB12 := 1)
    {
        this.anexo := anexo
        this.RB12 := RB12
        this.faixa := this.__definirFaixa(RB12)
        this.__definirImpostos()
    }
    /**
     * Enumerador personalizado com as propriedades disponíveis
     * @param prop
     * @param value Opcional
     * @returns {$enumerator}
     */
    __Enum(prop, value?)
    {
        num := 0, arr := ['pis', 'cofins', 'irpj', 'csll', 'inss', 'ipi', 'icms', 'iss']
        Enumerator(&prop, &value?)
        {
            if num >= 8
                return false

            num++, prop := arr[num], value := this.%prop%
            return true
        }
        return Enumerator
    }
    anexos := Map(
        'I',[
        { aliq: 0.04,  ded:     0,  pis: 0.0276, cofins: 0.1274, irpj: 0.0550, csll: 0.0350, inss: 0.4150, ipi: 0.0000, icms: 0.3400, iss: 0.0000 },
        { aliq: 0.073, ded:  5940,  pis: 0.0276, cofins: 0.1274, irpj: 0.0550, csll: 0.0350, inss: 0.4150, ipi: 0.0000, icms: 0.3400, iss: 0.0000 },
        { aliq: 0.095, ded: 13860,  pis: 0.0276, cofins: 0.1274, irpj: 0.0550, csll: 0.0350, inss: 0.4200, ipi: 0.0000, icms: 0.3350, iss: 0.0000 },
        { aliq: 0.107, ded: 22500,  pis: 0.0276, cofins: 0.1274, irpj: 0.0550, csll: 0.0350, inss: 0.4200, ipi: 0.0000, icms: 0.3350, iss: 0.0000 },
        { aliq: 0.143, ded: 87300,  pis: 0.0276, cofins: 0.1274, irpj: 0.0550, csll: 0.0350, inss: 0.4200, ipi: 0.0000, icms: 0.3350, iss: 0.0000 }
    ], 'II', [
        { aliq: 0.045, ded:     0,  pis: 0.0249, cofins: 0.1151, irpj: 0.0550, csll: 0.0350, inss: 0.3750, ipi: 0.0750, icms: 0.3200, iss: 0.0000 },
        { aliq: 0.078, ded:  5940,  pis: 0.0249, cofins: 0.1151, irpj: 0.0550, csll: 0.0350, inss: 0.3750, ipi: 0.0750, icms: 0.3200, iss: 0.0000 },
        { aliq: 0.100, ded: 13860,  pis: 0.0249, cofins: 0.1151, irpj: 0.0550, csll: 0.0350, inss: 0.3750, ipi: 0.0750, icms: 0.3200, iss: 0.0000 },
        { aliq: 0.112, ded: 22500,  pis: 0.0249, cofins: 0.1151, irpj: 0.0550, csll: 0.0350, inss: 0.3750, ipi: 0.0750, icms: 0.3200, iss: 0.0000 },
        { aliq: 0.147, ded: 85500,  pis: 0.0249, cofins: 0.1151, irpj: 0.0550, csll: 0.0350, inss: 0.3750, ipi: 0.0750, icms: 0.3200, iss: 0.0000 }
    ], 'III', [
        { aliq: 0.060, ded:      0, pis: 0.0278, cofins: 0.1282, irpj: 0.0400, csll: 0.0350, inss: 0.4340, ipi: 0.0000, icms: 0.0000, iss: 0.3350 },
        { aliq: 0.112, ded:   9360, pis: 0.0305, cofins: 0.1405, irpj: 0.0400, csll: 0.0350, inss: 0.4340, ipi: 0.0000, icms: 0.0000, iss: 0.3200 },
        { aliq: 0.135, ded:  17640, pis: 0.0296, cofins: 0.1364, irpj: 0.0400, csll: 0.0350, inss: 0.4340, ipi: 0.0000, icms: 0.0000, iss: 0.3250 },
        { aliq: 0.160, ded:  35640, pis: 0.0296, cofins: 0.1364, irpj: 0.0400, csll: 0.0350, inss: 0.4340, ipi: 0.0000, icms: 0.0000, iss: 0.3250 },
        { aliq: 0.210, ded: 125640, pis: 0.0278, cofins: 0.1282, irpj: 0.0400, csll: 0.0350, inss: 0.4340, ipi: 0.0000, icms: 0.0000, iss: 0.3350 }
    ], 'IV', [
        { aliq: 0.045, ded:      0, pis: 0.0383, cofins: 0.1767, irpj: 0.1880, csll: 0.1520, inss: 0.0000, ipi: 0.0000, icms: 0.0000, iss: 0.4450 },
        { aliq: 0.090, ded:   8100, pis: 0.0445, cofins: 0.2055, irpj: 0.1980, csll: 0.1520, inss: 0.0000, ipi: 0.0000, icms: 0.0000, iss: 0.4000 },
        { aliq: 0.102, ded:  12420, pis: 0.0427, cofins: 0.1973, irpj: 0.2080, csll: 0.1520, inss: 0.0000, ipi: 0.0000, icms: 0.0000, iss: 0.4000 },
        { aliq: 0.140, ded:  39780, pis: 0.0410, cofins: 0.1890, irpj: 0.1780, csll: 0.1920, inss: 0.0000, ipi: 0.0000, icms: 0.0000, iss: 0.4000 },
        { aliq: 0.220, ded: 183780, pis: 0.0392, cofins: 0.1808, irpj: 0.1880, csll: 0.1920, inss: 0.0000, ipi: 0.0000, icms: 0.0000, iss: 0.4000 }
    ], 'V', [
        { aliq: 0.155, ded:     0,  pis: 0.0305, cofins: 0.1410, irpj: 0.2500, csll: 0.1500, inss: 0.2885, ipi: 0.0000, icms: 0.0000, iss: 0.1400 },
        { aliq: 0.180, ded:  4500,  pis: 0.0305, cofins: 0.1410, irpj: 0.2300, csll: 0.1500, inss: 0.2785, ipi: 0.0000, icms: 0.0000, iss: 0.1700 },
        { aliq: 0.195, ded:  9900,  pis: 0.0323, cofins: 0.1492, irpj: 0.2400, csll: 0.1500, inss: 0.2385, ipi: 0.0000, icms: 0.0000, iss: 0.1900 },
        { aliq: 0.205, ded: 17100,  pis: 0.0341, cofins: 0.1574, irpj: 0.2100, csll: 0.1500, inss: 0.2385, ipi: 0.0000, icms: 0.0000, iss: 0.2100 },
        { aliq: 0.230, ded: 62100,  pis: 0.0305, cofins: 0.1410, irpj: 0.2300, csll: 0.1250, inss: 0.2385, ipi: 0.0000, icms: 0.0000, iss: 0.2350 }
    ] )

    faixa := [ { FatMin:          0, FatMax:  180000 },
               { FatMin:  180000.01, FatMax:  360000 },
               { FatMin:  360000.01, FatMax:  720000 },
               { FatMin:  720000.01, FatMax: 1800000 },
               { FatMin: 1800000.01, FatMax: 3600000 } ]

    /**
     * Permite definir quais impostos serão excluídos do cálculo e atualiza as propriedades
     *
     * Por exemplo, em vendas com ST, o ICMS deve ser excluído.
     * Já para saídas monofásicas, o PIS e COFINS deve ser excluído.
     * @param {variadic} params Os impostos a serem excluídos
     */
    definirExclusoes(params*)
    {
        f := this.faixa, a := this.anexo, exclusoes := 0
        for prop in params
        {
            exclusoes += this.anexos[a][f].%prop%
            this.%prop% := 0
        }
        this.aliqEfetiva *= (1 - exclusoes)
    }
    /**
     * Método interno. Não deve ser chamado diretamente
     * Define as propriedades de cada imposto de acordo com a alíquota devida
     */
    __definirImpostos()
    {
        f := this.faixa, a := this.anexo
        this.aliqEfetiva := (this.RB12 * this.anexos[a][f].aliq - this.anexos[a][f].ded) / this.RB12

        for prop in ['pis', 'cofins', 'irpj', 'csll', 'inss', 'ipi', 'icms', 'iss']
            this.DefineProp(prop, {value: this.anexos[a][f].%prop% * this.aliqEfetiva})
    }
    /**
     * Método interno. Não deve ser chamado diretamente
     * Define qual é a faixa de impostos e deduções que será aplicada
     */
    __definirFaixa(RB12)
    {
        for i, v in this.faixa
        {
            if (RB12 >= this.faixa[A_Index].FatMin and RB12 <= this.faixa[A_Index].FatMax)
            {
                ind := A_Index
                break
            }
        }
        if not IsSet(ind)
            throw ValueError('Receita Bruta fora da faixa de limite', -2)

        return ind
    }
}