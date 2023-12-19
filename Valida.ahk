/************************************************************************
 * @description Biblioteca com funções diversas para validação de CPF,
 * CNPJ, Inscrição estadual e chave de acesso NFe
 * @file Valida.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/08/22
 * @version 2.0.5
 ***********************************************************************/

#Requires AutoHotkey v2.0

/**
 * Valida CNPJ e opcionalmente retorna reformatado 
 * @param {String} CNPJ CNPJ a ser validado 
 * @param {&VarRef} pad Variável de retorno com padding de zeros
 * @param {&VarRef} form Variável de retorno com CNPJ formatado 
 * @returns {Boolean} verdadeiro para CNPJ válido
 */
validaCNPJ(CNPJ, &pad:='', &form:='')
{
    CNPJ := RegExReplace(CNPJ, '\D'), pad := Format('{:014}', CNPJ)
    form := SubStr(pad, 1, 2) '.' SubStr(pad, 3, 3) '.'
    . SubStr(pad, 6, 3) '/' SubStr(pad, 9, 4) '-' SubStr(pad, 13, 2)
    (pad = '00000000000000') && Exit(0)

    return validaMod11(SubStr(pad, 1, 13)) and validaMod11(SubStr(pad, 1, 14)) ? true : false
}

/**
 * Valida CPF e opcionalmente retorna reformatado 
 * @param {String} CPF CPF a ser validado 
 * @param {&VarRef} pad Variável de retorno com padding de zeros
 * @param {&VarRef} form Variável de retorno com CPF formatado 
 * @returns {Boolean} verdadeiro para CPF válido
 */
validaCPF(CPF, &pad:='', &form:='')
{
    CPF := RegExReplace(CPF, '\D'), pad := Format('{:011}', CPF)
    form := SubStr(pad, 1, 3) '.' SubStr(pad, 4, 3) '.' SubStr(pad, 7, 3) '-' SubStr(pad, 10, 2)

    if CPF = '00000000000' or CPF = '11111111111' or CPF = '22222222222' or CPF = '33333333333'
    or CPF = '44444444444' or CPF = '55555555555' or CPF = '66666666666'
    or CPF = '77777777777' or CPF = '88888888888' or CPF = '99999999999'
        return false

    soma:=0
    For n in [10,9,8,7,6,5,4,3,2]
        soma += n * SubStr(pad, A_Index, 1)

    dv := 11 - Mod(soma, 11) > 9 ? 0 : 11 - Mod(soma, 11)
    if dv != SubStr(pad, 10, 1)
        return false

    soma:=0
    For n in [11,10,9,8,7,6,5,4,3,2]
        soma += n * SubStr(pad, A_Index, 1)

    dv := 11 - Mod(soma, 11) > 9  ? 0 : 11 - Mod(soma, 11)
    if dv != SubStr(pad, 11, 1)
        return false

    return true
}

/**
 * Valida uma Inscrição Estadual com base nas regras extraidas do site Sintegra.
 * AM, CE, ES, PB, PI, SC e SE usam o mesmo cálculo do dígito verificador.
 * AC e DF também usam a mesma fórmula.
 * Obs.: A fórmula do Tocantins está errada no site do Sintegra.
 * @param {String} IE String da inscrição estadual. Não importa se está com ou sem pontuação
 * @param {String} UF Abreviação da UF ou código do IBGE para validação
 * @param {&VarRef} IEpad String formatada com o padding dos zeros à esquerda
 * @returns {Boolean} 
 */
validaIE(IE, UF, &IEpad:='')
{
    IE := RegExReplace(IE, '\D')

    if UF = 'AM' or UF = 13 or UF = 'CE' or UF = 23 or UF = 'ES' or UF = 32 or UF = 'PB' or UF = 25
    or UF = 'PI' or UF = 22 or UF = 'SC' or UF = 42 or UF = 'SE' or UF = 28
    {
		if StrLen(IE) > 9 or !validaMod11(IE)
			return false

        IEpad := Format('{:09}', IE)
		return true
    }
    if UF = 'AC' or UF = 12 or UF = 'DF' or UF = 53
    {
        if StrLen(IE) > 13 or !validaMod11(SubStr(IE, 1, 12)) or !validaMod11(SubStr(IE, 1, 13))
            return false

        IEpad := Format('{:013}', IE)
        return true
    }
    if UF = 'MT' or UF = 51
    {
		if StrLen(IE) > 11 or !validaMod11(IE)
			return false

        IEpad := Format('{:011}', IE)
        return true
    }
    if UF = 'RS' or UF = 43
    {
		if StrLen(IE) > 10 or !validaMod11(IE)
			return false

        IEpad := Format('{:010}', IE)
        return true
    }
    if UF = 'RO' or UF = 11
    {
		if StrLen(IE) > 14 or !validaMod11(IE)
			return false

        IEpad := Format('{:014}', IE)
        return true
    }
    if UF = 'MA' or UF = 21
    {
        if SubStr(IE, 1, 2) != 12 or StrLen(IE) != 9 or !validaMod11(IE)
            return false

        IEpad := IE
        return true
    }
    if UF = 'MS' or UF = 50
    {
        if SubStr(IE, 1, 2) != 28 or StrLen(IE) != 9 or !validaMod11(IE)
            return false

        IEpad := IE
        return true
    }
    if UF = 'PA' or UF = 15
    {
        if SubStr(IE, 1, 2) != 15 or StrLen(IE) != 9 or !validaMod11(IE)
            return false

        IEpad := IE
        return true
    }
    if UF = 'TO' or UF = 17
    {
        if StrLen(IE) != 9 or SubStr(IE, 1, 2) != 29 or !validaMod11(IE)
            return false

        IEpad := IE
        return true
    }
    if UF = 'GO' or UF = 52
    {
        if StrLen(IE) != 9 or !validaMod11(IE)
            return false

        d := SubStr(IE, 1, 2)
        if d != 10 and d != 11 and d != 20 and d != 29
            return false

        IEpad := IE
        return true
    }
    if UF = 'MG' or UF = 31
    {
        IE := IEpad := Format('{:013}', IE)
		if StrLen(IE) != 13
			return false

        strTrb := SubStr(IE, 1, 3) '0' SubStr(IE, 4, 8)
        soma := 0
        Loop 12
            somaAlgarismos .= SubStr(strTrb, A_index, 1) * (Mod(A_Index, 2) ? 1 : 2)

        Loop StrLen(somaAlgarismos)
            soma += SubStr(somaAlgarismos, A_index, 1)

        sub := (SubStr(soma, 1, -1) . 0) + 10 - soma
        dv := sub = 10 ? 0 : sub
        if dv != SubStr(IE, 12, 1)
            return false

        soma := 0
        For n in [3,2,11,10,9,8,7,6,5,4,3,2]
            soma += n * SubStr(IE, A_Index, 1)

        dv := Mod(soma, 11) < 2 ? 0 : 11 - Mod(soma, 11)
        if dv != SubStr(IE, 13, 1)
            return false

        return true
    }
    if UF = 'AL' or UF = 27
    {
        if StrLen(IE) != 9 or SubStr(IE, 1, 2) != 24
            return false

        soma := 0
        For n in [9,8,7,6,5,4,3,2]
            soma += SubStr(IE, A_Index, 1) * n

        produto := soma * 10
        resto := produto - Integer(produto / 11) * 11
        dv := resto = 10 ? 0 : resto

        if dv != SubStr(IE, 9, 1)
            return false

        IEpad := IE
        return true
    }
    if UF = 'PR' or UF = 41
    {
        IE := IEpad := Format('{:010}', IE)
		if StrLen(IE) != 10
			return false

        soma := 0
        For n in [3,2,7,6,5,4,3,2]
            soma += SubStr(IE, A_Index, 1) * n

        dv := Mod(soma, 11) < 2 ? 0 : 11 - Mod(soma, 11)
        if dv != SubStr(IE, 9, 1)
            return false

        soma := 0
        For n in [4,3,2,7,6,5,4,3,2]
            soma += SubStr(IE, A_Index, 1) * n

        dv := Mod(soma, 11) < 2 ? 0 : 11 - Mod(soma, 11)
        if dv != SubStr(IE, 10, 1)
            return false

        return true
    }
    if UF = 'RJ' or UF = 33
    {
        IE := IEpad := Format('{:08}', IE)
		if StrLen(IE) > 8
			return false

        soma := 0
        For n in [2,7,6,5,4,3,2]
            soma += SubStr(IE, A_Index, 1) * n

        dv := Mod(soma, 11) < 2 ? 0 : 11 - Mod(soma, 11)
        if dv != SubStr(IE, 8, 1)
            return false

        return true
    }
    if UF = 'RR' or UF = 14
    {
        if StrLen(IE) != 9 or SubStr(IE, 1, 2) != 24
            return false

        soma := 0
        For n in [1,2,3,4,5,6,7,8]
            soma += SubStr(IE, A_Index, 1) * n

        dv := Mod(soma, 9)
        if dv != SubStr(IE, 9, 1)
            return false

        IEpad := IE
        return true
    }
    if UF = 'AP' or UF = 16
    {
        IE := IEpad := Format('{:09}', IE)
        if SubStr(IE, 1, 2) != '03' or StrLen(IE) != 9
            return false

        strTrb := SubStr(IE, 1, 8)
        if strTrb >= 03000001 and strTrb <= 03017000
            p := 5, d := 0
        if strTrb >= 03017001 and strTrb <= 03019022
            p := 9, d := 1
        if strTrb >= 03019023
            p := 0, d := 0

        soma := p
        For n in [9,8,7,6,5,4,3,2]
            soma += SubStr(IE, A_Index, 1) * n

        dv := 11 - Mod(soma, 11) = 10 ? 0 : 11 - Mod(soma, 11) = 11 ? d : 11 - Mod(soma, 11)
        if dv != SubStr(IE, 9, 1)
            return false

        return true
    }
    if UF = 'BA' or UF = 29
    {
		if StrLen(IE) > 9
			return false

        if StrLen(IE) = 8
        {
            if SubStr(IE, 1, 1) <= 5 or SubStr(IE, 1, 1) = 8
            {
                soma := 0
                For n in [7,6,5,4,3,2]
                    soma += SubStr(IE, A_Index, 1) * n

                dv := Mod(soma, 10) ? 10 - Mod(soma, 10) : 0
                if dv != SubStr(IE, 8, 1)
                    return false

                soma := 0
                For n in [8,7,6,5,4,3]
                    soma += SubStr(IE, A_Index, 1) * n

                soma += 2 * SubStr(IE, 8, 1)
                dv := Mod(soma, 10) ? 10 - Mod(soma, 10) : 0
                if dv != SubStr(IE, 7, 1)
                    return false

                return true
            }
            if SubStr(IE, 1, 1) = 6 or SubStr(IE, 1, 1) = 7 or SubStr(IE, 1, 1) = 9
            {
                dv := modulo11(SubStr(IE, 1, 6))
                if dv != SubStr(IE, 8, 1)
                    return false

                str := SubStr(IE, 1, 6) . SubStr(IE, 8, 1)
                dv := modulo11(str)

                if dv != SubStr(IE, 7, 1)
                    return false

                return true
            }
        }
        if StrLen(IE) = 9
        {
            if SubStr(IE, 2, 1) <= 5 or SubStr(IE, 2, 1) = 8
            {
                soma := 0
                For n in [8,7,6,5,4,3,2]
                    soma += SubStr(IE, A_Index, 1) * n

                dv := Mod(soma, 10) ? 10 - Mod(soma, 10) : 0
                if dv != SubStr(IE, 9, 1)
                    return false

                soma := 0
                For n in [9,8,7,6,5,4,3]
                    soma += SubStr(IE, A_Index, 1) * n

                soma += 2 * SubStr(IE, 9, 1)
                dv := Mod(soma, 10) ? 10 - Mod(soma, 10) : 0
                if dv != SubStr(IE, 8, 1)
                    return false

                return true
            }
            if SubStr(IE, 2, 1) = 6 or SubStr(IE, 2, 1) = 7 or SubStr(IE, 2, 1) = 9
            {
                dv := modulo11(SubStr(IE, 1, 7))
                if dv != SubStr(IE, 9, 1)
                    return false

                str := SubStr(IE, 1, 7) . SubStr(IE, 9, 1)
                dv := modulo11(str)
                if dv != SubStr(IE, 8, 1)
                    return false

                return true
            }
        }
    }
    if UF = 'PE' or UF = 26
    {
        if StrLen(IE) <= 9
        {
            IEpad := Format('{:09}', IE)
            if !validaMod11(SubStr(IE, 1, 8))
                return false

            if !validaMod11(SubStr(IE, 1, 9))
                return false

            return true
        }
        else if StrLen(IE) > 9 and StrLen(IE) <= 14
        {
            IE := IEpad := Format('{:014}', IE)
            soma := 0
            For n in [5,4,3,2,1,9,8,7,6,5,4,3,2]
                soma += SubStr(IE, A_Index, 1) * n

            resto := 11 - Mod(soma, 11)
            dv := resto > 9 ? 10 - resto : resto
            if dv != SubStr(IE, 14, 1)
                return false

            return true
        }
		else
			return false
    }
    if UF = 'RN' or UF = 24
    {
        if SubStr(IE, 1, 2) != 20 or StrLen(IE) > 10
            return false

        if StrLen(IE) = 9
        {
            soma := 0
            For n in [9,8,7,6,5,4,3,2]
                soma += SubStr(IE, A_Index, 1) * n

            dv := Mod(soma * 10, 11) > 9 ? 0 : Mod(soma * 10, 11)
            if dv != SubStr(IE, 9, 1)
                return false

            IEpad := IE
            return true
        }
        else if StrLen(IE) = 10
        {
            soma := 0
            For n in [10,9,8,7,6,5,4,3,2]
                soma += SubStr(IE, A_Index, 1) * n

            dv := Mod(soma * 10, 11) > 9 ? 0 : Mod(soma * 10, 11)
            if dv != SubStr(IE, 10, 1)
                return false

            IEpad := IE
            return true
        }
        else
            return false
    }
    if UF = 'SP' or UF = 35
    {
        if SubStr(IE, 1, 1) = 'P'
        {
            soma := 0
            For n in [1,3,4,5,6,7,8,10]
                soma += SubStr(IE, A_Index, 1) * n

            dv := SubStr(Mod(soma, 11), -1)
            if dv != SubStr(IE, 10, 1)
                return false
        }
        else
        {
            IE := IEpad := Format('{:012}', IE)
			if StrLen(IE) > 12
				return false

            soma := 0
            For n in [1,3,4,5,6,7,8,10]
                soma += SubStr(IE, A_Index, 1) * n

            dv := SubStr(Mod(soma, 11), -1)
            if dv != SubStr(IE, 9, 1)
                return false

            soma := 0
            For n in [3,2,10,9,8,7,6,5,4,3,2]
                soma += SubStr(IE, A_Index, 1) * n

            dv := SubStr(Mod(soma, 11), -1)
            if dv != SubStr(IE, 12, 1)
                return false
        }
        return true
    }
}

/**
 * Valida uma chave de acesso NFe opcionalmente retornando suas propriedades
 * @param {String} chave Chave de acesso. Não importa se está com pontuação
 * @param {&VarRef} obj Objeto com as seguintes propriedades:
 * @prop {prop} UF abreviação do estado
 * @prop {prop} data ano e mês de emissão da nota
 * @prop {prop} CNPJ CNPJ do emissor da nota
 * @prop {prop} modelo Modelo da nota. NFe, CTe, etc...
 * @prop {prop} serie Série do documento eletrônico
 * @prop {prop} nNF Número do documento já convertido em inteiro
 * @prop {prop} tpEmissao Tipo de emissão da nota. Normal, Contingência, etc...
 * @returns {Boolean} 
 */
validaChave(chave, &obj:='')
{
    chave := RegExReplace(chave, '\D')
    if StrLen(chave) != 44 or !validaMod11(chave)
        return false

    cUF := Map(11, 'RO', 12, 'AC', 13, 'AM', 14, 'RR', 15, 'PA',
     16, 'AP', 17, 'TO', 21, 'MA', 22, 'PI', 23, 'CE', 24, 'RN',
     25, 'PB', 26, 'PE', 27, 'AL', 28, 'SE', 29, 'BA', 31, 'MG',
     32, 'ES', 33, 'RJ', 35, 'SP', 41, 'PR', 42, 'SC', 43, 'RS',
     50, 'MS', 51, 'MT', 52, 'GO', 53, 'DF' )

    obj := { UF: cUF[+SubStr(chave, 1, 2)], data: '20' SubStr(chave, 3, 4),
        CNPJ: SubStr(chave, 7, 14), modelo: SubStr(chave, 21, 2), serie: SubStr(chave, 23, 3),
        nNF: +SubStr(chave, 26, 9), tpEmissao: SubStr(chave, 35, 1) }

    return true
}

/**
 * Retorna verdadeiro se o digito verificador estiver correto
 * Esta é uma fórmula muito comum e por isso a função pode ser
 * reutilizada em diferentes validações
 * @param {Integer} num Sequência a ser validada
 * @returns {Boolean} 
 */
validaMod11(num)
{
    peso := [2,3,4,5,6,7,8,9], soma := 0
    Loop StrLen(num) - 1
    {
        i := Mod(A_Index, 8) ? Mod(A_Index, 8) : 8
        soma += peso[i] * SubStr(num, -A_index -1, 1)
    }
    dv := Mod(soma, 11) < 2 ? 0 : 11 - Mod(soma, 11)
    return dv = SubStr(num, -1) ? true : false
}

/**
 * Retorna o dígito verificador do módulo 11
 * @param {Integer} num número
 * @returns {Integer} dígito verificador
 */
modulo11(num)
{
    peso := [2,3,4,5,6,7,8,9], soma := 0
    Loop StrLen(num)
    {
        i := Mod(A_Index, 8) ? Mod(A_Index, 8) : 8
        soma += peso[i] * SubStr(num, -A_index, 1)
    }
    return Mod(soma, 11) < 2 ? 0 : 11 - Mod(soma, 11)
}