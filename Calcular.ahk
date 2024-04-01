/************************************************************************
 * @description calula uma expressão numérica
 * @author Pedro Henrique C. Xavier
 * @date 2023-08-23
 * @version 2.0.5
 ***********************************************************************/
#Requires AutoHotKey v2.0

/*
expr1 := '2,5+3*7,9-6/9,8'
expr2 := '(3 + 2) * 2'
expr3 := '((4-2)*(5+3,4))/(3*2)'
expr4 := '-4*5'
expr5 := '5*6%'
expr6 := '4+25%'
expr7 := 2**3
expr8 := (2+3)(4+9)+2(3/2)
*/

/**
 * Permite calcular uma expressão de soma, subtração, adição e divisão
 * * Suporta exponenciações com '**' ou '^'
 * * Suporta parênteses e precedência de operador
 * * Suporta operações com percentagem e notação científica
 * @param expr - Expressão a ser calculada
 * @returns {number} - valor em número nativo AutoHotKey
 */
Calcular(expr)
{
    ; Verifica caracteres inválidos na expressão
    if expr ~= '[^\d\s\.\,\+\-\*\/\%\^\x28\x29Ee]'
        throw ValueError('caracteres inválidos', -1)

    ; Remove espaços em branco, tabulações e quebras de linha
    expr := RegExReplace(expr, '\s')

    ; Converte os números para o formato nativo AutoHotKey ou retorna 0 para strings vazias
    expr := StrReplace(StrReplace(expr ? expr : 0, '.'), ',', '.')

    ; Adiciona multiplicação aos numeros encostados no parênteses de abertura
    expr := RegExReplace(expr, '(\d+\.?\d*|\x29)(\x28)', '$1*$2')

    ; Resolve as Sub-expressões dentro dos parênteses, uma a uma
    while RegExMatch(expr, '(\x28([^\x28\x29]+)\x29)', &par)
    {
        res := eval(par[2])
        expr := StrReplace(expr, par[1], res)
    }

    ; Se ainda sobrar sobrar algum parênteses retorna mensagem de erro
    if expr ~= '(\x28|\x29)'
        throw ValueError('parênteses inválidos', -1)

    calc := eval(expr)
    return IsNumber(calc) ? calc : 0
}

; Função de apoio
eval(expr)
{
    ; Substitui os percentuais multiplicadores e divisores
    while RegExMatch(expr, '(\*|\/)(\d+\.?\d*)\%', &per)
    {
        switch per[1]
        {
            case '*': expr := StrReplace(expr, per[0], '*' per[2] / 100)
            case '/': expr := StrReplace(expr, per[0], '/' per[2] / 100)
        }
    }

    ; Substitui os percentuais de soma e subtração
    while RegExMatch(expr, '(\+|\-)(\d+\.?\d*)\%', &per)
    {
        switch per[1]
        {
            case '+': expr := StrReplace(expr, per[0], '*' 1 + per[2] / 100)
            case '-': expr := StrReplace(expr, per[0], '*' 1 - per[2] / 100)
        }
    }

    ; Calcula as exponenciações ('**' ou '^')
    while RegExMatch(expr, '(-?\d+(\.\d+)?([eE]-?\d+)?)(\*\*|\^)(-?\d+(\.\d+)?([eE]-?\d+)?)', &op)
    {
        switch op[4]
        {
            case '**': res := op[1] ** op[5]
            case '^': res := op[1] ** op[5]
        }
        expr := StrReplace(expr, op[0], res)
    }

    ; Calcula todas as multiplicações e divisões
    while RegExMatch(expr, '(-?\d+(\.\d+)?([eE]-?\d+)?)(\/\/|\/|\*)(-?\d+(\.\d+)?([eE]-?\d+)?)', &op)
    {
        switch op[4]
        {
            case '*': res := op[1] * op[5]
            case '/': res := op[1] / op[5]
            case '//': res := op[1] // op[5]
        }
        expr := StrReplace(expr, op[0], res)
    }

    ; Calcula todas as adições e subtrações
    while RegExMatch(expr, '(-?\d+(\.\d+)?([eE]-?\d+)?)(\+|\-)(-?\d+(\.\d+)?([eE]-?\d+)?)', &op)
    {
        switch op[4]
        {
            case '+': res := op[1] + op[5]
            case '-': res := op[1] - op[5]
        }
        expr := StrReplace(expr, op[0], res)
    }
    return expr
}
