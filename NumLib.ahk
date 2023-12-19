/************************************************************************
 * @description Biblioteca para trabalhar com números em padrões diferentes
 * @author Pedro Henrique C. Xavier
 * @date 2023/12/04
 * @version 2.0.10
 ***********************************************************************/

#Requires AutoHotKey v2.0

/**
 * Retorna um número formatado com separador de milhar. Ex 1.234,56
 * @param {number} num Número a ser convertido, tanto nativo quanto no formato brasileiro
 * @param {number} dec Quantidade de casas decimais retornadas
 * @returns {string} String formatada
 */
milhar(num := 0, dec := 2)
{
    num := !IsNumber(num) ? ToNum(num) : num
    return RegExReplace(StrReplace(Round(num, dec), '.', ','), '\d(?=(\d{3})+(\D|$))', '$0.')
}

/**
 * Converte um número no padrão brasileiro para o formato nativo AutoHotkey
 * @param {string} str Número no padrão brasileiro. Ex. 1.234,56
 * @returns {number} Número nativo
 */
ToNum(str := 0) => +StrReplace(StrReplace(str || 0, '.'), ',', '.')

/**
 * Retorna verdadeiro se o valor estiver entre o valor mínimo e o máximo
 * @param {number} value Valor a ser comparado
 * @param {number} valueMin Valor mínimo
 * @param {number} valueMax Valor máximo
 * @returns {boolean} verdadeiro ou falso
 */
IsBetween(value, valueMin, valueMax) => Value >= valueMin && Value <= ValueMax

/**
 * Converte números bancários em números nativos AutoHotKey
 * Ex1.: 125,25C   -> 125.25
 * Ex2.: 1.335,37D -> -1335.37
 * @param {string} num Número em formato bancário
 * @returns {number}
 */
ConverterNumBancario(num)
{
    if RegExMatch(num, '(-?\d+(?:\.\d{3})*\,\d\d)\s?([CD]?)', &re)
        return (re[2] = 'D') ? -ToNum(re[1]) : ToNum(re[1])
}
