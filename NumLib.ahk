/************************************************************************
 * @description Biblioteca para trabalhar com números em padrões diferentes
 * @author Pedro Henrique C. Xavier
 * @date 2024-04-01
 * @version 2.1-alpha.9
 ***********************************************************************/

#Requires AutoHotKey v2.0

/**
 * Retorna um número formatado com separador de milhar. Ex 1.234,56
 * @param {number} num Número a ser convertido, tanto nativo quanto no formato brasileiro
 * @param {number} dec Quantidade de casas decimais retornadas
 * @returns {string} String formatada
 */
Milhar(num := 0, dec := 2)
{
    IsNumber(num) || num := ToNum(num)
    return RegExReplace(StrReplace(Round(num, dec), '.', ','), '\G\d+?(?=(\d{3})+(?:\D|$))', '$0.')
}

/**
 * Retorna uma string numérica no padrão brasileiro utilizando a API do Windows
 * @param {number} num Número a ser convertido
 * @returns {string} String formatada
 */
Milhar2(num)
{
    VarSetStrCapacity(&out, 260)
    if DllCall('GetNumberFormatEx', 'str', 'pt-BR', 'int', 0, 'str', num, 'int', 0, 'str', out, 'int', 260)
        return out
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

/**
 * Converte um número nativo em formato moeda com R$
 * Ex.: 12.35 -> R$ 12,35
 * @param {number} num
 * @returns {string}
 */
Moeda(num)
{
    VarSetStrCapacity(&out, 260)
    if DllCall('GetCurrencyFormatEx', 'str', 'pt-BR', 'int', 0, 'str', num, 'int', 0, 'str', out, 'int', 260)
        return out
}

/**
 * Converte um numero nativo AutoHotKey em formato brasileiro,
 * apenas trocando o ponto decimal pela vírgula.
 * @param {number} num Número ou string numérica
 * @returns {string} String numérica
 */
NumBr(num)
{
    IsNumber(num) || num := ToNum(num)
    return StrReplace(num, '.', ',')
}

/**
 * Retorna a soma dos valores passados. Pode ser um número nativo ou string numérica
 * @param {variadic} values Números ou strings numéricas
 * @returns {string} String numérica
 */
Somar(values*)
{
    sum := 0
    for value in values
        sum += toNum(value)
   return milhar(sum)
}

/**
 * Retorna a quantidade dos argumentos passados
 * @param {Variadic} values Qualquer valor
 * @returns {Integer}
 */
Contar(values*) => Values.Length

/**
 * Retorna o valor máximo dos argumentos da função
 * @param {Variadic} values Números ou strings numéricas
 * @returns {string} String numérica
 */
Maximo(values*)
{
    arr := []
    for value in values
        arr.Push(ToNum(value))
    return milhar(Max(arr*))
}

/**
 * Retorna o valor mínimo dos argumentos da função
 * @param {Variadic} values Números ou strings numéricas
 * @returns {string} String numérica
 */
Minimo(values*)
{
    arr := []
    for value in values
        arr.Push(ToNum(value))
    return milhar(Min(arr*))
}

/**
 * Retorna o valor médio dos argumentos da função
 * @param {Variadic} values Números ou strings numéricas
 * @returns {string} String numérica
 */
Media(values*)
{
    sum := 0
    for value in values
        sum += toNum(value)
    return milhar(sum / values.Length)
}