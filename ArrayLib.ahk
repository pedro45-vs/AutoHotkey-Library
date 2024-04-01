/************************************************************************
 * @description Acrescenta métodos para os objetos do tipo Array
 * @author Pedro Henrique C. Xavier
 * @date 2024-03-07
 * @version 2.1-alpha.8
 ***********************************************************************/

#Requires AutoHotkey v2.0

Array.Prototype.Base := ExtendedArray

class ExtendedArray
{
    /**
     * Retorna o índice do valor contido na array
     * @param value
     * @returns {number}
     */
    static IndexOf(value)
    {
        for item in this
            if item == value
                return A_Index
        return false
    }
    /**
     * Retorna o índice do valor contido na array usando os mesmos parâmentros
     * da função InStr()
     * @param {value} value valor a ser procurado
     * @param {string} CaseSense Mesmas opções da função InStr()
     * @param {integer} StartingPos indice para inicio da pesquisa
     * @param {integer} Occurrence Se diferente de 0, retornará uma array com os indices
     * @returns {integer}
     */
    static Search(value, CaseSense := 0, StartingPos := 1, Occurrence := false)
    {
        if Occurrence
        {
            arr := []
            for item in this
                if A_Index >= StartingPos and InStr(item, value, CaseSense)
                    arr.push(A_Index)

            return arr
        }
        else
        {
            for item in this
                if A_Index >= StartingPos and InStr(item, value, CaseSense)
                    return A_Index
        }
        return false
    }
    /**
     * Classifica uma array de strings
     * @param {string} opt mesmas opções da função Sort()
     * @returns {Array}
     */
    static Sort(opt := '')
    {
        for item in this
            str .= item '`n'
        
        this.Length := 0
        this.Push( StrSplit(Sort(SubStr(str, 1, -1), opt), '`n')* )
        return this
    }
    /**
     * Retorna uma string com o conteúdo da array separado por vírgulas
     */
    static ToString()
    {
        for item in this
            str .= item ', '
        return SubStr(str, 1, -2)
    }
    /**
     * Aplica uma função para cada elemento da array, modificando-o
     * @param func Função que aceita um argumento
     */
     static Map(func)
     {
         for item, value in this
            this[item] := func(value)
            
        return this
     }
}
