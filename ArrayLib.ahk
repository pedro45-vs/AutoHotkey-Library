/************************************************************************
 * @description Acrescenta métodos para os objetos do tipo Array
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/09
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0

Array.Prototype.Base := ExtendedArray()

class ExtendedArray
{
    /**
     * Retorna o índice do valor contido na array
     * @param value
     * @returns {string|number}
     */
    IndexOf(value)
    {
        for item in this
            if item == value
                return A_Index
        return false
    }
    /**
     * Retorna uma array classificada
     * @param {string} opt mesmas opções da função Sort()
     */
    Sort(opt := '')
    {
        sorted := []
        for item in this
            str .= item '`n'
        return StrSplit(Sort(SubStr(str, 1, -1), opt), '`n')
    }
    /**
     * Retorna uma string com o conteúdo da array separado por vírgulas
     */
    ToString()
    {
        for item in this
            str .= item ', '
        return SubStr(str, 1, -2)
    }
}
