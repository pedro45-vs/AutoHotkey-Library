/************************************************************************
 * @description Acrescenta métodos para os objetos do tipo Map
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/09
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0

Map.Prototype.Base := MapExtended()

class MapExtended
{
    /**
     * Retorna o valor do Map no indice especificado
     * @param index
     * @returns {any}
     */
    item(index)
    {
        ind := index < 0 ? this.Count + index + 1 : index
        for k in this
            if A_Index = ind
                return this[k]
    }
    ; Retorna uma array com as chaves do Map
    keys   => [this*]
    ; Retorna uma array com os valores do Map
    values => [this.__Enum(2).Bind(&_)*]
}