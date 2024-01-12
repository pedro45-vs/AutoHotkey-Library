/************************************************************************
 * @description Acrescenta métodos para os objetos do tipo Map
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/11
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0

Map.Prototype.Base := MapExtended()

class MapExtended
{
    ; Retorna a chave no índice especificado
    key[index] => this.Keys[index]
    ; Retorna uma array com as chaves do Map
    keys => [this*]
    ; Retorna o valor no índice especificado
    value[index] => this.values[index]
    ; Retorna uma array com os valores do Map
    values => [this.__Enum(2).Bind(&_)*]
}