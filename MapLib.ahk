/************************************************************************
 * @description Acrescenta métodos para os objetos do tipo Map
 * @author Pedro Henrique C. Xavier
 * @date 2024-03-13
 * @version 2.1-alpha.8
 ***********************************************************************/

#Requires AutoHotkey v2.0

Map.Prototype.Base := MapExtended

class MapExtended
{
    ; Retorna a chave no índice especificado
    static key[index] => this.Keys[index]
    ; Retorna uma array com as chaves do Map
    static keys => [this*]
    ; Retorna o valor no índice especificado
    static value[index] => this.values[index]
    ; Retorna uma array com os valores do Map
    static values => [this.__Enum(2).Bind(&_)*]
}