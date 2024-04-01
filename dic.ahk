/************************************************************************
 * @description Classe para emular as funções nativas do Map mas usando
 * as características do Scripting.Dictionary object.
 * Uma das caracteríticas é que a ordem de inserção das chaves é mantida
 * @author Pedro Henrique C. Xavier
 * @date 2024-03-07
 * @version 2.1-alpha.8
 ***********************************************************************/

#Requires AutoHotkey v2.0

class dic
{
    ; Limpa todas as chaves-valor do dicionário
    Clear() => this.dic.RemoveAll()
    ; Retorna uma cópia do dicionário
    Clone()
    {
        newdic := dic()
        for key in this.dic
            newdic[key] := (this.dic.Items)[A_Index -1]
        return newdic
    }
    ; Apaga uma chave-valor do dicionário
    Delete(key) => this.dic.Remove(key)
    ; Recupera o valor da chave, ou um valor padrão
    Get(key, default := unset)
    {
        if this.has(key)
            return this.dic.item[key]
        else
            return default
    }
    ; Checa de há uma chave no dicionário
    Has(key) => this.dic.Exists(key)
    ; Insere as chaves-valor separadas por vírgula.
    Set(items*)
    {
        for item in items
        {
            if A_Index & 1
                this.dic.Add(items[A_Index], items[A_Index + 1])
        }
    }
    ; Cria novo dicionário, opcionalmente inserindo as chaves-valor separadas por vírgula.
    __New(items*)
    {
        this.dic := ComObject('Scripting.Dictionary')
        this.Set(items*)
    }
    ; Enumera as chaves-valor do dicionário utilizando o comando FOR
    __Enum(key, value?)
    {
        num := 0
        Enumerator(&key, &value?)
        {
            if num >= this.count
                return false

            key := (this.dic.Keys)[num]
            value := (this.dic.Items)[num]
            num++
            return true
        }
        return Enumerator
    }
    ; Retorna a quantidade de chaves-valor do dicionário
    Count => this.dic.Count
    ; Define o valor padrão para quando a chave não for encontrada
    Default := unset
    ; Retorna a chave no índice especificado
    Key[index] => this.Keys[index -1]
    ; Retorna uma array com as chaves do Map
    Keys => this.dic.Keys
    ; Retorna o valor no índice especificado
    Value[index] => this.Values[index -1]
    ; Retorna uma array com os valores do dicionário
    Values => this.dic.Items
    ; Seta ou retorna uma chave-valor do dicionário
    __Item[key]
    {
        get => this.has(key) ? this.dic.item[key] : this.default
        set => this.dic.item[key] := value
    }
}