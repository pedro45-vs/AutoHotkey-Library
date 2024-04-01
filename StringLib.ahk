/************************************************************************
 * @description Biblioteca com funções diversas para trabalhar com Strings
 * @author Pedro Henrique C Xavier
 * @date 2023/08/22
 * @version 2.0.5
 ***********************************************************************/

/**
 * Remove a acentuação das palavras
 * @param {string} str String acentuada
 * @returns {string} String sem acentos
 */
RemoverAcentuacao(str)
{
	CaracAcentuados := 'àáâãèéêìíîòóôõùúûüÀÁÂÃÈÉÊÌÍÎÒÓÔÕÙÚÛÜçÇ&'
	CaracSubstituto := 'aaaaeeeiiioooouuuuAAAAEEEIIIOOOOUUUUcCE'

	Loop StrLen(CaracAcentuados)
		str := StrReplace(str, SubStr(CaracAcentuados, A_Index, 1), SubStr(CaracSubstituto, A_Index, 1))

	return str
}

/**
 * Retorna a primeira letra maíuscula e as demais minúsculas
 * @param {string} str
 * @returns {string} 
 */
modofrase(str)
{
    return StrUpper(SubStr(str, 1, 1)) StrLower(SubStr(str, 2))
}

/**
 * Retorna a string invertida (de trás para frente)
 * @param {string} str 
 * @param {Boolean} world Se verdadeiro, a inversão será palavra por palavra
 * @returns {string} 
 */
InvStr(str, world:=false)
{
    if world
    {
        pos := 0
        Loop Parse, str, '`s`t`n'
        {
            pos += StrLen(A_LoopField) + 1
            del := SubStr(str, pos, 1)

            Loop StrLen(A_LoopField)
                res .= SubStr(A_LoopField, -A_Index, 1)

            res .= del
        }
        return res
    }
    else
    {
        Loop StrLen(str)
            res .= SubStr(str, -A_Index, 1)

        return res
    }
}

/**
 * Função para classificar uma string em ordem aleatória.
 * Permite embaralhar letras individuais, diferentemente da função nativa Sort()
 * @param {string} str 
 * @param {string} delimiter Delimitador 
 * @returns {string} 
 */
SortStr(str, delimiter:='')
{
    StrNormal := StrSplit(str, delimiter)
    StrRandom := Array()
    Loop StrNormal.Length
    {
        i := Random(1, StrNormal.Length)
        StrRandom.push(StrNormal[i])
        StrNormal.RemoveAt(i)
    }
    For val in StrRandom
        StrR .= val delimiter
    return StrR
}

/**
 * Retorna uma string classificada pela coluna indicada.
 * Útil para classificar um arquivo CSV usando uma coluna específica.
 * @param {string} str String a ser classificada
 * @param {number} col Número da coluna considerada
 * @param {string} delimiter Delimitador das colunas da string
 * @param {bolelan} cabecalho ignora a primeira linha
 * @returns {string}
 */
SortCol(str, col, delimiter := ';', cabecalho := false)
{
    if cabecalho
    {
        newstr := str, str := ''
        loop parse newstr, '`n', '`r'
            (A_index != 1) ? str .= A_LoopField '`n' : cab := A_LoopField '`n'
    }
    SortColumn(lin1, lin2, *)
    {
        lin1 := StrSplit(lin1, delimiter)[col]
        lin2 := StrSplit(lin2, delimiter)[col]
        return StrCompare(lin1, lin2)
    }
    return (cab ?? '') . Sort(str, , SortColumn)
}

/**
 * Retorna a string com case aleatória
 * @param {string} str 
 * @returns {string} 
 */
CaseCrazy(str)
{
    For char in StrSplit(str)
        r .= (~A_Index & 1) ? StrUpper(char) : StrLower(char)
    return r
}

/**
 * Retorna verdadeiro se o valor estiver entre o valor mínimo e o máximo
 * @param {String, Number} value 
 * @param {String, Number} valueMin 
 * @param {String, Number} valueMax 
 * @returns {number} 
 */
IsStrBetween(value, valueMin, valueMax)
{
    if IsNumber(value)
        return Value >= valueMin && Value <= ValueMax

    else if IsAlpha(value)
        return StrCompare(value, valueMin, 'Locale') >= 0
        && StrCompare(value, valueMax, 'Locale') <= 0
}
/**
 * Recupera os dados usando notação de objeto
 * Permite alterar os campos e gerar o nome do arquivo
 * Usa o método .ToString() que habilita a função String(obj)
 * @param {String} Path Caminho do arquivo
*/
class PathSplit
{
    __New(Path)
    {
        SplitPath Path, &OutFileName, &OutDir, &OutExtension, &OutNameNoExt, &OutDrive
        this.FileName := OutFileName
        this.Dir := OutDir
        this.Extension := OutExtension
        this.NameNoExt := OutNameNoExt
        this.Drive := OutDrive
    }
    /**
     * Extensão do arquivo
     * @prop {String} Ext
     */
    Ext {
        get => this.Extension
        set => this.FileName := This.NameNoExt '.' Value
    }
    /**
     * Nome do arquivo
     * @prop {String} Name
     */
    Name {
        get => this.NameNoExt
        set => this.FileName := Value '.' this.Extension
    }
    /**
     * Retorna as propriedades do caminho em string
     * @returns {string} 
     */
    ToString() => this.Dir '\' this.FileName
}

/**
 * Returns all RegExMatch results in an array: [RegExMatchInfo1, RegExMatchInfo2, ...]
 * @param haystack The string whose content is searched.
 * @param needleRegEx The RegEx pattern to search for.
 * @param startingPosition If StartingPos is omitted, it defaults to 1 (the beginning of haystack).
 * @returns {Array}
 */
RegExMatchAll(haystack, needleRegEx, startingPosition := 1)
{
	reg := []
	while startingPosition := RegExMatch(haystack, needleRegEx, &outputVar, startingPosition)
    {
		reg.Push(outputVar)
        startingPosition += outputVar[0] ? StrLen(outputVar[0]) : 1
	}
	return reg
}

/**
 * Substitui os caractes de Escape comuns pelo padrão AutoHotKey
 * @param {string} str String a ser convertida
 * @returns {string} 
 */
EscapeString(str)
{
    for esc, char in Map('\n', '`n', '\r', '`r', '\t', '`t')
        str := StrReplace(str, esc, char)
    return str
}

/**
 * Normaliza nomes antes de pesquisar.
 * Junta as siglas para facilitar a pesquisa
 * @param {string} str String a ser normalizada
 * @returns {string} 
 */
NormalizarNomes(str)
{
	; Remove todos os pontos e caracteres especiais
	sub := RegExReplace(str, '(*UCP)[^a-zA-Z0-9 ]')
	; Remove todos os espaços entre duas letras sozinhas (junta as siglas)
	return RegExReplace(sub, '\s+(?=\b\w\b)') 
}

/**
 * Formata uma string numérica com o formato padrão do CPF ou CNPJ
 *
 */
FormatarCNPJCPF(str)
{
	if StrLen(str) = 14
		return RegExReplace(str,'(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})','$1.$2.$3/$4-$5')

	else if StrLen(str) = 11
		return RegExReplace(str,'(\d{3})(\d{3})(\d{3})(\d{2})','$1.$2.$3-$4')
}