/************************************************************************
 * @description Classe para trabalhar com dados tabulares
 * @author Pedro Henrique C. Xavier
 * @date 2024-03-29
 * @version 2.1-alpha.8
 ***********************************************************************/

#Requires AutoHotkey v2.0

class ClassTable
{
    ; Retorna a quantidade de linhas da tabela, sem contar o cabeçalho
    count => this.table.Length
    /**
     * Permite iniciar a classe definindo os dados carregados diretamente
     * @param {String} table String em formato tabular
     * @param {Array} header Se omitido, usará a primeira linha dos dados
     * @param {String} Delimiter Delimitador dos dados
     * @returns {ClassTable}
     */
    __New(table := '', header := unset, Delimiter := ';')
    {
        this.table := [], this.header := []
        Loop parse table, '`n', '`r'
        {
            if not A_LoopField
                continue

            columns := StrSplit(A_LoopField, Delimiter)
            if A_Index = 1
            {
                if IsSet(header)
                {
                    this.Header := header
                    this.table.Push(columns)
                }
                else
                    this.Header := columns
            }
            else
                this.table.Push(columns)
        }
        this.Delimiter := Delimiter, this.Encoding := 'UTF-8'
        return this
    }
    /**
     * Carrega o arquivo com opção para definir o delimitador e o enconding
     * @param FilePath Caminho do arquivo para abrir
     * @param {String} Delimiter Opcional. Delimitador usado, padrão é ';'
     * @param {String} Encoding Opcional. Encondig usado, padrão é 'UTF-8'
     * @returns {ClassTable}
     */
    Load(FilePath, Delimiter := this.Delimiter, Encoding := this.Encoding)
    {
        FileEncoding(this.Encoding)
        this.table := [], this.header := []
        Loop read FilePath
        {
            columns := StrSplit(A_LoopReadLine, Delimiter)
            A_index = 1 ? this.header := columns : this.table.Push(columns)
        }
        this.Encoding := Encoding
        this.FilePath := FilePath
        this.Delimiter := Delimiter
        return this
    }
    /**
     * Salva o arquivo, opcionalmente definindo novo delimitador e enconding
     * @param FilePath Caminho do arquivo para salvar.
     * @param {String} Delimiter Opcional. Delimitador usado, padrão é ';'
     * @param {String} Encoding Opcional. Encondig usado, padrão é 'UTF-8'
     */
    Save(FilePath := this.FilePath, Delimiter := this.Delimiter, Encoding := this.Encoding)
    {
        csvstr .= this.toCSV(this.header)
        for row in this.table
            csvstr .= this.toCSV(row)

        FileOpen(FilePath, 'w', Encoding).Write(RTrim(csvstr, '`n'))
    }
    /**
     * Filtra uma coluna da tabela, podendo ser uma string ou função
     * @param {String} field Nome da coluna
     * @param {String|Func} search String ou função callback com um argumento contendo o valor
     * @returns {ClassTable}
     */
    Filter(field, search)
    {
        index := this.HeaderIndex(field), TempTable := []
        for row in this.table
        {
            if Type(search) ~= '(Number|String)' and search = row[index]
                TempTable.Push(row)
            else if Type(search) ~= '(Func|BoundFunc)' and search(row[index])
                TempTable.Push(row)
        }
        this.table := TempTable
        return this
    }
    /**
     * Seleciona as colunas da tabela
     * @param {Variadic} columns Strings com o nome das colunas
     * @returns {ClassTable}
     */
    Select(columns*)
    {
        TempTable := [], TempHeader := []
        for row in this.table
        {
            TempRow := []
            for field in columns
                TempRow.push(row[ this.HeaderIndex(field) ])
            TempTable.push(TempRow)
        }
        for field in columns
            TempHeader.push(this.header[ this.HeaderIndex(field) ])

        this.table := TempTable, this.header := TempHeader
        return this
    }
    /**
     * Classifica a tabela de acordo com a coluna
     * @param {String} field Coluna a ser usada para classificar as linhas da tabela
     * @param {String} order ('asc'|'desc') determina a ordem de classificação
     * @param {Boolean} unique Remove resultados duplicados
     * @returns {ClassTable}
     */
    OrderBy(field, order := 'asc', unique := false)
    {
        index := this.HeaderIndex(field), datatype := DetectType(this.table[1][index])
        for row in this.table
        {
            strTable := ''
            for col in row
                strTable .= col this.Delimiter
            strTableLine .= SubStr(strTable, 1, -1) '`n'
        }
        strTableLine := Sort(RTrim(strTableLine, '`n'), unique ? 'U' : '', CustomSort)
        table := []
        Loop parse RTrim(strTableLine, '`n'), '`n'
            table.push(StrSplit(A_LoopField, this.Delimiter))

        this.table := table
        return this

        CustomSort(First, Second, Offset)
        {
            a1 := StrSplit(First, this.Delimiter)[index]
            a2 := StrSplit(Second, this.Delimiter)[index]

            if datatype = 'string'
                return StrCompare(a1, a2, 'Locale') * (order = 'desc' ? -1 : 1)

            else if datatype = 'number'
            {
                a1 := IsNumber(a1) ? a1 : +StrReplace(StrReplace(a1 || 0, '.'), ',', '.')
                a2 := IsNumber(a2) ? a2 : +StrReplace(StrReplace(a2 || 0, '.'), ',', '.')
                return (a1 > a2 ? 1 : a1 < a2 ? -1 : 0) * (order = 'desc' ? -1 : 1)
            }
            else if datatype = 'date'
            {
                a1 := RegExReplace(a1, '(\d{2})\/(\d{2})\/(\d{4})', '$3$2$1')
                a2 := RegExReplace(a2, '(\d{2})\/(\d{2})\/(\d{4})', '$3$2$1')
                return DateDiff(a1, a2, 'days') * (order = 'desc' ? -1 : 1)
            }
        }
        DetectType(value) => (value ~= '[0-9\,\.]') ? 'number' : (value ~= '\d\d\/\d\d\/\d{4}') ? 'date' : 'string'
    }
    /**
     * Transforma o conteúdo de uma coluna de acordo com a função passada
     * @param {String} field Coluna a ser usada para classificar as linhas da tabela
     * @param {Func} function Função callback com um argumento contendo o valor
     * @returns {ClassTable}
     */
    Transform(field, function)
    {
        index := this.HeaderIndex(field)
        for row in this.table
            row[index] := function( row[index] )
        return this
    }
    /**
     * Cria uma nova coluna na tabela com base na função passada
     * @param {Func} function Função usada para criar a nova coluna
     * @param {Array} callbackFields Array com as colunas usadas como callback
     * @param {String} NameCol Nome da nova coluna. Se omitido, usa o nome da função.
     * @returns {ClassTable}
     */
    Create(function, callbackFields := [], NameCol := function.Name)
    {
        this.Header.Push(NameCol)
        fields_index := []
        for item in callbackFields
            fields_index.Push( this.HeaderIndex(item) )

        for row in this.table
        {
            fields_value := []
            for item in fields_index
                fields_value.push( row[item] )
            row.Push( function(fields_value*) )
        }
        return this
    }
    /**
     * Combina o valor das colunas de acordo com campo passado
     * @param {Array} columns Array com o nome das colunas para serem combinadas
     * @param {String} filter Nome da coluna usada como filtro
     * @param {Func} function Função variada usada para combinar os valores
     * @returns {ClassTable}
     */
    Combine(columns, filter, function)
    {
        TempHeader := [filter], MapRow := Map(), fields_index := []
        for column in columns
        {
            TempHeader.Push(column)
            fields_index.Push(this.HeaderIndex(column))
        }
        index_filter := this.HeaderIndex(filter)
        for row in this.table
        {
            TempTable := [], TempTable.Length := fields_index.Length
            MapRow.has(value := row[index_filter]) || MapRow[value] := TempTable

            for index, field in fields_index
                MapRow[value].has(index) ? MapRow[value][index].Push( row[field] ) : MapRow[value][index] := [ row[field] ]
        }
        for key, arrays in MapRow
            for index, arr in arrays
                arrays[index] := function(arr*)

        TempTable := []
        for key, arrays in MapRow
            TempTable.Push([key, arrays*])

        this.table := TempTable
        this.Header := TempHeader
        return this
    }
    /**
     * Limita a quantidade de linhas na tabela, excluindo todas as outras
     * @param {Integer} value Quantidade de linhas
     * @returns {ClassTable}
     */
    Limit(value)
    {
        this.table.Length := value
        return this
    }
    /**
     * Retorna o índice do cabeçalho
     * @param {String} field Nome da coluna
     * @returns {Integer}
     */
    HeaderIndex(field)
    {
        for item, value in this.header
            if value == field
                return item
        throw ValueError('Coluna não encontrada', -2, field)
    }
    /**
     * Converte uma array em uma string
     * @param {Array} arr Array com os valores
     * @returns {String}
     */
    toCSV(arr)
    {
        for item in arr
            str .= item this.Delimiter
        return SubStr(str, 1, -1) '`n'
    }
}

ClassTable.Prototype.Base := TableViewer

class TableViewer
{
    static Show()
    {
        GuiP := Gui('Resize', this.HasProp('FilePath') ? this.FilePath : A_ScriptName)
        GuiP.OnEvent('Close', (*)=>ExitApp())
        GuiP.OnEvent('Size', (o, m, w, h) => this.LV.Move(, , w - 20, h - 40))
        GuiP.SetFont('s10', 'Segoe UI')
        this.LV := GuiP.AddListView('LV0x4000 LV0x10000 vList Count' this.Count , this.header)
        GuiP.AddStatusBar(, 'Colunas: ' this.header.Length '`tLinhas: ' this.table.Length)
        this.LV.Opt('-Redraw')

        for row in this.table
            this.LV.Add(, row*)

        Loop this.header.Length
        {
            this.LV.ModifyCol(A_Index, 'Auto')
            this.LV.ModifyCol(A_Index, 'AutoHDR')
        }
        for index, value in this.table[1]
            if value ~= '^[0-9,.\/]+$'
                this.LV.ModifyCol(index, 'Right')

        this.LV.Opt('Redraw')
        GuiP.Show(Format('w{} h{}', A_ScreenWidth * 0.6, A_ScreenHeight * 0.6))

        HotIfWinActive('ahk_id' GuiP.Hwnd)
        Hotkey('^c', this.CopySelect.Bind(this))
    }
    static CopySelect(*)
    {
        while RowNumber := this.LV.GetNext(RowNumber ?? 0)
        {
            str := ''
            Loop this.LV.GetCount('Column')
                str .= this.LV.GetText(RowNumber, A_Index) '`t'
            retrieve .= SubStr(str, 1, -1) '`n'
        }
        A_Clipboard := RTrim(retrieve, '`n')
    }
}