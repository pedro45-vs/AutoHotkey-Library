/************************************************************************
 * @description Wrapper do Everything baseada no SDK
 * Documentação da API:
 * https://www.voidtools.com/support/everything/sdk/
 * @author Pedro Henrique C Xavier
 * @date 2024/01/14
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0


; Parâmetros possíveis para SetSort:
; ----------------------------------------------------------------
; EVERYTHING_SORT_NAME_ASCENDING                      (1) Default
; EVERYTHING_SORT_NAME_DESCENDING                     (2)
; EVERYTHING_SORT_PATH_ASCENDING                      (3)
; EVERYTHING_SORT_PATH_DESCENDING                     (4)
; EVERYTHING_SORT_SIZE_ASCENDING                      (5)
; EVERYTHING_SORT_SIZE_DESCENDING                     (6)
; EVERYTHING_SORT_EXTENSION_ASCENDING                 (7)
; EVERYTHING_SORT_EXTENSION_DESCENDING                (8)
; EVERYTHING_SORT_TYPE_NAME_ASCENDING                 (9)
; EVERYTHING_SORT_TYPE_NAME_DESCENDING                (10)
; EVERYTHING_SORT_DATE_CREATED_ASCENDING              (11)
; EVERYTHING_SORT_DATE_CREATED_DESCENDING             (12)
; EVERYTHING_SORT_DATE_MODIFIED_ASCENDING             (13)
; EVERYTHING_SORT_DATE_MODIFIED_DESCENDING            (14)
; EVERYTHING_SORT_ATTRIBUTES_ASCENDING                (15)
; EVERYTHING_SORT_ATTRIBUTES_DESCENDING               (16)
; EVERYTHING_SORT_FILE_LIST_FILENAME_ASCENDING        (17)
; EVERYTHING_SORT_FILE_LIST_FILENAME_DESCENDING       (18)
; EVERYTHING_SORT_RUN_COUNT_ASCENDING                 (19)
; EVERYTHING_SORT_RUN_COUNT_DESCENDING                (20)
; EVERYTHING_SORT_DATE_RECENTLY_CHANGED_ASCENDING     (21)
; EVERYTHING_SORT_DATE_RECENTLY_CHANGED_DESCENDING    (22)
; EVERYTHING_SORT_DATE_ACCESSED_ASCENDING             (23)
; EVERYTHING_SORT_DATE_ACCESSED_DESCENDING            (24)
; EVERYTHING_SORT_DATE_RUN_ASCENDING                  (25)
; EVERYTHING_SORT_DATE_RUN_DESCENDING                 (26)

class Everything
{
    /**
     * Inicia o módulo e define a classificação e o número de resultados máximos
     * @param {string} search Pesquisa a ser realizada
     * @param {number} SetSort Define a classificação utilizada. Padrão é por ordem ascendente do nome
     * @param {number} MaxResult Número máximo de resultados retornados. Padrão é todos os resultados
     */
    __New(search, SetSort := 1, MaxResult := -1)
    {
        this.hModule := DllCall('LoadLibrary', 'Str', A_LineFile '\..\Everything64.dll', 'Ptr')

        if Err := this.__LastError()
            MsgBox Err, 'Class Everything', 16

        DllCall('Everything64\Everything_SetMax', 'int', MaxResult)
        this.search := search
        this.sort := SetSort
        this.__ExecuteQuery()
    }
    /**
     * Retorna o nome do arquivo baseado no índice
     * @param {number} index Índice do resultado retornado. Padrão é 1.
     * @returns {string} filename Nome completo do arquivo
     */
    getFile(index := 1)
    {
        if index > this.NumRes
            return ''
        if !Ptr := DllCall('Everything64\Everything_GetResultFileName', 'Int', index - 1)
            throw ValueError(-1, this.__LastError())
        return StrGet(Ptr)
    }
    /**
     * Retorna todos os nomes de arquivos do resultado definidos na pesquisa.
     * @returns {array} com os resultados da pesquisa
     */
    getFiles()
    {
        arqs := []
        Loop this.NumRes
            arqs.push(this.getFile(A_Index))
        return arqs
    }
    /**
     * Retorna o nome dos arquivo do resultado da pesquisa baseado no índice
     * @param {number} index Índice do resultado.
     * @returns {string} Nome do arquivo.
     */
    getFullPath(index := 1)
    {
        if index > this.NumRes
            return ''
        ; Na primeira chamada, retorna a quantidade de caracteres em TCHAR para criar o buffer do tamanho suficiente
        size := DllCall('Everything64\Everything_GetResultFullPathName', 'int', index - 1, 'ptr', 0)
        buf := Buffer(size * 4)
        ; Na segunda chamada coloca o resultado no buffer
        DllCall('Everything64\Everything_GetResultFullPathName', 'int', index - 1, 'ptr', buf)
        return StrGet(buf)
    }
    /**
     * Retorna o caminho completo dos arquivos do resultado da pesquisa
     * @returns {array} com os resultados da pesquisa
     */
    getFullPaths()
    {
        arqs := []
        Loop this.NumRes
            arqs.push(this.getFullPath(A_Index))
        return arqs
    }
    /**
     * Retorna o nome do arquivo, o caminho completo, o tamanho e a data de modificação
     * @returns {string} string no formato CSV
     */
    getTable(flags := 0x53)
    {
        ; https://www.voidtools.com/support/everything/sdk/everything_setrequestflags/
        ; 0x53 = File + FullFilePath + size + date mod
        DllCall('Everything64\Everything_SetRequestFlags', 'int', 0x53)
        this.__ExecuteQuery()
        table := [ ['ARQUIVO', 'CAMINHO', 'TAMANHO (BYTES)', 'DATA MODIFICAO'] ]
        Loop this.NumRes
        {
            size := Buffer(8)
            if !DllCall('Everything64\Everything_GetResultSize', 'int', A_Index - 1, 'UPtr', size.ptr)
                MsgBox this.__LastError()

            filetime := Buffer(8)
            if !DllCall('Everything64\Everything_GetResultDateModified', 'int', A_Index - 1, 'UPtr', filetime.ptr)
                MsgBox this.__LastError()

            table.Push( [this.getFile(A_Index), this.getFullPath(A_Index), NumGet(size, 'Int64'), ConvTime(NumGet(filetime, 'Int64')) ] )
        }
        return table
        /**
         * Converte de nanossegundos para segundos, e desconta o fuso horário brasileiro
         * Uma struture filetime retorna o tempo em 100 nanossegundos
         * @param {timestamp} time timestamp da API Windows
         * @returns {string} string no padrão ISO
         */
        ConvTime(time)
        {
            date := DateAdd(16010101000000, time // 10000000, 'seconds')
            date := DateAdd(date, -3, 'hours')
            return FormatTime(date, 'yyyy-MM-dd HH:mm:ss')
        }
    }
    /**
     * Método interno, não usar diretamente
     * Define os parâmetros da pesquisa e chama a API.
     */
    __ExecuteQuery()
    {
        ; Coloca a string a ser pesquisada na memória
        DllCall('Everything64\Everything_SetSearch', 'Ptr', StrPtr(this.search))
        ; Define o tipo de classificação a ser usado. Padrão = 1
        DllCall('Everything64\Everything_SetSort', 'int', this.sort)
        ; Executa o comando da pesquisa
        DllCall('Everything64\Everything_Query', 'int', 1)
        ; Retorna o número do resultado da pesquisa
        this.NumRes := DllCall('Everything64\Everything_GetNumResults')
    }
    /**
     * Método interno, não usar diretamente
     * Retorna uma mensagem de erro da API do Everything
     */
    __LastError()
    {
        switch DllCall('Everything64\Everything_GetLastError')
        {
            case 0: err := false
            case 1: err := 'Failed to allocate memory for the search query.'
            case 2: err := 'IPC is not available.'
            case 3: err := 'Failed to register the search query window class.'
            case 4: err := 'Failed to create the search query window.'
            case 5: err := 'Failed to create the search query thread.'
            case 6: err := 'Invalid index. The index must be greater or equal to 0 and less than the number of visible results.'
            case 7: err := 'Invalid call.'
        }
        return err
    }
    __Delete() => DllCall('FreeLibrary', 'Ptr', this.hModule)
}
