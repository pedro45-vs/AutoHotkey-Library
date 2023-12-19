/************************************************************************
 * @description Funções diversas para trabalhar com datas
 * @file DataLib.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/08/21
 * @version 2.0.5
 ***********************************************************************/
#Requires AutoHotkey v2.0

/**
 * Adiciona ou subtrai um número de meses a uma data no formato AutoHotkey
 * @param {Integer} num Quantidade de meses a somar ou diminuir (use número negativo)
 * @param {string} data Data no formato AutoHotkey. Se omitido usará a data atual
 * @param {string} formato Formato de saída. Se omitido usará o padrão AutoHotkey
 * @returns {string} 
 */
AdicionarMeses(num, data := A_YYYY A_MM, formato := 'yyyyMM')
{
    ano := SubStr(data, 1, 4)
    mes := SubStr(data, 5, 2) + num

    mesMod := Mod(mes, 12)
    mesInt := Floor(mes / 12)

    modMes := mesMod <= 0 ? mesMod + 12 : mesMod
    ano += mesMod ? mesInt : mesInt - 1

    return FormatTime(ano . Format('{:02}', modMes), formato)
}

/**
 * Converte datas em diferentes formatos
 * @param data String em diversos formatos.
 * @param {string} formato da data de saída
 * @returns {string} 
 */
converterData(data, formato := 'yyyyMMdd')
{
    data := RegExReplace(data, '\D')
    if IsTime(data)
        return FormatTime(data, formato)

    data := RegExReplace(data, '(\d{2})(\d{2})(\d{4})', '$3$2$1')
    if IsTime(data)
        return FormatTime(data, formato)

    data := RegExReplace(data, '(\d{2})(\d{2})(\d{2})', '20$3$2$1')
    if IsTime(data)
        return FormatTime(data, formato)
    else
        throw ValueError('Não foi possível recnhecer essa data', -2)
}

/**
 * Converte um TimeStamp AutoHotkey para o padrão ISO8601
 * @param date Data no formato AutoHotkey
 * @returns {string} TimeStamp padrão ISO
 */
DateTimeToISO(date)
{
    return IsTime(date) && FormatTime(date, 'yyyy-MM-ddTHH:mm:ss')
}

/**
 * Converte um TimeStamp no padrão ISO8601 para data formato Autohotkey
 * @param str TimeStamp padrão ISO
 * @returns {string} Data no formato AutoHotkey
 */
ISOToDateTime(str)
{
    return IsTime(str := RegExReplace(str, '\D')) ? str : 0
}

/**
 * Retorna o mês anterior no formato especificado
 * @param {string} formato string de retorno Ex. ('dd/MM/yyyy')
 * @param {boolean} opcional 
 * @returns {string} 
 */
mesPassado(formato := 'yyyyMMdd', opcional := false)
{
    DataMesPassado := DateAdd(A_YYYY A_MM '01', -1, 'days')

    if !opcional
        DataMesPassado := SubStr(DataMesPassado, 1, 6)

    return FormatTime(DataMesPassado, formato)
}

/**
 * Retorna o mês seguinte no formato especificado
 * @param {string} formato string de retorno Ex. ('dd/MM/yyyy')
 * @param {boolean} opcional 
 * @returns {string} 
 */
mesFuturo(formato := 'yyyyMMdd', opcional := false)
{
    MesFuturo := A_MM + 1, AnoFuturo := A_YYYY
    if MesFuturo = 13
        MesFuturo := 1, AnoFuturo := A_YYYY + 1

    DataMesFuturo := Format('{:}{:02}', AnoFuturo, MesFuturo)

    if opcional
        DataMesFuturo := ultimoDiaMes(, DataMesFuturo)

    return FormatTime(DataMesFuturo, formato)
}


/**
 * Retorna o último dia do mês especificado
 * @param {string} formato string de retorno Ex. ('dd/MM/yyyy')
 * @param {string} AnoMes - Se omitido retorna o mês atual
 * @returns {string} 
 */
ultimoDiaMes(formato := 'yyyyMMdd', AnoMes := A_Now)
{
    Ano := SubStr(AnoMes, 1, 4)
    Mes := SubStr(AnoMes, 5, 2)

    MesFuturo := Mes + 1 > 12 ? 1 : Mes + 1
    AnoFuturo := MesFuturo = 1 ? Ano + 1 : Ano

    data := DateAdd(AnoFuturo . Format('{:02}', MesFuturo) . '01', -1, 'days')
    return FormatTime(data, formato)
}

/**
 * Retorna uma array com os dias úteis entre duas datas
 * @param {string} data date-time stamp in the YYYYMMDDHH24MISS format
 * @param {string} data date-time stamp in the YYYYMMDDHH24MISS format
 * @param {string} formato Formato de saída. Ex. dd/MM/yyyy
 * @returns {array} 
 */
listarDiasUteis(data, data2 := A_Now, formato := 'yyyyMMdd')
{
    arr := []
    while DateDiff(data2, data1, 'days') >= 0
    {
        if checarDiaUtil(data1)
            arr.push(FormatTime(data1, formato))

        data1 := DateAdd(data1, 1, 'days')
    }
    return arr
}

/**
 * Verifica se uma data é ou não dia útil
 * @param {string} data date-time stamp in the YYYYMMDDHH24MISS format
 * @returns {boolean} 
 */
checarDiaUtil(data := A_Now)
{
    if FormatTime(data, 'WDay') = 1 or FormatTime(data, 'WDay') = 7 or isHoliday(data)
        return false
    else
        return true
}

/**
 * Retorna o dia útil anterior a uma determinada data
 * @param {string} data A date-time stamp in the YYYYMMDDHH24MISS format
 * @param {string} formato Formato de saída. Ex. dd/MM/yyyy
 * @returns {string}
 */
antecipaData(data := A_Now, formato := 'yyyyMMdd')
{
    ; Se a data coincidir com feriado, volta um dia
    if isHoliday(data)
        data := DateAdd(data, -1, 'days')

    ; Checa o dia da semana da nova data
    diasem := FormatTime(data, 'WDay')

    ; Se cair num domingo, volta dois dias (sexta-feira)
    if (diasem = 1)
        data := DateAdd(data, -2, 'days')

    ; Se cair num sábado, volta um dia (sexta-feira)
    if (diasem = 7)
        data := DateAdd(data, -1, 'days')

    return FormatTime(data, formato)
}

/**
 * Retorna o próximo dia útil a uma determinada data
 * @param {string} data A date-time stamp in the YYYYMMDDHH24MISS format
 * @param {string} formato Formato de saída. Ex. dd/MM/yyyy
 * @returns {string}
 */
postergaData(data := A_Now, formato := 'yyyyMMdd')
{
    ; Se a data coincidir com feriado, avança um dia
    if isHoliday(data)
        data := DateAdd(data, 1, 'days')

    ; Checa o dia da semana da nova data
    diasem := FormatTime(data, 'WDay')

    ; Se cair num domingo, avança um dia (segunda-feira)
    if (diasem = 1)
        data := DateAdd(data, 1, 'days')

    ; Se cair num sábado, avança dois dias (segunda-feira)
    if (diasem = 7)
        data := DateAdd(data, 2, 'days')

    return FormatTime(data, formato)
}

/**
 * Retorna um inteiro maior que zero caso a data seja um feriado
 * @param date A date-time stamp in the YYYYMMDDHH24MISS format
 * @returns {number} 
 */
isHoliday(date)
{
    year := FormatTime(date, 'yyyy')

    ; Feriados móveis de 2021 até 2078
    feriados := '20210215|20210216|20210402|20210603|20220228|20220301|20220415|20220616|'
        . '20230220|20230221|20230407|20230608|20240212|20240213|20240329|20240530|'
        . '20250303|20250304|20250418|20250619|20260216|20260217|20260403|20260604|'
        . '20270208|20270209|20270326|20270527|20280228|20280229|20280414|20280615|'
        . '20290212|20290213|20290330|20290531|20300304|20300305|20300419|20300620|'
        . '20310224|20310225|20310411|20310612|20320209|20320210|20320326|20320527|'
        . '20330228|20330301|20330415|20330616|20340220|20340221|20340407|20340608|'
        . '20350205|20350206|20350323|20350524|20360225|20360226|20360411|20360612|'
        . '20370216|20370217|20370403|20370604|20380308|20380309|20380423|20380624|'
        . '20390221|20390222|20390408|20390609|20400213|20400214|20400330|20400531|'
        . '20410304|20410305|20410419|20410620|20420217|20420218|20420404|20420605|'
        . '20430209|20430210|20430327|20430528|20440229|20440301|20440415|20440616|'
        . '20450220|20450221|20450407|20450608|20460205|20460206|20460323|20460524|'
        . '20470225|20470226|20470412|20470613|20480217|20480218|20480403|20480604|'
        . '20490301|20490302|20490416|20490617|20500221|20500222|20500408|20500609|'
        . '20510213|20510214|20510331|20510601|20520304|20520305|20520419|20520620|'
        . '20530217|20530218|20530404|20530605|20540209|20540210|20540327|20540528|'
        . '20550301|20550302|20550416|20550617|20560214|20560215|20560331|20560601|'
        . '20570305|20570306|20570420|20570621|20580225|20580226|20580412|20580613|'
        . '20590210|20590211|20590328|20590529|20600301|20600302|20600416|20600617|'
        . '20610221|20610222|20610408|20610609|20620206|20620207|20620324|20620525|'
        . '20630226|20630227|20630413|20630614|20640218|20640219|20640404|20640605|'
        . '20650209|20650210|20650327|20650528|20660222|20660223|20660409|20660610|'
        . '20670214|20670215|20670401|20670602|20680305|20680306|20680420|20680621|'
        . '20690225|20690226|20690412|20690613|20700210|20700211|20700328|20700529|'
        . '20710302|20710303|20710417|20710618|20720222|20720223|20720408|20720609|'
        . '20730206|20730207|20730324|20730525|20740226|20740227|20740413|20740614|'
        . '20750218|20750219|20750405|20750606|20760302|20760303|20760417|20760618|'
        . '20770222|20770223|20770409|20770610|20780214|20780215|20780401|20780602|'

    ; Feriados fixos - incluíndo o feriado municipal de Papagaios
    feriados .= year '0101|' year '0421|' year '0501|' year '0907|' year '1012|'
        . year '1102|' year '1115|' year '1225|' year '0120|'

    return InStr(feriados, date)
}
