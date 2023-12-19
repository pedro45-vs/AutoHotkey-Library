/************************************************************************
 * @description Converte TimeStamp AutoHotkey para o UnixTime e vice-versa
 * @file UnixTime.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/08/21
 * @version 2.0.5
 ***********************************************************************/

#Requires AutoHotkey v2.0

/**
 * Converte TimeStamp AutoHotkey para UnixTime
 * @param {string} time a YYYYMMDDHH24MISS timestamp 
 * @returns {number} UnixTime
 */
TimetoUnix(time)
{
	unix := DateAdd(time, 3, 'hours')
	unix := DateDiff(unix, 19700101000000, 'seconds')
	return unix
}

/**
 * Converte TimeStamp UnixTime para AutoHotkey
 * @param {string} unix a UnixTime
 * @param {string} formato Formato de saída
 * @returns {string} string formatada
 */
TimeFromUnix(unix, formato := 'yyyy-MM-ddTHH:mm:ss')
{
	date := DateAdd(19700101000000, unix, 'seconds')
	date := DateAdd(date, -3, 'hours')
	return FormatTime(date, formato)
}
