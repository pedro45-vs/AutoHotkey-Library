/************************************************************************
 * @description Retorna a saudação apropriada conforme a hora do dia
 * @author Pedro Henrique C Xavier
 * @date 2023-08-22
 * @version 2.0.5
 ***********************************************************************/

/**
 * Retorna a saudação apropriada conforme a hora do dia
 * @param {Integer} hr 
 * @param {string} sep 
 * @returns {string} 
 */
saudacao(hr:=A_Hour, sep:=',')
{
	if hr >= 0 and hr <= 5
		sds := 'Boa madrugada'
	if hr > 5 and hr <= 11
		sds := 'Bom dia'
	if hr > 11 and hr <= 17
		sds := 'Boa tarde'
	if hr > 17 and hr <= 23
		sds := 'Boa noite'
	return sds ','
}
