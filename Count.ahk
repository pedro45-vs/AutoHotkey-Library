﻿/************************************************************************
 * @description Cria um contador para calculo de benchmark
 * @file Count.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/08/21
 * @version 2.0.5
 ***********************************************************************/
#Requires AutoHotkey v2.0

/**
 * Implementa uma forma mais prática de criar contadores para benchmark
 * pode ser invocado com os métodos estáticos ou por instância de classe
 * @method start() inicia o contador
 * @prop value retorna o valor em milisegundos decorridos
 */
class count
{
	static CounterBefore := 0, CounterAfter := 0, freq := 0
	static value
	{
		get {
			DllCall('QueryPerformanceCounter', 'Int64*', &CounterAfter := 0)
			return Round((CounterAfter - this.CounterBefore) / this.freq * 1000)
		}
	}
	; Inicia o contador
	static start()
	{
		DllCall('QueryPerformanceFrequency', 'Int64*', &freq := 0)
		DllCall('QueryPerformanceCounter', 'Int64*', &CounterBefore := 0)
		this.CounterBefore := CounterBefore, this.freq := freq
	}
}
