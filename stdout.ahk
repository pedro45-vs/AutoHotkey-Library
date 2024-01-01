/************************************************************************
 * @description Retorna o resultado de um comando CMD do stdout com opção
 * de streaming do resultado.
 * @author Diversos, com modificações minhas
 * @date 2023/12/18
 * @version 2.0.10
 ***********************************************************************/

#Requires AutoHotKey v2.0

/**
 * Retorna o resultado de um comando executando pelo CMD
 * @param {string} sCmd Comando a ser executado
 * @param {func} Callback(output, index) função que processa o resultado a cada atualização
 * @param {string} CP Encoding de retorno, normalmente CP0 ou CP850
 * @returns {string} Texto do standart output
 */
stdout(sCmd, Callback := '', CP := '850')
{
    DllCall('CreatePipe', 'UIntP', &hPipeRead := 0, 'UIntP', &hPipeWrite := 0, 'UInt', 0, 'UInt', 0)
    DllCall('SetHandleInformation', 'UInt', hPipeWrite, 'UInt', 1, 'UInt', 1)

    STARTUPINFO := Buffer(104, 0)                ; STARTUPINFO          ;  http://goo.gl/fZf24
    NumPut('UInt', 68, STARTUPINFO, 0)           ; cbSize
    NumPut('UInt', 0x100, STARTUPINFO, 60)       ; dwFlags    =>  STARTF_USESTDHANDLES = 0x100
    NumPut('UInt', hPipeWrite, STARTUPINFO, 88)  ; hStdOutput
    NumPut('UInt', hPipeWrite, STARTUPINFO, 96)  ; hStdError

    PROCESS_INFORMATION := Buffer(32)           ; PROCESS_INFORMATION  ;  http://goo.gl/b9BaI

    if not DllCall('CreateProcess', 'UInt', 0, 'UInt', StrPtr(sCmd), 'UInt', 0, 'UInt', 0,  ;  http://goo.gl/USC5a
        'UInt', 1, 'UInt', 0x08000000, 'UInt', 0, 'UInt', 0, 'UInt', STARTUPINFO.ptr, 'UInt', PROCESS_INFORMATION.ptr)
    {
        DllCall('CloseHandle', 'UInt', hPipeWrite)
        DllCall('CloseHandle', 'UInt', hPipeRead)
        DllCall('SetLastError', 'Int', -1)
        return ''
    }

    hProcess := NumGet(PROCESS_INFORMATION, 0, 'UInt')
    hThread := NumGet(PROCESS_INFORMATION, 8, 'UInt')

    DllCall('CloseHandle', 'UInt', hPipeWrite)
    Buf := Buffer(4096), nSz := 0

    while DllCall('ReadFile', 'UInt', hPipeRead, 'UPtr', Buf.ptr, 'UInt', 4094, 'UIntP', &nSz, 'Int', 0)
    {
        tOutput := StrGet(Buf, nSz, CP)
        IsObject(Callback) ? Callback(tOutput, A_Index) : sOutput .= tOutput
    }

    DllCall('GetExitCodeProcess', 'UInt', hProcess, 'UIntP', &ExitCode := 0)
    DllCall('CloseHandle', 'UInt', hProcess)
    DllCall('CloseHandle', 'UInt', hThread)
    DllCall('CloseHandle', 'UInt', hPipeRead)
    DllCall('SetLastError', 'UInt', ExitCode)

    return IsObject(Callback) ? Callback('', 0) : sOutput
}
