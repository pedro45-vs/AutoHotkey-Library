/************************************************************************
 * @description Biblioteca com funções para uso no Windows Explorer
 * @file Explorer.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/08/25
 * @version 2.0.5
 ***********************************************************************/

#Requires AutoHotkey v2.0

/**
 * Recupera o caminho da pasta aberta no Explorer ativo
 * @param {Hwnd} ActiveHwnd da janela
 * @returns {string} Caminho da pasta
 */
GetExplorerPath(ActiveHwnd)
{
    if WinActive('ahk_exe explorer.exe')
    {
        Texto := WinGetText('ahk_id' ActiveHwnd)
        Try
        {
            RegExMatch(Texto, 'Endereço:\s(.*)', &folder)
            Switch folder[1]
            {
                case 'Downloads': pasta := 'C:\Users\' A_UserName '\Downloads'
                case 'Área de Trabalho': pasta := A_Desktop
                Default: pasta := folder[1]
            }
            return pasta
        }
    }
}
/**
 * Abre uma pasta do explorer e a coloca em primeiro plano
 * @param {string} path pasta do explorer
 */
openfolder(path)
{
    ult := StrSplit(path, '\')
    tit := ult[ult.length]
    RunWait(path)
    WinWait(tit, , 1)
    WinActivate(tit)
}
/**
 * Executa e ativa um programa ou script opcionalmente movendo-o para
 * O monitor em que o mouse se encontra
 * @param {string} program Caminho completo do programa a ser executado
 * @param {string} WinTitle Parâmetro opcional para melhorar a correspondência
 * @param {boolean} Move Se verdadeiro, moverá o programa para o monitor ativo
 */
RunAct(program, WinTitle:='', Move:=false)
{
    SplitPath(program, &exe)
    WinMatch := WinTitle ? WinTitle : 'ahk_exe ' exe

    if not WinExist(WinMatch)
        Run(program)
    else
        WinActivate()

    if Move
    {
        CoordMode('Mouse')
        MouseGetPos(&X)
		if not WinWait(WinMatch,, 1)
			return
		WinGetPos(,, &OutWidth, &OutHeight)
        movX := (A_ScreenWidth - OutWidth) // 2
        movY := (A_ScreenHeight - OutHeight) // 2

		if X > A_ScreenWidth
            movX += A_ScreenWidth

		WinMove(movX,movY,,, WinMatch)
    }
}
/**
 * Resolve os caminhos relativos em absolutos
 * @param {string} relative_path
 */
PathCanonicalize(relative_path)
{
    VarSetStrCapacity(&out, 260)
    if DllCall('shlwapi\PathCanonicalize', 'str', out, 'str', relative_path)
        return out
}
/**
 * Chega se uma pasta está vazia ou não
 * @param {string} path
 */
PathIsDirectoryEmpty(path) => DllCall('shlwapi\PathIsDirectoryEmpty', 'str', path)
