/************************************************************************
 * @description Técnica para enviar um cliques do mouse para um controle específico
 * Permite acionamento sem que o mouse seja movido e funciona em janelas em segundo plano
 * As coordenadas são relativas ao controle e não a janela. É necessario calcular.
 * 0x0201 -> LButton Down, 0x0202 -> LButton Up, 0x0001 -> MousePressionado
 * @file SendClick.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/09/27
 * @version 2.0.10
 ***********************************************************************/

#Requires AutoHotkey v2.0

/**
 * Envia cliques do mouse para um controle específico
 * @param {number} X Coordenada X
 * @param {number} Y Coordenada Y
 * @param {hwnd} CtrlID ID do controle
 * @param {hwnd} WinID  ID da janela
 */
SendClick(X, Y, CtrlID, WinID := WinActive('A'))
{
    for msg in [0x0201, 0x0202]
        PostMessage(msg, 0x0001, lParam(X, Y), CtrlID, WinID)

    lParam(X, Y) => Y << 16 | X
}
