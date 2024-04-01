/************************************************************************
* @description Exibe uma janela de notificação com o texto desejado
* @author Pedro Henrique C. Xavier
* @date 2024-02-20
* @version 2.1-alpha.8
***********************************************************************/

#Requires AutoHotkey v2.0

/**
* Exibe uma janela de notificação com o texto desejado.
* @param string_alert string a ser exibida.
* @param options string com o esquema de cores, o tempo em seguido da letra 'T'
* e a letra 'N' indicando se é para exibir a janela no canto das notificações
*/
Alert(string_alert, options := '')
{
    ; Aguarda indefinidamente até que outra GUI do mesmo tipo seja fechada antes de exibir a atual
    WinWaitClose('alert_ahk ahk_class AutoHotkeyGUI')

    ; Determina o esquema de cores para ser usado
    switch (RegExMatch(options, '[a-zA-Z]{2,}', &theme) ? theme[0] : 'info'), 'Locale'
    {
        case 'ok'      : cback := 'D1E7DD', ctext := 'c23961D', strTitu := '✔ OK'
        case 'info'    : cback := 'BADDEF', ctext := 'c043A62', strTitu := 'ℹ️ INFO'
        case 'alerta'  : cback := 'FFF3CD', ctext := 'cCAA202', strTitu := '⚠ ALERTA'
        case 'erro'    : cback := 'FEE2E2', ctext := 'cCC0000', strTitu := '☠ ERRO'
        case 'lembrete': cback := 'FFFFFF', ctext := 'c495057', strTitu := '⏰ LEMBRETE'
    }

    GuiA := Gui('-Caption ToolWindow AlwaysOnTop Border', 'alert_ahk')
    GuiA.SetFont(, 'Segoe UI Emoji')
    GuiA.Addtext(ctext, strTitu).SetFont('s14 bold')
    GuiA.Addtext('ys', '✕').OnEvent('Click', (*) => GuiA.Hide())
    GuiA.Addtext('xm section y+20', string_alert).SetFont('s12')
    GuiA.BackColor := cback

    ; Determina o tamanho da GUI, respeitando os valores mínimos e máximos
    MinW := 200, MaxW := 1200, MinH := 20, MaxH := 600
    SizeCtrl := GetTextSize(GuiA['Static3'], string_alert)
    width := MinMax(SizeCtrl.width, MinW, MaxW)
    height := MinMax(SizeCtrl.height, MinH, MaxH)

    GuiA['Static1'].Move(, , width, height + 10)
    GuiA['Static2'].Move(width + 20)
    GuiA['Static3'].Move(, , width, height + 10)

    GuiA.Show('Autosize Hide')

    ; Se especificado nas opções, mostrará a Gui perto das notificações, caso contrário, exibirá centralizado na tela
    if options ~= '\b[nN]\b'
        SurgirDeBaixo()
    else
        GuiA.Show('NoActivate')

    ; Se espeficicado nas opções, fará a Gui ser destruida após os milisegundos definidos
    if RegExMatch(options, '\b[tT](\d+\.*\d*)\b', &timer)
        SetTimer(SumirAosPoucos, timer[1] * -1000)

    SumirAosPoucos()
    {
        Loop (N := 255) - 238
            WinSetTransparent(N += -15, GuiA), Sleep(30)
        GuiA.Destroy()
    }

    SurgirDeBaixo()
    {
        GuiA.GetPos(, , &s_width, &s_height)
        GuiA.Move(A_ScreenWidth - s_width - 2)
        GuiA.Show('NoActivate')
        Y := SysGet(62) - 24

        Loop s_height
            GuiA.Move(, Y--)
    }

    ; Essa função permite que se possa mover a janela clicando em qualquer posição
    if A_AhkVersion ~= '^2.1'
        GuiA.OnMessage(0x201, (*) => PostMessage(0xA1, 2, , , GuiA.hwnd))

    ; Calcula o valor necessário para que a GUI fique entre o tamanho mínimo e o máximo
    MinMax(Value, ValueMin, ValueMax) => Max(ValueMin, Min(ValueMax, Value))

    ; Função para calcular o tamanho que um controle precisará para exibir o texto com a fonte usada no Gui_control
    GetTextSize(textCtrl, text)
    {
        static WM_GETFONT := 0x0031, DT_CALCRECT := 0x400
        hDC := DllCall('GetDC', 'Ptr', textCtrl.Hwnd, 'Ptr')
        hPrevObj := DllCall('SelectObject', 'Ptr', hDC, 'Ptr', SendMessage(WM_GETFONT, , , textCtrl), 'Ptr')
        height := DllCall('DrawText', 'Ptr', hDC, 'Str', text, 'Int', -1, 'Ptr', buf := Buffer(16), 'UInt', DT_CALCRECT)
        width := NumGet(buf, 8, 'Int') - NumGet(buf, 'Int')
        DllCall('SelectObject', 'Ptr', hDC, 'Ptr', hPrevObj, 'Ptr')
        DllCall('ReleaseDC', 'Ptr', textCtrl.Hwnd, 'Ptr', hDC)
        return { width: width * 96 // A_ScreenDPI, height: height * 96 // A_ScreenDPI }
    }
}
