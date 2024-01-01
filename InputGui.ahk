/************************************************************************
 * @description Exibe uma Gui para inserir dados do ClipBoard
 * Fiz isso porque o InputBox bloqueia a thread enquanto está ativa
 * @file InputGui.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/08/23
 * @version 2.0.5
 ***********************************************************************/

#Requires AutoHotKey v2.0

InputGui(coord?)
{
	GuiI := Gui('Resize MinSize MaxSizex26 -MaximizeBox -MinimizeBox AlwaysOnTop', 'Copiar para o ClipBoard')
	GuiI.OnEvent('Close', (*) => GuiI.Destroy())
	GuiI.OnEvent('Escape', (*) => GuiI.Destroy())
	GuiI.OnEvent('Size', Gui_Size)
	GuiI.Add('Edit', 'w250')
	GuiI.Add('Button', 'w80 default', 'OK').OnEvent('Click', Clip)
	GuiI['Edit1'].SetFont('s10', 'JetBrains Mono')
	GuiI.Show(coord ?? '')

	; Essa função permite que se possa mover a janela clicando em qualquer posição
	; OnMessage(0x201, (wParam, lParam, msg, hwnd) => PostMessage(0xA1, 2, , , hwnd))

	Gui_Size(GuiObj, MinMax, Width, Height)
	{
		GuiI['Edit1'].Move(, , Width - 20)
		GuiI['Button1'].Move((Width - 100) >> 1) ; Equivale a dividir por 2
	}

	Clip(*)
	{
		A_Clipboard := GuiI['Edit1'].Text
		GuiI.Destroy()
	}
}
