/************************************************************************
 * @description GUI com o controle RichEdit para relatórios RichText
 * @author Pedro Henrique C. Xavier
 * @date 2023/12/22
 * @version 2.0.10
 ***********************************************************************/

#Requires AutoHotkey v2.0
#Include %A_LineFile% \..\RichEdit.ahk
#Include %A_LineFile% \..\RichEditMenu.ahk

class RichReport extends RichEdit
{
    __New()
    {
        prevIconFile := A_IconFile, prevIconNumber := A_IconNumber
        TraySetIcon(A_LineFile '\..\..\icons\report.ico')
        this.GuiR := Gui('Resize', 'RichReport: ' A_ScriptName)
        (prevIconFile) && TraySetIcon(prevIconFile, prevIconNumber)
        this.GuiR.OnEvent('Size', (o, m, w, h) => this.Ctrl.Move(,, w, h))
        this.GuiR.OnEvent('Close', (*) => ExitApp())
        this.GuiR.SetFont('s10', 'Segoe UI')
        this.GuiR.MarginX := this.GuiR.MarginY := 0
        super.__New(this.GuiR, 'vRich VScroll HScroll ReadOnly')
        this.SetMargins(20, 20)
    }
}