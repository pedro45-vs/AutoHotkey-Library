/************************************************************************
 * @description RichEdit para exibição de relatórios com formatação
 * @author Pedro Henrique C. Xavier
 * @date 2023/12/18
 * @version 2.0.10
 ***********************************************************************/

#Requires AutoHotkey v2.0
#Include %A_LineFile% \..\RichEdit.ahk

RichReport()
{
    GuiR := Gui('Resize -DPIScale', 'RichReport: ' A_ScriptName)
    GuiR.SetFont('s10', 'Segoe UI')
    GuiR.MarginX := GuiR.MarginY := 0
    GuiR.OnEvent('Close', (*) => GuiR.Destroy())
    Rich := RichEdit(GuiR, 'vRich VScroll HScroll ReadOnly')
    GuiR.OnEvent('Size', (o, m, w, h) => Rich.Ctrl.Move(,, w, h))
    
    Rich.SetMargins(20, 20)
    Rich.Ctrl.OnEvent('ContextMenu', RichContextMenu)
    return Rich

    /**
     * Função Callback disparada quando se clica no Menu de Contexto
     */
    RichContextMenu(*)
    {
        CM := Menu()
        CM.Add('Copiar', (*)=> Send('^c'))
        CM.Add('Colar', (*)=> Send('^v'))
        CM.Add('Salvar', SaveDialog)
        CM.Show()
    }
    
    SaveDialog(*)
    {
        if not select_file := FileSelect('S16', A_MyDocuments '\RichReport.rtf', 'Salvar arquivo', 'Rich Text Format(*.rtf)')
            Exit()

        if not select_file ~= '[Rr][Tt][Ff]$'
            select_file .= '.rtf'

        FileExist(select_file) && FileDelete(select_file)
        Rich.ITextDocument.Save(select_file, tomRTF := 0x1, codePage := 1200)
    }
}