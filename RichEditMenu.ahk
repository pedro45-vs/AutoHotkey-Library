/************************************************************************
 * @description Cria Menu de Contexto para o RichEdit com diversas funções
 * Para criar o menu, inclua essa lib e chame o método Instancia.CriarMenu()
 * @author Pedro Henrique C. Xavier
 * @date 2023/12/21
 * @version 2.0.10
 ***********************************************************************/

RichEdit.Prototype.Base := RichEditMenu()

class RichEditMenu
{
    /**
     * Ativa o Menu de contexto;
     * Cria a Gui da Caixa de diálogo para pesquisa.
     */
    ShowMenu()
    {
        this.Ctrl.OnEvent('ContextMenu', this.RichContextMenu.Bind(this))

        this.GuiP := GuiP := Gui('-MinimizeBox', 'Localizar', this)
        GuiP.Opt('+Owner' this.Ctrl.Hwnd)
        GuiP.OnEvent('Close', (*) => GuiP.Hide())
        GuiP.OnEvent('Escape', (*) => GuiP.Hide())
        GuiP.SetFont(, 'Segoe UI')
        GuiP.AddEdit('vPesq w280')
        GuiP.AddButton('w110 x+10 default', 'Localizar Próxima').OnEvent('Click', this.Search.Bind(, GuiP, this))
        GuiP.AddCheckbox('vCheck w280 xm', 'Diferenciar maíusculas de minúsculas')
        GuiP.AddButton('w110 x+10', 'Localizar Anterior').OnEvent('Click', this.Search.Bind(, GuiP, this))

        ; Atalhos personalizados para alguns itens do Menu.
        ; Copiar e Colar já fazem parte do controle RichEdit.
        HotIfWinActive('ahk_id' this.Ctrl.Gui.hwnd)
        Hotkey('^s', (*) => this.SaveDialog())
        Hotkey('^f', (*) => this.ShowGuiSearch())
    }
    /**
     * @params {callback}
     * Mostra o Menu de Contexto
     */
    RichContextMenu(*)
    {
        CM := Menu()
        CM.Add('Copiar `tCtrl+C', (*)=> Send('^c'))
        CM.Add('Colar `tCtrl+V', (*)=> Send('^v'))
        CM.Add('Selecionar tudo `tCtrl+A', (*)=> Send('^a'))
        CM.Add()
        CM.Add('Localizar `tCtrl+F', this.ShowGuiSearch.Bind(this))
        CM.Add('Salvar `tCtrl+S', this.SaveDialog.Bind(this))
        CM.Add('Abrir com Wordpad', this.OpenWithWordPad.Bind(this))
        CM.Show()
    }
    /**
     * @params {callback}
     * Mostra caixa de diálogo para salvar conteúdo do controle em arquivo RTF
     */
    SaveDialog(*)
    {
        if not select_file := FileSelect('S16', A_MyDocuments '\RichText.rtf', 'Salvar arquivo', 'Rich Text Format(*.rtf)')
            Exit()

        if not select_file ~= '[Rr][Tt][Ff]$'
            select_file .= '.rtf'

        FileExist(select_file) && FileDelete(select_file)
        this.ITextDocument.Save(select_file, tomRTF := 0x1, codePage := 1200)
    }
    /**
     * @params {callback}
     * Mostra caixa de pesquisa para localizar palavras no controle RichEdit
     */
    ShowGuiSearch(*)
    {
        this.Ctrl.Gui.GetPos(&X, &Y, &Width, &Height)
        this.GuiP.Show('X ' X + Width / 2 - 218 ' Y ' Y + Height / 2 - 51)
        this.occurrence := 0
    }
    /**
     * @params {callback}
     * Executa a função de localizar, selecionando a palavra encontrada no controle RichEdit
     */
    Search(GuiObj, Rich, Info)
    {
        (needle := GuiObj['Pesq'].Text) || Exit()
        case_sense := GuiObj['Check'].Value ? 'On' : 'Locale'
        inc := this.Text = 'Localizar Próxima' ? 1 : -1

        if Rich.occurrence <= 1 and inc = -1
            Rich.occurrence := 1
        else
            Rich.occurrence += inc

        haystack := Rich.ITextDocument.Range(0, Rich.End).text
        if pos := InStr(haystack, needle, case_sense, 1, Rich.occurrence)
            Rich.ITextDocument.Range(pos -= 1, pos + StrLen(needle)).Select()
        else
        {
            MsgBox('Não foi possível localizar', 'Localizar', 'Icon! 4096 T1')
            Rich.occurrence := 0
        }
    }
    /**
     * @params {callback}
     * Salva um arquivo RTF temporário e o abre com o Wordpad
     */
    OpenWithWordPad(*)
    {
        FileExist(A_Temp '\temp.rtf') && FileDelete(A_Temp '\temp.rtf')
        this.ITextDocument.Save(A_Temp '\temp.rtf', tomRTF := 0x1, codePage := 1200)
        Run Format('"{}" "{}"', 'C:\Program Files (x86)\Windows NT\Accessories\wordpad.exe', A_Temp '\temp.rtf')     
    }
}