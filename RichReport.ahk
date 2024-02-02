/************************************************************************
 * @description GUI com o controle RichEdit para relatórios RichText
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/24
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0
#Include %A_LineFile% \..\RichEdit.ahk
#Include %A_LineFile% \..\RichEditMenu.ahk
#Include %A_LineFile% \..\cJSON.ahk

class RichReport extends RichEdit
{
    __New(value?)
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

        A_TrayMenu.Add()
        A_TrayMenu.Add('Show GUI', this.GuiShow.Bind(this))
        A_TrayMenu.Default := 'Show GUI'
        A_TrayMenu.ClickCount := 1

        IsSet(value) && this.DebugMode(value)
    }
    /**
     * Insere um cabeçalho pré-formatado
     * @param str Texto a ser exibido
     * @param {number} level
     * @returns {object} Objeto tipo Range
     */
    Header(str, level := 1)
    {
        rng := this.ITextDocument.Range(this.end, this.end)
        rng.Text := str '`n'
        rng.Size := -2 * level + 22
        rng.Bold := this.tomTrue
        rng.SpaceAfter := rng.SpaceBefore := -2 * level + 22
        this.End := rng.End
        return rng
    }
    /**
     * Insere uma lista numerada pré-formatada
     * @param {array} arr_list array com strings
     */
    List(arr_list)
    {
        for item in IsObject(arr_list) ? arr_list : StrSplit(RTrim(arr_list, '`n'), '`n')
            str_list .= Format('`t{}.  {}`n', A_Index, item)

        rng := this.ITextDocument.Range(this.end, this.end)
        rng.Para.ClearAllTabs()
        rng.Para.AddTab(20, this.tomAlignLeft, this.tomSpaces)
        rng.text := str_list
        rng.SetLineSpacing(tomLineSpaceAtLeast := 3, 18)
        this.End := rng.End
        return rng
    }
    /**
     * Insere uma tabela pré-formatada
     * @param {array} table_array array de arrays com conteúdo da tabela
     * @param {array} width array com valores de largura para cada coluna
     * @param {array} align array com strings para o alinhamento de coluna
     */
    Table(table_array, width, align := [])
    {
        nCols := table_array[1].length, twips := 15, color := this.Gray[6]
        SetAlign := Map(), SetAlign.CaseSense := false
        SetAlign.Set('l', this.tomAlignLeft, 'c', this.tomAlignCenter, 'r', this.tomAlignRight)
        align.Length := nCols, align.Default := 'l'

        this.Ctrl.Move(, , 60)
        rng := this.ITextDocument.Range(this.End, this.End)

        for index, row_array in table_array
        {
            try rng.InsertTable(nCols, 1, 0)
            catch
                break

            rng.Move(tomTable := 15, -1)
            row := rng.Row
            Loop nCols
            {
                row.CellIndex := A_Index - 1
                row.SetCellBorderColors(color, color, color, color)
                row.CellWidth := twips * width[A_Index]
                row.CellAlignment := SetAlign['c']

                if index = 1
                    row.CellColorBack := this.Gray[8]
                else if index & 1
                    row.CellColorBack := this.Gray[10]
                else
                    row.CellColorBack := this.White
            }
            rng.Move(tomRow := 10, -1)
            for value in row_array
            {
                rng.Text := value ?? ''
                rng.Alignment := SetAlign[ align[A_Index] ]
                rng.SpaceAfter := rng.SpaceBefore := 2
                (index = 1) && rng.Bold := this.tomTrue
                rng.Move(tomCell := 12, 1)
            }
            ; Aplicação das modificações da estrutura da tabela
            row.Apply(1, 0)
            this.End += rng.End + 2
        }
    }
    /**
     * Ajusta o tamanho do controle e invoca o comando Show da Gui hospedeira
     * Opcionalmente definindo o modo edição ou somente leitura para o controle
     * @param {string|integer} options se o parâmetro for um inteiro, será a razão do tamanho da tela.
     * Se for uma string, será as mesmas opções para o método Gui.Show()
     * @param {boolean} ReadOnly
     */
    Show(options := 60, ReadOnly := true)
    {
        if IsInteger(options)
        {
            this.Ctrl.Move(, , A_ScreenWidth * options / 100, A_ScreenHeight * options / 100)
            this.GuiR.Show('AutoSize')
        }
        else
            this.GuiR.Show(options)

        this.ReadOnly(ReadOnly)
        this.ShowMenu()
    }
    /**
     * Minimiza ou Maximiza a janela do relatório clicando no ícone de bandeija.
     */
    GuiShow(*)
    {
        win := 'AHK_id' this.GuiR.Hwnd
        WinGetMinMax(win) ? WinRestore(win) : WinMinimize(win)
    }
    /**
     * Mostra o valor com fonte monoespaçada e em fundo cinza.
     * Se o valor for um objeto, usará a função JSON.Dump()
     * @param {value} Number, String, Map or Array
     */
    DebugMode(value)
    {
        IsObject(value) && value := JSON.Dump(value)
        this.SetBkgnColor(this.color['#F0F0F0'])
        this.Text(value).Name := 'Consolas'
        this.Show(, false)
    }
}
