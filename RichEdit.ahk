/************************************************************************
 * @description Controle RichEdit para criação e exibição de RichText
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/03
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0
#DllLoad 'Msftedit.dll'

/**
 * TOM Constants diponível em TOM.h
 * RichEdit Constants disponível em RichEdit.h
 *
 * Interface ITextDocument
 *   https://learn.microsoft.com/pt-br/windows/win32/api/tom/nn-tom-itextdocument2
 *
 * PRINCIPAIS MÉTODOS E PROPRIEDADES
 *
 * Interface ITextRange
 *   ScrollIntoView(tomEnd := 0 | tomStart := 32)
 *
 * Interface ITextPara
 *   SpaceAfter     - Insere espaço após o parágrafo (float)
 *   SpaceBefore    - Insere espaço antes do parágrafo (float)
 *   Alignment      - Alinha o parágrafo entre esquerda, direita, centralizado (tomAlignment)
 *   LineSpacing    - Alterna o espaço entre linhas (float)
 *
 * Interface ITextFont
 *   Name           - Define o nome da fonte. (string)
 *   Size           - Define o tamanho da fonte (float)
 *   ForeColor      - Define a cor da fonte (RGB Value)
 *   BackColor      - Define a cor de fundo da fonte (RGB Value)
 *   Bold           - Define o estilo da fonte como negrito (tomBoolean)
 *   Italic         - Define o estilo da fonte como itálico (tomBoolean)
 *   Underline      - Define o estilo da fonte como sublinhado (tomBoolean)
 *   StrikeThrough  - Define o estilo da fonte como riscado (tomBoolean)
 *   Shadow         - Define o estilo da fonte como sombreamento (tomBoolean)
 *   Spacing        - Define o espaço entre caracteres (float)
 *   AllCaps        - Define o texto como maiúsculas (tomBoolean)
 *   SmallCaps      - Define o texto como minúsculas (tomBoolean)
 */

class RichEdit
{
    ; DEFAULT STYLES
    STYLES := ES_SAVESEL := 0x00008000 | ES_NOHIDESEL := 0x00000100
        | ES_MULTILINE := 0x00000004 | ECO_WANTRETURN := 0x00001000

    ; TOM CONSTANTS
    tomFalse := 0, tomTrue := -1, tomToggle := -9999998, tomUndefined := -9999999,
    tomAlignLeft := 0, tomAlignCenter := 1, tomAlignRight := 2,
    tomAlignDecimal := tomAlignJustify := 3, tomAlignBar := 4,
    tomSpaces := 0, tomDots := 1, tomDashes := 2, tomLines := 3,
    tomLowerCase := 0, tomUpperCase := 1, tomTitleCase := 2,
    tomSentenceCase := 4, tomToggleCase := 5,
    end := 0

    ; PALETA DE CORES
    Lum := [15, 25, 35, 45, 55, 65, 75, 85, 90, 95]
    Black := 0x000000, LightGray := 0xF0F0F0, White := 0xFFFFFF
    Color[str] => Integer('0x' SubStr(str, -2, 2) SubStr(str, -4, 2) SubStr(str, -6, 2))
    Gray[index := 5]    => this.hsl(  0,  0, this.Lum[index])
    Brown[index := 5]   => this.hsl( 15, 25, this.Lum[index])
    Slate[index := 5]   => this.hsl(200, 18, this.Lum[index])
    Red[index := 5]     => this.hsl(  0, 75, this.Lum[index])
    Orange[index := 5]  => this.hsl( 30, 75, this.Lum[index])
    Yellow[index := 5]  => this.hsl( 60, 75, this.Lum[index])
    Lime[index := 5]    => this.hsl( 90, 75, this.Lum[index])
    Green[index := 5]   => this.hsl(120, 75, this.Lum[index])
    Cyan[index := 5]    => this.hsl(180, 75, this.Lum[index])
    Blue[index := 5]    => this.hsl(240, 75, this.Lum[index])
    Purple[index := 5]  => this.hsl(270, 75, this.Lum[index])
    Magenta[index := 5] => this.hsl(300, 75, this.Lum[index])
    Pink[index := 5]    => this.hsl(330, 75, this.Lum[index])

    ; Recupera ou define uma Interface ITextFont
    ITextFont
    {
        get => this.ITextDocument.Documentfont
        set => this.ITextDocument.Documentfont := value
    }
    /**
     * Retorna um handle para o controle RichText
     * @param {GuiObj} Gui
     * @param {string} Opções para o controle
     * @returns {COMObject}
     * https://www.autohotkey.com/boards/viewtopic.php?f=82&t=121656&hilit=rich
     */
    __New(GuiObj, Options)
    {
        this.Ctrl := GuiObj.AddCustom('ClassRICHEDIT50W ' Options ' ' this.Styles)

        static EM_GETOLEINTERFACE := 0x43C, VT_DISPATCH := 9, VT_UNKNOWN := 13, F_OWNVALUE := 1,
            IID_ITextDocument2 := '{01C25500-4268-11D1-883A-3C8B00C10000}'

        DllCall('SendMessage', 'Ptr', this.Ctrl.hwnd, 'UInt', EM_GETOLEINTERFACE, 'Ptr', 0,
            'PtrP', IRichEditOle := ComValue(VT_UNKNOWN, 0))

        ITextDocument := ComObjQuery(IRichEditOle, IID_ITextDocument2)
        this.ITextDocument := ComValue(VT_DISPATCH, ITextDocument, F_OWNVALUE)

        this.Ctrl.OnNotify(EN_LINK := 0x70B, this.ClickLink.Bind(, this.ITextDocument))
    }
    /**
     * Define a cor do plano de fundo do controle RichEdit
     * @param {integer} BackColor valor RGB
     */
    SetBkgnColor(BackColor)
    {
        SendMessage(WM_USER := 0x0400 | EM_SETBKGNDCOLOR := 67, 0, BackColor, this.Ctrl)
    }
    /**
     * Define as margens do Controle RichEdit
     * @param {integer} Left
     * @param {integer} Right
     */
    SetMargins(Left := 0, Right := 0)
    {
        SendMessage(EM_SETMARGINS := 0x00D3, EC_LEFTMARGIN := 0x0001, Left, this.Ctrl)
        SendMessage(EM_SETMARGINS := 0x00D3, EC_RIGHTMARGIN := 0x0002, Right << 16, this.Ctrl)
    }
    /**
     * Mostra a GUI com a opção de habilitar ou não a edição
     * @param {bool}
     */
    ReadOnly(bool := true)
    {
        tomNormalCaret := 0, tomNullCaret := 0x2
        this.ITextDocument.CaretType := bool ? tomNullCaret : tomNormalCaret
        SendMessage(EM_SETREADONLY := 0x00CF, bool, , this.Ctrl)
    }
    /**
     * Ajusta o tamanho do controle e invoca o comando Show da Gui hospedeira
     * Opcionalmente definindo o modo edição ou somente leitura para o controle
     * @param {string|integer} options se o parâmetro for um inteiro, será a razão do tamanho da tela.
     * Se for uma string, será as mesmas opções para o método Gui.Show()
     * @param {boolean} ReadOnly
     */
    Show(options := 'AutoSize', ReadOnly := true)
    {
        if IsInteger(options)
        {
            this.Ctrl.Move(, , A_ScreenWidth * options / 100, A_ScreenHeight * options / 100)
            this.Ctrl.Gui.Show('AutoSize')
        }
        else
            this.Ctrl.Gui.Show(options)

        this.ReadOnly(ReadOnly)
    }
    /**
     * Limpa o controle RichEdit
     */
    Clear() => this.ITextDocument.New()
    /**
     * Carrega um arquivo rtf para ser exibido no controle
     * @param {string} fileName
     * @param {number} codePage
     */
    Load(fileName, codePage := 1200) => this.ITextDocument.Open(fileName, 0, codePage)
    /**
     * Salva o conteúdo do controle em arquivo
     * @param {string} FileName
     */
    Save(FileName := A_Desktop '\RichText.rtf')
    {
        FileExist(FileName) && FileDelete(FileName)
        this.ITextDocument.Save(FileName, tomRTF := 0x1, codePage := 1200)
    }
    /**
     * Exibe o texto com a opção de alinhamento e retorna um objeto do tipo ITextDocument.Range.Font
     * @param {string} str Texto a ser exibido
     * @returns {object} Objeto tipo Range
     */
    Text(str)
    {
        rng := this.ITextDocument.Range(this.end, this.end)
        rng.Text := str
        this.End := rng.End
        return rng
    }
    /**
     * Insere um espaço vertical entre parágrafos
     * @param {number} num tamanho do espaço
     */
    Space(num := 1)
    {
        rng := this.ITextDocument.Range(this.end, this.end)
        rng.Text := '`n'
        rng.Para.SpaceAfter := num
        this.End := rng.End
        return rng
    }
    /**
     * Insere uma linha com o comprimento passado
     * @param {string} len
     * @returns {object} objeto tipo Range
     */
    Line(len, char := 1)
    {
        switch char
        {
            case 1: char := '▔'
            case 2: char := '━'
            case 3: char := '▁'
            case 4: char := '▀'
            case 5: char := '■'
            case 6: char := '▄'
            case 7: char := '═'
        }
        rng := this.ITextDocument.Range(this.end, this.end)
        rep := Format('{:0' len '}', 0)
        rng.Text := StrReplace(rep, 0, char) '`n'
        this.End := rng.End
        return rng
    }
    /**
     * Insere o texto com padrão de cores intercalada
     * @param {string} str texto separado por quebra de linha
     * @param {number} color1 cor a ser usada
     * @param {number} color2 cor a ser usada
     * @returns {object} objeto tipo Range
     */
    ColorLines(str, color1 := this.LightGray, color2 := this.White)
    {
        end := this.end
        loop parse str, '`n', '`r'
        {
            if A_Index & 1
                this.text(A_LoopField '`n').BackColor := color2
            else
                this.text(A_LoopField '`n').BackColor := color1
        }
        return this.ItextDocument.Range(end, this.End)
    }
    /**
     * Adiciona as paradas de tabulação e o alinhamento
     * @param {string} caracteres para espaçamento da tabulação: Spaces, Dots, Dashes, Lines
     * @param {variadic} tabs paramentro de parada seguida do alinhamento.
     * Ex. AddTabs(, 240, 'right', 300, 'left')
     */
    AddTabs(tbLeader := this.tomSpaces, tabs*)
    {
        rng := this.ITextDocument.Range(this.end, this.end)
        rng.Spacing := 1
        rng.Para.ClearAllTabs()
        this.Tabs := []
        for param in tabs
        {
            if A_Index & 1
                tbPos := param
            else
            {
                switch param, false
                {
                    case 'l', 'left': tbAlign := this.tomAlignLeft
                    case 'c', 'center': tbAlign := this.tomAlignCenter
                    case 'r', 'right': tbAlign := this.tomAlignRight
                    case 'd', 'decimal': tbAlign := this.tomAlignDecimal
                    default: tbAlign := this.tomAlignLeft
                }
                this.Tabs.Push([tbPos, tbAlign])
                rng.Para.AddTab(tbPos, tbAlign, tbLeader)
            }
        }
    }
    /**
     * Formata o resultado da pesquisa com as propriedades selecionadas
     * @param {string} needleRegEx Regex da pesquisa
     * @param {object} obj Objeto com as propriedades do ITextFont que serão alteradas
     */
    SetFormatInSearch(needleRegEx, obj)
    {
        haystack := this.ITextDocument.Range(0, this.End).text
        while pos := RegExMatch(haystack, needleRegEx, &re, pos ?? 1)
        {
            style := this.ITextDocument.Range(pos - 1, pos + re.len - 1).Font
            for prop, value in obj.OwnProps()
                style.%prop% := value

            this.ITextDocument.Range(pos - 1, (pos += re.len) - 1).Font := style
        }
    }
    /**
     * Insere uma imagem ao controle RichEdit.
     * Conversão para HIMETRIC = PIXEL * 2540 / 96
     * @param {string} filepath Caminho completo do arquivo
     * @param {integer} width Largura em pixels
     * @param {integer} height Altura em pixels
     * @param {integer} Align Alinhamento da imagem
     */
    InsertImage(filepath, width, height, Align := this.tomAlignLeft)
    {
        rng := this.ITextDocument.Range(this.end, this.end)
        imgbuf := FileRead(filepath, 'RAW')
        pIStream := DllCall('Shlwapi\SHCreateMemStream', 'Ptr', imgbuf, 'UInt', imgbuf.size, 'Ptr')
        this.Space()
        rng.InsertImage(width *= 2540 / 96, height *= 2540 / 96, 0, 0, '', ComValue(VT_UNKNOWN := 13, pIStream))
        this.Space().Alignment := Align
    }
    /**
     * Insere uma linha de tabela, no formato desejado
     * @param {array} arr_str array com strings
     * @param {object} obj objeto com as seguintes propriedades (todas são opcionais)
     * @prop {integer} Height Altura das colunas
     * @prop {integer} CellBorderWidth Espessura das bordas
     * @prop {integer} CellBorderColor Cor das bordas no valor RGB
     * @prop {integer|array} CellWidth Integer ou Array com as larguras das colunas
     * @prop {integer|array} CellColorBack Integer ou Array com a cor de fundo das colunas
     * @prop {integer|array} CellAlignment Integer ou Array com o alinhamento vertical das colunas
     * @prop {ITextFont|array} Font Objetos ou Array com os objetos ITextFont para cada coluna
     * @prop {integer|array} Alignment Integer ou Array com o alinhamento horizontal das colunas
     */
    InsertRow(arr_str, obj := {})
    {
        this.Ctrl.Move(, , 60)
        rng := this.ITextDocument.Range(this.End, this.End)
        rng.InsertTable(arr_str.length, 1, 0)
        rng.Move(tomTable := 15, -1)
        row := rng.Row

        ; Alterações na estrutura da Tabela, navegação pelo método CellIndex
        if obj.HasProp('Height')
            row.Height := obj.Height * 15

        loop arr_str.length
        {
            row.CellIndex := A_Index - 1
            if obj.HasProp('CellBorderWidth') && width := obj.CellBorderWidth * 15
                row.SetCellBorderWidths(width, width, width, width)

            if obj.HasProp('CellBorderColor') && color := obj.CellBorderColor
                row.SetCellBorderColors(color, color, color, color)

            if obj.HasProp('CellWidth')
                row.CellWidth := (IsInteger(obj.CellWidth) ? obj.CellWidth : obj.CellWidth[A_Index]) * 15

            if obj.HasProp('CellAlignment')
                row.CellAlignment := IsInteger(obj.CellAlignment) ? obj.CellAlignment : obj.CellAlignment[A_Index]

            if obj.HasProp('CellColorBack')
                row.CellColorBack := IsInteger(obj.CellColorBack) ? obj.CellColorBack : obj.CellColorBack[A_Index]
        }
        ; Aplicação das modificações da estrutura da tabela
        row.Apply(1, 0)

        ; Alterações nas marcas de tabulação, navegação pelo método Move(tomRow)
        rng.Move(tomRow := 10, -1)
        if this.HasProp('Tabs')
            for tab in this.tabs
                rng.AddTab(tab[1], tab[2], 0)

        ; Alterações no conteúdo da Tabela, navegação pela célula Move(tomCell)
        for field in arr_str
        {
            if obj.HasProp('Font')
                rng.Font := (obj.Font is Array) ? obj.Font[A_Index] : obj.Font

            if obj.HasProp('Alignment')
                rng.Alignment := (obj.Alignment is Array) ? obj.Alignment[A_Index] : obj.Alignment

            rng.text := field, rng.Move(tomCell := 12, 1)
        }
        this.End += rng.End + 2
        return rng
    }
    /**
     * Converte a cor no padrão hsl para o valor RGB
     * @param {number} hue valor da cor entre de 0 e 360
     * @param {number} sat saturação entre 0 e 100
     * @param {number} light luminosidade entre 0 e 100
     * @returns {number} valor RGB da cor
     */
    hsl(hue, sat, light)
    {
        sat /= 100, light /= 100, a := sat * Min(light, 1 - light)
        f(n)
        {
            k := Mod((n + hue / 30), 12)
            color := light - a * Max(Min(k - 3, 9 - k, 1), -1)
            return Round(255 * color)
        }
        return f(4) << 16 | f(8) << 8 | f(0)
    }
    /**
     * Insere um link no controle RichText
     * @param {string} text Texto a ser exibido.
     * @param {string|func} value Se for uma string, será usada como parâmetro da função Run().
     * Se for uma função, essa será chamada ao clicar no link.
     * @returns {object} objeto tipo Range
     */
    Link(text, value)
    {
        SendMessage(EM_SETEVENTMASK := 0x0445, , ENM_LINK := 0x04000000, this.Ctrl)
        rng := this.ITextDocument.Range(this.end, this.end)
        rng.Font.ForeColor := this.Blue, rng.Text := text

        if value is func
        {
            RichEdit.%value.name% := value
            link := value.name
        }
        else
            link := value

        rng.Url := '"' link '"'
        this.End := rng.End
        return rng
    }
    /**
     * Função Callback disparada quando se clica no link
     */
    ClickLink(ITextDocument, lParam)
    {
        if NumGet(lParam, A_PtrSize * 3, 'UInt') = WM_LBUTTONDOWN := 0x0201
        {
            start := NumGet(lParam, A_PtrSize * 5 + 4, 'UInt')
            end := NumGet(lParam, A_PtrSize * 5 + 8, 'UInt')
            value := ITextDocument.Range(start, end).Text

            RichEdit.HasOwnProp(value) ? (RichEdit.%value%)() : Run(value)
        }
    }
}