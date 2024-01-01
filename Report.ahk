/************************************************************************
 * @description Mostra uma GuiEdit com o conteúdo da variável.
 * Se for um Objeto, Array ou Map usará a lib json para mostrar.
 * Útil para mostrar relatórios sem precisar salvar em arquivo.
 * @file report.ahk
 * @author Pedro Henrique C Xavier
 * @date 2023/09/26
 * @version 2.0.10
 ***********************************************************************/

#Requires AutoHotkey v2.0

/**
 * Visualizador simples estilo bloco de notas para exibir relatórios
 */
class report extends Gui
{
    /**
     * @param {value} value Valor a ser exibido. Se for um objeto usará o estilo json
     * @param {string} header Cabeçalho opcional
     * @param {boolean} editable Opção para edição
     */
    __New(value, header := '', editable := false)
    {
        value := IsObject(value) ? report.stringify(value) : RTrim(value, '`n')

        super.__New('Resize MinSize240x40', 'Report: ' A_ScriptName, this)
        this.MarginX := this.MarginY := 2
        this.OnEvent('Size', 'Gui_Size')
        this.OnEvent('Close', (*) => this.Destroy())
        this.SetFont('S10', 'Consolas')
        this.SetFont('S10', 'JetBrains Mono')
        this.AddEdit('w0 h0')
        this.AddEdit('vEdit 0x4 -VScroll -HScroll -Wrap -WantReturn ' (editable ? '' : 'ReadOnly'), report.cab(header) . value)
        this.GetTextSize(this['Edit'], value)
        this.Show('AutoSize')
    }
    /**
     * Função executada sempre que se redimensionar a Gui
     * Calcula o novo tamanho e ativa ou desativa o scroll
     */
    Gui_Size(min_max, gui_width, gui_height)
    {
        this['Edit'].Opt((this.sizeWidth > gui_width) ? 'HScroll' : '-HScroll')
        this['Edit'].Opt((this.sizeHeight > gui_height) ? 'VScroll' : '-VScroll')
        this['Edit'].Move(, , gui_width - 4, gui_height - 6)
    }
    /**
     * Função para calcular o tamanho que um controle precisará para exibir o texto
     * com a fonte usada no GuiControl
     * @param textCtrl GuiControl
     * @param text string
     */
    GetTextSize(textCtrl, text)
    {
        static WM_GETFONT := 0x0031, DT_CALCRECT := 0x400
        hDC := DllCall('GetDC', 'Ptr', textCtrl.Hwnd, 'Ptr')
        hPrevObj := DllCall('SelectObject', 'Ptr', hDC, 'Ptr', SendMessage(WM_GETFONT, , , textCtrl), 'Ptr')
        height := DllCall('DrawText', 'Ptr', hDC, 'Str', text, 'Int', -1, 'Ptr', buf := Buffer(16), 'UInt', DT_CALCRECT)
        width := NumGet(buf, 8, 'Int') - NumGet(buf, 'Int')
        DllCall('SelectObject', 'Ptr', hDC, 'Ptr', hPrevObj, 'Ptr')
        DllCall('ReleaseDC', 'Ptr', textCtrl.Hwnd, 'Ptr', hDC)
        this.sizeWidth := width * 96 // A_ScreenDPI
        this.sizeHeight := height * 96 // A_ScreenDPI
    }
    /**
     * Cria um cabeçalho para a string informada
     * @param {string} str Cabeçalho
     * @returns {string}
     */
    static cab(str)
    {
        if len := StrLen(str)
        {
            rep := Format('{:0' len '}', 0)
            return str '`n' StrReplace(rep, 0, '═') '`n'
        }
    }
    /**
     * Cria um rodapé para a string informada
     * @param {string} str Rodapé
     * @returns {string}
     */
    static rod(str)
    {
        if len := StrLen(str)
        {
            rep := Format('{:0' len '}', 0)
            return '`n' StrReplace(rep, 0, '—') '`n' str
        }
    }
    static null := ComValue(1, 0), true := ComValue(0xB, 1), false := ComValue(0xB, 0)
    /**
     * https://github.com/thqby/ahk2_lib/blob/master/JSON.ahk
     * Converts a AutoHotkey Array/Map/Object to a Object Notation JSON string.
     * @param obj A AutoHotkey value, usually an object or array or map, to be converted.
     * @param expandlevel The level of JSON string need to expand, by default expand all.
     * @param space Adds indentation, white space, and line break characters to the return-value JSON text to make it easier to read.
     */
    static stringify(obj, expandlevel := unset, space := '  ')
    {
        expandlevel := IsSet(expandlevel) ? Abs(expandlevel) : 10000000
        return Trim(CO(obj, expandlevel))
        CO(O, J := 0, R := 0, Q := 0)
        {
            static M1 := '{', M2 := '}', S1 := '[', S2 := ']', N := '`n', C := ',', S := '- ', E := '', K := ':'
            if (OT := Type(O)) = 'Array'
            {
                D := !R ? S1 : ''
                for key, value in O
                {
                    F := (VT := Type(value)) = 'Array' ? 'S' : InStr('Map,Object', VT) ? 'M' : E
                    Z := VT = 'Array' && value.Length = 0 ? '[]' : ((VT = 'Map' && value.count = 0) || (VT = 'Object' && ObjOwnPropCount(value) = 0)) ? '{}' : ''
                    D .= (J > R ? '`n' CL(R + 2) : '') (F ? (%F%1 (Z ? '' : CO(value, J, R + 1, F)) %F%2) : ES(value)) (OT = 'Array' && O.Length = A_Index ? E : C)
                }
            }
            else
            {
                D := !R ? M1 : ''
                for key, value in (OT := Type(O)) = 'Map' ? (Y := 1, O) : (Y := 0, O.OwnProps())
                {
                    F := (VT := Type(value)) = 'Array' ? 'S' : InStr('Map,Object', VT) ? 'M' : E
                    Z := VT = 'Array' && value.Length = 0 ? '[]' : ((VT = 'Map' && value.count = 0) || (VT = 'Object' && ObjOwnPropCount(value) = 0)) ? '{}' : ''
                    D .= (J > R ? '`n' CL(R + 2) : '') (Q = 'S' && A_Index = 1 ? M1 : E) ES(key) K (F ? (%F%1 (Z ? '' : CO(value, J, R + 1, F)) %F%2) : ES(value)) (Q = 'S' && A_Index = (Y ? O.count : ObjOwnPropCount(O)) ? M2 : E) (J != 0 || R ? (A_Index = (Y ? O.count : ObjOwnPropCount(O)) ? E : C) : E)
                    if J = 0 && !R
                        D .= (A_Index < (Y ? O.count : ObjOwnPropCount(O)) ? C : E)
                }
            }
            if J > R
                D .= '`n' CL(R + 1)
            if R = 0
                D := RegExReplace(D, '^\R+') (OT = 'Array' ? S2 : M2)
            return D
        }
        ES(S)
        {
            switch Type(S)
            {
                case 'Float':
                    if (v := '', d := InStr(S, 'e'))
                        v := SubStr(S, d), S := SubStr(S, 1, d - 1)
                    if ((StrLen(S) > 17) && (d := RegExMatch(S, '(99999+|00000+)\d{0,3}$')))
                        S := Round(S, Max(1, d - InStr(S, '.') - 1))
                    return S v
                case 'Integer':
                    return S
                case 'String':
                    S := StrReplace(S, '\', '\\')
                    S := StrReplace(S, '`t', '\t')
                    S := StrReplace(S, '`r', '\r')
                    S := StrReplace(S, '`n', '\n')
                    S := StrReplace(S, '`b', '\b')
                    S := StrReplace(S, '`f', '\f')
                    S := StrReplace(S, '`v', '\v')
                    S := StrReplace(S, '"', '\"')
                    return '"' S '"'
                default:
                    return S == this.true ? 'true' : S == this.false ? 'false' : 'null'
            }
        }
        CL(i)
        {
            Loop (s := '', space ? i - 1 : 0)
                s .= space
            return s
        }
    }
}