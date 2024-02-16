#Requires AutoHotkey v2.0
#Include %A_LineFile% \..\cJson.ahk
#Include %A_LineFile% \..\RichEdit.ahk
#Include %A_LineFile% \..\RichEditMenu.ahk
#SingleInstance

Debug(var)
{
    GuiD := Gui('Resize', 'Debug: ' A_ScriptName)
    GuiD.OnEvent('Close', (*)=> ExitApp())
    GuiD.OnEvent('Size', (o, m, w, h) => Rich.Ctrl.Move(,, w-20, h-60))
    GuiD.SetFont('s11', 'Calibri')
    GuiD.MarginX := MarginY := 10
    GuiD.AddButton('w100', 'Play').OnEvent('Click', TogglePause)
    GuiD.AddButton('w110 x+10', 'Reload').OnEvent('Click', (*) => Reload())
    GuiD.AddButton('w110 x+10', 'ListVars').OnEvent('Click', (*) => ListVars())
    GuiD.AddButton('w110 x+10', 'ListLines').OnEvent('Click', (*) => ListLines())
    GuiD.AddButton('w110 x+10', 'Windows Spy').OnEvent('Click', (*) => Run(A_AhkPath '\..\..\WindowSpy.ahk'))
    GuiD.AddButton('w110 x+10', 'Help').OnEvent('Click', (*) => Run(A_AhkPath '\..\AutoHotkey.chm'))

    TogglePause(CtrlObj, info)
    {
        CtrlObj.Text := CtrlObj.Text = 'Play' ? 'Pause' : 'Play'
        Pause(-1)
    }

    Rich := RichEdit(GuiD, 'xm vScroll hScroll w300 h200 Border')
    Rich.SetMargins(10, 10)
    Rich.SetBkgnColor(0xF0F0F0)
    Rich.ShowMenu()
    fon := Rich.ITextFont
    fon.Bold := Rich.tomTrue

    cab := Rich.ItextFont
    cab.Name := 'Verdana', cab.Bold := Rich.tomTrue, cab.Size := 24, cab.ForeColor := Rich.Blue[3]
    Rich.Text(var_type := Type(var) '`n').Font := cab
    lin := Rich.Line(17, 1)
    lin.Font := cab, lin.SpaceBefore := 0

    if var is ComObject
    {
        rowformat := {CellWidth: [50, 160, 400, 80, 80, 80]}
        cabFormat := rowformat.Clone()
        cabFormat.DefineProp('CellColorBack', {value: [0xf0f0f0, 0xf0f0f0, 0xf0f0f0, 0xf0f0f0, 0xf0f0f0, 0xf0f0f0]})
        cabFormat.DefineProp('Font', {value: [fon, fon, fon, fon, fon, fon]})
        Rich.InsertRow(['ID', 'Name', 'Description', 'Kind', 'Arg count', 'Opt args'], cabFormat)
        for id, info in EnumComMembers(var)
            Rich.InsertRow([id, info.name, info.desc, info.kind, info.argNum, info.optArgs], rowformat)

        GuiD.Show('w920 h600')
    }
    else if var is Primitive
    {
        Rich.Text(var).Name := 'Consolas'
        GuiD.Show('w720 h400')
    }
    else
    {
        EnumProps(var)
        GuiD.Show('w720 h600')
    }
    Rich.ReadOnly()
    Pause()


    EnumProps(obj)
    {
        Prop := [], Methods := []
        ; Enumeração dos métodos e propriedades Base
        for iProp in obj.base.OwnProps()
        {
            iDesc := obj.base.GetOwnPropDesc(iProp)

            if iDesc.HasProp('Call')
            {
                Methods.Push([iProp '()' ,obj.%iProp%.IsBuiltIn, obj.%iProp%.isVariadic, obj.%iProp%.MinParams, obj.%iProp%.MaxParams])
            }
            else if iDesc.HasProp('Get')
            {
                try Prop.Push([iProp, (T := Type(V := var.%iProp%)), (T ~= 'Integer|Float|String') ? var.%iProp% : ''])
            }
            if iProp = '__Item' or iProp = 'OwnProps'
                itens := Json.Dump(var)
        }
        for iProp in var.OwnProps()
        {
            value := (T := Type(V := var.%iProp%)) ~= 'Float|Integer|String' ? V : ''
            Prop.Push([iProp, T, value])
        }

        fon := Rich.ITextFont
        fon.Size := 16, fon.Underline := fon.Bold := Rich.tomTrue
        color := Rich.Color['9BCAEF']

        if IsSet(itens)
        {
            Rich.ItextDocument.DefaultTabStop := 15
            Rich.Text('Itens:').Font := fon
            Rich.Space(4)
            Rich.Text(itens).Name := 'Consolas'
            Rich.Space(14)
        }
        if Prop.Length
        {
            cell := {CellWidth: [200, 120, 120]}
            cab := cell.Clone()
            cab.DefineProp('CellColorBack', {value: [color, color, color]})
            Rich.Text('Properties:').Font := fon
            Rich.Space(14)
            Rich.InsertRow(['Description', 'Type', 'Value'], cab)
            for arr in Prop
                Rich.InsertRow(arr, cell)
            Rich.Space(14)
        }
        if Methods.Length
        {
            cell := {CellWidth: [200, 85, 85, 85, 85]}
            cab := cell.Clone()
            cab.DefineProp('CellColorBack', {value: [color, color, color, color, color]})
            Rich.Text('Methods:').Font := fon
            Rich.Space(14)
            Rich.InsertRow(['Description', 'IsBuiltIn', 'IsVariadic', 'MinParams', 'MaxParams'], cab)
            for arr in Methods
                Rich.InsertRow(arr, cell)
        }
    }
    /**
     * https://www.autohotkey.com/boards/viewtopic.php?p=540210&sid=b5af05d451d318f85282e4fa19308e51#p540210
     * @param IDispatch
     * @returns {map}
     */
    EnumComMembers(IDispatch)
    {
        static VT_DISPATCH := 9, VT_UNKNOWN := 13, MEMBERID_NIL := -1
        if ComObjType(IDispatch) != VT_DISPATCH
        {
            return MsgBox('IDispatch type only supported')
        }
        pIDispatch := ComObjValue(IDispatch)
        ComCall(GetTypeInfoCount := 3, pIDispatch, 'UIntP', &hasInfo := 0)
        if !hasInfo
        {
            return MsgBox('ITypeInfo Interface not supported')
        }
        ComCall(GetTypeInfo := 4, pIDispatch, 'UInt', 0, 'UInt', 0, 'PtrP', ITypeInfo := ComValue(VT_UNKNOWN, 0))
        name := 'IDispatch'
        Loop
        {
            ComCall(GetTypeAttr := 3, ITypeInfo, 'PtrP', &pTypeAttr := 0)
            cFuncs := NumGet(pTypeAttr, 40 + A_PtrSize, 'Short')
            cImplTypes := NumGet(pTypeAttr, 44 + A_PtrSize, 'Short')
            ComCall(ReleaseTypeAttr := 19, ITypeInfo, 'Ptr', pTypeAttr)

            if cImplTypes
            {
                ComCall(GetRefTypeOfImplType := 8, ITypeInfo, 'Int', 0, 'PtrP', &pRefType := 0)
                ComCall(GetRefTypeInfo := 14, ITypeInfo, 'Ptr', pRefType, 'PtrP', ITypeInfo2 := ComValue(VT_UNKNOWN, 0))
                ComCall(GetDocumentation := 12, ITypeInfo2, 'Int', MEMBERID_NIL, 'PtrP', &pName := 0, 'Ptr', 0, 'Ptr', 0, 'Ptr', 0)
                name := StrGet(pName)
                DllCall('OleAut32\SysFreeString', 'Ptr', pName)
                if name != 'IDispatch'
                {
                    ITypeInfo := ITypeInfo2
                }
            }
        }
        until name == 'IDispatch'
        info := Map()
        Loop cFuncs
        {
            ComCall(GetFuncDesc := 5, ITypeInfo, 'UInt', A_Index - 1, 'PtrP', &pFuncDesc := 0)
            id := NumGet(pFuncDesc, 0, 'Short')
            invkind := NumGet(pFuncDesc, 4 + A_PtrSize * 3, 'Short')
            argNum := NumGet(pFuncDesc, 12 + A_PtrSize * 3, 'Short')
            optArgs := NumGet(pFuncDesc, 14 + A_PtrSize * 3, 'Short')
            ComCall(ReleaseFuncDesc := 20, ITypeInfo, 'Ptr', pFuncDesc)
            try ComCall(GetDocumentation := 12, ITypeInfo, 'Int', id, 'PtrP', &pName := 0, 'PtrP', &pDesc := 0, 'Ptr', 0, 'Ptr', 0)
            catch
            {
                continue
            }
            name := StrGet(pName), desc := ''
            (pDesc && desc := StrGet(pDesc))
            DllCall('OleAut32\SysFreeString', 'Ptr', pName), DllCall('OleAut32\SysFreeString', 'Ptr', pDesc)
            kind := ['method', 'get', , 'put', , , , 'putref'][invkind]
            if info.Has(id) && info[id].kind != 'method' && !(info[id].kind ~= '(^|,\s)' . kind . '(,|$)')
            {
                info[id].kind .= ', ' . kind, info[id].argNum := argNum, info[id].optArgs := optArgs
            }
            else
            {
                info[id] := { name: name, desc: desc, kind: kind, argNum: argNum, optArgs: optArgs }
            }
        }
        return info
    }
}