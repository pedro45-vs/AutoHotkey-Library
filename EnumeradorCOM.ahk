/************************************************************************
 * @description Visualiza as propriedades e métodos de uma interface COM
 * @author Pedro Henrique C. Xavier
 * @date 2023-11-16
 * @version 2.1-alpha.7
 ***********************************************************************/

#Requires AutoHotKey v2

/**
 * Mostra um ListView com as propriedades e métodos da interface COM
 * @param IDispatch Interface COM
 */
EnumeradorCOM(IDispatch)
{
    GuiC := Gui('Resize', Type(IDispatch))
    GuiC.MarginX := GuiC.MarginY := 0
    GuiC.OnEvent('Size', Gui_Size)
    GuiC.SetFont('s9', 'Segoe UI')
    GuiC.AddListView('Grid -Multi r16 vList', ['ID', 'Name', 'Description', 'Kind', 'Arg count', 'Opt args'])
    GuiC['List'].OnEvent('ContextMenu', Ctrl_ContextMenu)

    for id, info in EnumComMembers(IDispatch)
        GuiC['List'].Add('Vis', id, info.name, info.desc, info.kind, info.argNum, info.optArgs)

    loop 6
        GuiC['List'].ModifyCol(A_Index, 'AutoHdr')

    GuiC.Show('w940')

    Gui_Size(GuiObj, MinMax, Width, Height) => GuiC['List'].Move(, , Width, Height)
    
    Ctrl_ContextMenu(GuiCtrlObj, Item, IsRightClick, X, Y)
    {
        ToolTip(A_Clipboard := GuiCtrlObj.GetText(item, 2))
        SetTimer () => ToolTip(), -1500
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
