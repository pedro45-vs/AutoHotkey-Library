/************************************************************************
 * @description Show a AutoHotkey value in StdOutput
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/05
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0
#Include %A_LineFile%\..\cJson.ahk

/**
 * Show a AutoHotkey value in StdOutput
 * @param {value} var string, number, array or map
 */
print(value)
{
    switch ValueType := Type(Value)
    {
        case 'Integer', 'Float', 'String': FileAppend(value '`n', '*')
        case 'Array', 'Map', 'Object': FileAppend(json.dump(value) '`n', '*')
        default: FileAppend(ValueType '`n', '*')
    }
}