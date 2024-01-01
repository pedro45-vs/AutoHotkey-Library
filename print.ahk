/************************************************************************
 * @description Show a AutoHotkey value in StdOutput
 * @file print.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/08/31
 * @version 2.0.6
 ***********************************************************************/

#Requires AutoHotkey v2.0
#Include %A_LineFile%\..\json.ahk

/**
 * Show a AutoHotkey value in StdOutput
 * @param {value} var string, number, array or map
 */
print(value)
{
    switch ValueType := Type(Value)
    {
        case 'Integer', 'Float', 'String': FileAppend(value '`n', '*')
        case 'Array', 'Map', 'Object': FileAppend(json.stringify(value) '`n', '*')
        default: FileAppend(ValueType '`n', '*')
    }
}