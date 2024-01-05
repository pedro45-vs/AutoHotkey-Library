#Requires AutoHotkey v2.0

Array.Prototype.Base := ExtendedArray()

class ExtendedArray
{
    Find(value)
    {
        for item in this
            if item == value
                return A_Index
        return false
    }
    toString()
    {
        for item in this
            str .= item ', '
        return SubStr(str, 1, -2)
    }
}