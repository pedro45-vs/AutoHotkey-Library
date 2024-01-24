#Requires AutoHotkey v2.0

ExpandRelativePath(relative_path)
{
    VarSetStrCapacity(&out, 260)
    if DllCall('shlwapi\PathCanonicalize', 'str', out, 'str', relative_path)
        return out
}