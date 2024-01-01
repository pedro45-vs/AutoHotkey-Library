#Requires AutoHotkey v2.0

ExpandRelativePath(relative_path)
{
    buf := Buffer(260, 0)
    DllCall('shlwapi\PathCombine', 'ptr', buf, 'wStr', A_WorkingDir, 'wStr', relative_path)
    return StrGet(buf)
}