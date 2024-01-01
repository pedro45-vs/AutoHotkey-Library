if not A_AhkVersion ~= '2.1'
{
    Run('"' StrReplace(A_AhkPath, 'v2', 'v2.1') '" /restart "' A_ScriptFullPath '"')
    ExitApp()
}