/************************************************************************
 * @description Recupera o conteúdo de uma linha específica em arquivo de texto
 * @file GetFileLine.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/09/26
 * @version 2.0.10
 ***********************************************************************/
#Requires AutoHotKey v2.0

GetFileLine(file, line := 1)
{
    loop read file
        if A_Index = line
            return A_LoopReadLine
}
