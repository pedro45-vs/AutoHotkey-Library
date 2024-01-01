/************************************************************************
 * @description Ferramentas para alterar o conteúdo do ClipBoard
 * @file ClipTools.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2023/08/21
 * @version 2.0.5
 ***********************************************************************/
#Requires AutoHotkey v2.0
#Include %A_LineFile% \..\calcular.ahk
#Include %A_LineFile% \..\NumLib.ahk

/**
 * Ferramentas para alterar o conteúdo do ClipBoard
 * @param {Number} tool 1 - Aspas simples; 2 - Aspas duplas; 3 - Parênteses; 4 - UpperCase;
 * 5 - LowerCase; 6 - TitleCase; 7 - Capital Letter; 8 - Remove parênteses e aspas do texto;
 * 9 -  Deixa apenas os números e caracteres de expressão
 */
ClipTools(tool)
{
    A_Clipboard := ''
    Send('^c')
    ClipWait(1)
    switch tool
    {
        case 1: A_Clipboard := "'" Trim(A_Clipboard) "'"  	; Insere Aspas Simples
        case 2: A_Clipboard := '"' Trim(A_Clipboard) '"'  	; Insere Aspas Duplas
        case 3: A_Clipboard := '(' Trim(A_Clipboard) ')'  	; Insere Parênteses
        case 4: A_Clipboard := StrUpper(A_Clipboard)
        case 5: A_Clipboard := StrLower(A_Clipboard)
        case 6: A_Clipboard := StrTitle(A_Clipboard)
        case 7: A_Clipboard := StrUpper(SubStr(A_Clipboard, 1, 1)) StrLower(SubStr(A_Clipboard, 2))
        case 8:
            if RegExMatch(A_Clipboard, 's)^[\x22\x27()](.*)[\x22\x27()]$', &Rem)
                A_Clipboard := Rem[1]                   ; Remove parênteses e aspas do texto
        case 9:                                         ; Deixa apenas os números e caracteres de expressão
            A_Clipboard := milhar(Calcular(RegExReplace(A_Clipboard, '[^\d\,\+\-\*\/\(\)\s]')))
    }
    ClipWait(1)
    Send('^v')
}
