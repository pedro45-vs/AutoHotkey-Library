/************************************************************************
 * @description 
 * @author Pedro Henrique C. Xavier
 * @date 2024-03-28
 * @version 2.1-alpha.9
 ***********************************************************************/

#Requires AutoHotkey v2.0
#Include %A_LineFile% \..\OCR.ahk

pdfOcrExtract(PDF_file)
{
   
    static PNG_root := A_Temp '\pdfOcrExtract'
    DirExist(PNG_root) || DirCreate(PNG_root)
    
    Loop files PNG_root '\*.png'
        try FileDelete(A_LoopFileFullPath)
        catch
            MsgBox('Não foi possível excluir o arquivo ' A_LoopFileName, 'pdfOcrExtract', 'IconX')
    
    ; Converte o arquivo pdf em png e salva em uma pasta temporária
    command := Format('"{}\..\pdftopng.exe" -q "{}" "{}"', A_LineFile, PDF_file, PNG_root '\PDF-OCR')
    RunWait(command,, 'Hide')
    
    Loop files PNG_root '\*.png'
        texto .= OCR.FromFile(A_LoopFileFullPath, 'pt-BR').text
    
    return texto
}