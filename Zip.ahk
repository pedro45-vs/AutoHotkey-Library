/************************************************************************
 * @description Wrapper simples do 7Zip que utiliza a linha de comando
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/02
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0

class Zip
{
    /**
    *  Além do programa 7zG.exe requer o arquivo 7zip.dll na mesma pasta para funcionar
    */
    static exe := A_LineFile "\..\7zG.exe"

    /**
     * Adiciona uma pasta para um arquivo compactado
     * @param {string} zipFile Arquivo compactado
     * @param {string} filePath Diretório que será compactado
     * @param {boolean} delete Deletar após compactação?
     */
    static Add(zipFile, filePath, delete := false)
    {
        mode := delete ? '-sdel' : ''
        RunWait Format('{} a "{}" "{}" {}', this.exe, zipFile, filePath, mode)
    }
    /**
     * Extrai um arquivo zip para um diretório específico
     * @param {string} zipFile Arquivo compactado
     * @param {string} filePath Diretório para extração
     * @param {boolean} root Extrai todos os arquivos para a pasta raiz
     */
    static Extract(zipFile, filePath, root := true)
    {
        mode := root ? 'e' : 'x'
        RunWait Format('{} {} "{}" -o"{}"', this.exe, mode, zipFile, filePath)
    }
}
