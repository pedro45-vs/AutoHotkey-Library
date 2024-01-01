#Requires AutoHotkey v2.0

class SevenZip
{
    static exe := A_LineFile "\..\7zG.exe"
    
    static Add(filePath, zipFile)
    {
        Run Format('{} a "{}" "{}"', this.exe, filePath, zipFile)
    }
    static Extract(dest, zipFile, path := false)
    {
        
    }
}
SevenZip.Add("C:\Users\pedro\Downloads\hdzogcom video original.mp4", A_Desktop '\teste.zip')


