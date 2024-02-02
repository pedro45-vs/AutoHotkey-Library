#Requires AutoHotkey v2.0

File.Prototype.Base := ExtendedFile()

class ExtendedFile
{
    /**
     * Permite ler linhas específicas de um arquivo
     * Inclusive selecionando intervalos arbitrários
     * Ex: 1-3;6;10-12
     */
    ReadLines(str_interval)
    {
        interval := this.__RangeInterval(str_interval)
        Loop
        {
            line := this.ReadLine()
            if InStr(interval, '|' A_Index '|')
                texto .= line '`n'
        }
        until this.AtEOF
        return texto
    }
    __RangeInterval(str_interval)
    {
        list := '|'
        for interval in StrSplit(str_interval, ';', A_Space)
        {
            sub_interval := StrSplit(interval, '-', A_Space)
            if sub_interval.Length = 2
            {
                for nItem in Range(sub_interval[1], sub_interval[2])
                    list .= nItem '|'
            }
            else
                list .= interval '|'
        }
        return list
        Range(Start, Stop, Step:=1) => (&n) => (n := Start, Start += Step, Step > 0 ? n <= Stop : n >= Stop)
    }
}