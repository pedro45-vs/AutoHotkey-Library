/************************************************************************
 * @description class para trabalhar com arquivos csv
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/25
 * @version 2.0.11
 ***********************************************************************/

#Requires AutoHotkey v2.0

class ClassCSV
{
    __New(delimiter := ';', Encoding := 'UTF-8')
    {
        this.Delimiter := delimiter
        this.Encoding := Encoding
    }

    CountRows
    {
        get
        {
            Loop parse this.text, '`n', '`r'
                cr := A_Index
            return cr
        }
    }

    CountColumns
    {
        get
        {
            Loop parse this.text, '`n', '`r'
                if A_Index = 1
                    return StrSplit(A_LoopField, this.Delimiter).Length
        }
    }

    Load(file)
    {
        FileEncoding(this.Encoding)
        this.text := Trim(FileRead(file), '`n')
        this.FilePath := file
    }

    ReadColumns(columns*)
    {
        FileEncoding(this.Encoding)
        Loop Read this.FilePath
        {
            str := StrSplit(A_LoopReadLine, this.Delimiter)
            For column in columns
                TotalColumms .= str[column] . this.Delimiter

            TotalColumms .= '`n'
        }
        return Trim(TotalColumms, '`n')
    }

    ReadRows(rows*)
    {
        FileEncoding(this.Encoding)
        For row in rows
        {
            Loop Read this.FilePath
                if A_Index = row
                    TotalRows .= A_LoopReadLine '`n'
        }
        return Trim(TotalRows, '`n')
    }

    SelectColumns(columns*)
    {
        this.text := this.ReadColumns(columns*)
    }

    SelectRows(rows*)
    {
        this.text := this.ReadRows(rows*)
    }

    ChangeDelimiter(delimiter)
    {
        this.text := StrReplace(this.text, this.delimiter, delimiter)
    }

    DeleteColumns(columns*)
    {
        Loop parse this.text, '`n', '`r'
        {
            str := StrSplit(A_LoopField, this.Delimiter)
            For column in columns
                str.Delete(column)

            For column in str
            {
                if str.Has(A_Index)
                    TotalColumms .= str[A_Index] . this.Delimiter
            }
            TotalColumms .= '`n'
        }
        this.text := Trim(TotalColumms, '`n')
    }

    SortCol(col)
    {
        SortProd(a1, a2, *)
        {
            a1 := StrSplit(a1, this.Delimiter)[col]
            a2 := StrSplit(a2, this.Delimiter)[col]
            return StrCompare(a1, a2)
        }
        this.text := Sort(this.text, , SortProd)
    }

    Save(FilePath := this.FilePath, Encoding := this.Encoding)
    {
        file := FileOpen(FilePath, 'w `n', Encoding)
        file.Write(this.text)
        file.Close()
    }
}
