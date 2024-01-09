#Requires AutoHotkey v2.0

RichEdit.Prototype.Base := RichTable()

class RichTable
{
    Header(str, level := 1)
    {
        rng := this.ITextDocument.Range(this.end, this.end)
        rng.Text := str '`n'
        rng.Size := -2 * level + 18
        rng.Bold := rng.Shadow := this.tomTrue
        rng.SpaceAfter := rng.SpaceBefore := -2 * level + 18
        this.End := rng.End
        return rng
    }

    Table(table_array, Width, Align := [])
    {
        static color := this.Gray[6]
        SetAlign := Map(), SetAlign.CaseSense := false
        SetAlign.Set('l', this.tomAlignLeft, 'c', this.tomAlignCenter, 'r', this.tomAlignRight)

        fonH := this.ItextFont, fonH.Bold := this.tomTrue
        rowp := {
            CellBorderColor: this.Gray[6],
            CellAlignment: this.tomAlignCenter,
            CellWidth: Width,
            Height: 24
        }

        Align.Length := table_array[1].Length, Alignment := []
        for value in Align
            Alignment.Push( SetAlign[value ?? 'l'] )
        rowp.Alignment := Alignment

        head := rowp.Clone(), head.CellColorBack := this.Gray[8], head.Font := fonH
        rowc := rowp.Clone(), rowc.CellColorBack := this.Gray[10]

        for row_array in table_array
        {
            if A_Index = 1
                this.InsertRow(row_array, head)

            else if A_Index & 1
                this.InsertRow(row_array, rowc)

            else
                this.InsertRow(row_array, rowp)
        }
    }

    Table2(table_array, width, align := [])
    {
        nCols := table_array[1].length, twips := 15, color := this.Gray[6]
        SetAlign := Map(), SetAlign.CaseSense := false
        SetAlign.Set('l', this.tomAlignLeft, 'c', this.tomAlignCenter, 'r', this.tomAlignRight)
        align.Length := nCols, align.Default := 'l'

        this.Ctrl.Move(, , 60)
        rng := this.ITextDocument.Range(this.End, this.End)

        for index, row_array in table_array
        {
            MsgBox(index, nCols)
            rng.InsertTable(nCols, 1, 0)
            rng.Move(tomTable := 15, -1)
            row := rng.Row
            row.Height := twips * 24
            Loop nCols
            {
                row.CellIndex := A_Index - 1
                row.SetCellBorderColors(color, color, color, color)
                row.CellWidth := twips * width[A_Index]
                row.CellAlignment := SetAlign['c']

                if index = 1
                    row.CellColorBack := this.Gray[8]
                else if index & 1
                    row.CellColorBack := this.Gray[10]
                else
                    row.CellColorBack := this.White
            }
            ; Aplicação das modificações da estrutura da tabela
            row.Apply(1, 0)

            rng.Move(tomRow := 10, -1)
            for value in row_array
            {
                rng.Text := value ?? ''
                rng.Alignment := SetAlign[ align[A_Index] ]
                (index = 1) && rng.Bold := this.tomTrue
                rng.Move(tomCell := 12, 1)
            }
            this.End += 300
        }
    }
}