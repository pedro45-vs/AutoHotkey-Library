

/**
 * Converte uma string RRGGBB para o valor correspondente
 * @param {string} str
 * @returns {integer}
 */
ColorHex(str) => Integer('0x' . SubStr(str, 5, 2) . SubStr(str, 3, 2) . SubStr(str, 1, 2))

/**
 * Retorna o valor RGB da cor
 * @param {r} valor Red (0-255)
 * @param {g} valor Green (0-255)
 * @param {b} valor Blue (0-255)
 * @returns {integer}
 */
RGB(r, g, b) => b << 16 | g << 8 | r


hsl(h, s, l) 
{
    l /= 100
    a := s * Min(l, 1 - l) / 100
    f(n)
    {
        k := Mod((n + h / 30), 12)
        color := l - a * Max(Min(k - 3, 9 - k, 1), -1)
        return Round(255 * color)
    }
    return f(4) << 16 | f(8) << 8 | f(0)
}

hslToHex(h, s, l) 
{
    l /= 100
    a := s * Min(l, 1 - l) / 100
    f(n)
    {
        k := Mod((n + h / 30), 12)
        color := l - a * Max(Min(k - 3, 9 - k, 1), -1)
        return Round(255 * color)
    }
    return Format('#{:02X}{:02X}{:02X}', f(0), f(8), f(4))
}