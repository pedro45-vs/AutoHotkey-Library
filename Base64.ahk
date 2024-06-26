﻿/************************************************************************
 * @description Converte string ou arquivo para Base64 e vice-versa
 * @author Pedro Henrique C. Xavier
 * @date 2023-08-21
 * @version 2.0.5
 ***********************************************************************/
#Requires AutoHotKey v2.0

/**
* Converte string ou arquivo para Base64 e vice-versa
* Documentação CryptBinaryToStringA:
* https://learn.microsoft.com/en-us/windows/win32/api/wincrypt/nf-wincrypt-cryptstringtobinarya
* Documentação CryptStringToBinaryA:
* https://learn.microsoft.com/en-us/windows/win32/api/wincrypt/nf-wincrypt-cryptbinarytostringa
*/
Class Base64
{
    /**
     * Encode string or buffer object to base64
     * @prop {fun} 
     * @param data String ou Objeto do tipo Buffer (FileRead no modo 'RAW')
     * @returns {string} String Base64
     */
    static Encode(data)
    {
        if (data is Buffer)
            bin := data
        else if (data is String)
            bin := Buffer(StrPut(data, 'UTF-16')), StrPut(data, bin, 'UTF-16')
        else
            throw TypeError('Tipo não suportado', -1)

        if !(DllCall('crypt32\CryptBinaryToString', 'Ptr', bin.ptr, 'Int', bin.size - 1, 'Int', 0x1, 'ptr', 0, 'Int*', &size := 0))
            throw ValueError('CryptBinaryToString falhou', -1)

        buf := Buffer(size * 2, 0)

        if !(DllCall('crypt32\CryptBinaryToString', 'Ptr', bin.ptr, 'Int', bin.size - 1, 'Int', 0x1, 'ptr', buf.ptr, 'Int*', buf.ptr))
            throw ValueError('CryptBinaryToString falhou', -1)

        return StrGet(buf, , 'UTF-16')
    }
    
	static Encode2(Buf, Codec := 0x00000001) 
	{
		if Buf is String
			p := StrPtr(Buf), s := StrLen(Buf) * 2
		else p := Buf, s := Buf.Size
		if (DllCall("crypt32\CryptBinaryToString", "Ptr", p, "UInt", s, "UInt", Codec, "Ptr", 0, "Uint*", &nSize := 0) &&
			(VarSetStrCapacity(&VarOut, nSize << 1), DllCall("crypt32\CryptBinaryToString", "Ptr", p, "UInt", s, "UInt", Codec, "Str", VarOut, "Uint*", &nSize)))
			return (VarSetStrCapacity(&VarOut, -1), VarOut)
	}

    /**
     * Decode base64 to string or buffer object
     * @param str String em Base64
     * @param {Bolean} false retorna Buffer, true retorna String 
     * @returns {buffer|string} 
     */
    static Decode(str, mode := false)
    {
        bin := Buffer(StrPut(str, 'UTF-8')), StrPut(str, bin, 'UTF-8')

        if !(DllCall('crypt32\CryptStringToBinaryA', 'Char', bin.ptr, 'Int', 0, 'Int', 0x1, 'ptr', 0, 'uint*', &size := 0, 'ptr', 0, 'ptr', 0))
            throw ValueError('CryptStringToBinary falhou', -1)

        buf := Buffer(size, 0)

        if !(DllCall('crypt32\CryptStringToBinaryA', 'Char', bin.ptr, 'Int', 0, 'Int', 0x1, 'ptr', buf.ptr, 'uint*', &size, 'ptr', 0, 'ptr', 0))
            throw ValueError('CryptStringToBinary falhou', -1)

        return mode ? buf : StrGet(buf, , 'UTF-8')
    }
}

; https://www.autohotkey.com/boards/viewtopic.php?f=83&t=112821

Base64Decode(s) {
   ; Trim whitespace and remove mime type.
   s := Trim(s)
   s := RegExReplace(s, "(?i)^.*?;base64,")

   ; Retrieve the size of bytes from the length of the base64 string.
   size := StrLen(RTrim(s, "=")) * 3 // 4
   bin := Buffer(size)

   ; Place the decoded base64 string into a binary buffer.
   flags := 0x1 ; CRYPT_STRING_BASE64
   DllCall("crypt32\CryptStringToBinary", "str", s, "uint", 0, "uint", flags, "ptr", bin, "uint*", size, "ptr", 0, "ptr", 0)

   ; Must reinterpret the binary bytes from UTF-8.
   return StrGet(bin, size, "UTF-8")
}

Base64Encode(s) {
   ; Convert the input string into a byte string of UTF-8 characters.
   size := StrPut(s, "UTF-8")
   bin := Buffer(size)
   StrPut(s, bin, "UTF-8")
   size := size - 1 ; A binary does not have a null terminator

   ; Calculate the length of the base64 string.
   length := 4 * Ceil(size / 3) + 1   ; A string has a null terminator
   VarSetStrCapacity(&str, length)    ; Allocates a ANSI or Unicode string
   ; This appends 1 or 2 zero byte null terminators respectively.

   ; Passing a pre-allocated string buffer prevents an additional memory copy via StrGet.
   flags := 0x40000001 ; CRYPT_STRING_NOCRLF | CRYPT_STRING_BASE64
   DllCall("crypt32\CryptBinaryToString", "ptr", bin, "uint", size, "uint", flags, "str", str, "uint*", &length)

   ; Returns an AutoHotkey native string.
   return str
}



