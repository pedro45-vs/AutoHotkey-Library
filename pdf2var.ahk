/************************************************************************
 * @description Extrai o texto de um arquivo PDF
 * @author Pedro Henrique C. Xavier
 * @date 09.11.2023
 * @version 2.0.10
 ***********************************************************************/
#Requires AutoHotkey v2

/**
 * Extrai o texto de um arquivo PDF
 * @param {string} PdfPath Caminho completo do arquivo PDF
 * @param {string} options Opções compatíveis com a biblioteca pdftotext.exe
 * @returns {string}
 */
pdf2var(PdfPath, options := '')
{
    try strpdf := StdoutToVar(A_LineFile '\..\pdftotext.exe -q ' options '`s' '"' PdfPath '"' ' -')
    catch
    {
        throw ValueError('Ocorreu um erro ao chamar a função StdoutToVar', -1)
    }
    switch strpdf.ExitCode
    {
        case 0:  return strpdf.Output
        case 1:  throw ValueError('Ocorreu um erro ao abrir o arquivo`n' PdfPath, -1)
        case 2:  throw ValueError('Ocorreu um erro ao criar arquivo de saída', -1)
        case 3:  throw ValueError('Ocorreu um erro de permissão de leitura', -1)
        case 99: throw ValueError('Ocorreu um erro não conhecido', -1)
    }

    ; Source: https://www.autohotkey.com/boards/viewtopic.php?t=109148
    ; Fiz algumas alterações para meu uso pessoal e removi alguns trechos que julgo desnecessários
    ; CP TESTADOS EM PORTUGUÊS: CP65000, CP850, CP857, CP858
    ; ----------------------------------------------------------------------------------------------------------------------
    StdoutToVar(sCmd, sEnc := '')
    {
        ; Create 2 buffer-like objects to wrap the handles to take advantage of the __Delete meta-function.
        oHndStdoutRd := { Ptr: 0, __Delete: delete(this) => DllCall('CloseHandle', 'Ptr', this) }
        oHndStdoutWr := { Base: oHndStdoutRd }
        
        if !DllCall('CreatePipe', 'PtrP', oHndStdoutRd, 'PtrP', oHndStdoutWr, 'Ptr', 0, 'UInt', 0)
            throw OSError(, , 'Error creating pipe.')
        if !DllCall('SetHandleInformation', 'Ptr', oHndStdoutWr, 'UInt', 1, 'UInt', 1)
            throw OSError(, , 'Error setting handle information.')
        
        PI := Buffer(A_PtrSize == 4 ? 16 : 24, 0)
        SI := Buffer(A_PtrSize == 4 ? 68 : 104, 0)
        NumPut('UInt', SI.Size, SI, 0)
        NumPut('UInt', 0x100, SI, A_PtrSize == 4 ? 44 : 60)
        NumPut('Ptr', oHndStdoutWr.Ptr, SI, A_PtrSize == 4 ? 60 : 88)
        NumPut('Ptr', oHndStdoutWr.Ptr, SI, A_PtrSize == 4 ? 64 : 96)

        if !DllCall('CreateProcess', 'Ptr', 0, 'Str', sCmd, 'Ptr', 0, 'Ptr', 0, 'Int', 1,
            'UInt', 0x08000000, 'Ptr', 0, 'Ptr', 0, 'Ptr', SI, 'Ptr', PI)
            throw OSError(, , 'Error creating process.')
        
        ; The write pipe must be closed before reading the stdout so we release the object.
        ; The reading pipe will be released automatically on function return.
        oHndStdoutWr := ''
        
        ; Before reading, we check if the pipe has been written to, so we avoid freezings.
        nAvail := nLen := 0
        
        while DllCall('PeekNamedPipe', 'Ptr', oHndStdoutRd, 'Ptr', 0, 'UInt', 0, 'Ptr', 0, 'UIntP', &nAvail, 'Ptr', 0) != 0
        {
            ; If the pipe buffer is empty, sleep and continue checking.
            if !nAvail && Sleep(10)
                continue
            cBuf := Buffer(nAvail + 1)
            DllCall('ReadFile', 'Ptr', oHndStdoutRd, 'Ptr', cBuf, 'UInt', nAvail, 'PtrP', &nLen, 'Ptr', 0)
            sOutput .= StrGet(cBuf, nLen, sEnc)
        }
        
        ; Get the exit code, close all process handles and return the output object.
        DllCall('GetExitCodeProcess', 'Ptr', NumGet(PI, 0, 'Ptr'), 'UIntP', &nExitCode := 0)
        DllCall('CloseHandle', 'Ptr', NumGet(PI, 0, 'Ptr'))
        DllCall('CloseHandle', 'Ptr', NumGet(PI, A_PtrSize, 'Ptr'))
        return { Output: sOutput, ExitCode: nExitCode }
    }
}


/*
pdftotext(1)                General Commands Manual               pdftotext(1)
------------------------------------------------------------------------------
NAME
pdftotext  -  Portable Document Format (PDF) to text converter (version
4.04)

SYNOPSIS
pdftotext [options] [PDF-file [text-file]]

DESCRIPTION
Pdftotext converts Portable Document Format (PDF) files to plain text.

Pdftotext reads the PDF file, PDF-file, and writes a text  file,  text-
file.   If  text-file  is not specified, pdftotext converts file.pdf to
file.txt.  If text-file is '-', the text is sent to stdout.

CONFIGURATION FILE
Pdftotext reads a configuration file at startup.   It  first  tries  to
find the user's private config file, ~/.xpdfrc.  If that doesn't exist,
it looks for a system-wide config file, typically /etc/xpdfrc (but this
location  can  be  changed when pdftotext is built).  See the xpdfrc(5)
man page for details.

OPTIONS
Many of the following options can be set with configuration  file  com-
mands.  These are listed in square brackets with the description of the
corresponding command line option.

-f number
Specifies the first page to convert.

-l number
Specifies the last page to convert.

-layout
Maintain (as best as possible) the original physical  layout  of
the  text.   The  default is to 'undo' physical layout (columns,
hyphenation, etc.) and output the text in reading order.  If the
-fixed  option is given, character spacing within each line will
be determined by the specified character pitch.

-simple
Similar to -layout, but optimized for simple  one-column  pages.
This  mode  will do a better job of maintaining horizontal spac-
ing, but it will only work properly  with  a  single  column  of
text.

-simple2
Similar to -simple, but handles slightly rotated text (e.g., OCR
output) better.  Only works for pages with a  single  column  of
text.

-table Table mode is similar to physical layout mode, but optimized for
tabular data, with the goal of keeping rows and columns  aligned
(at  the  expense of inserting extra whitespace).  If the -fixed
option is given, character spacing  within  each  line  will  be
determined by the specified character pitch.

-lineprinter
Line  printer  mode  uses  a  strict  fixed-character-pitch  and
-height layout.  That is, the page is broken into  a  grid,  and
characters  are  placed  into that grid.  If the grid spacing is
too small for the actual characters, the result is extra  white-
space.   If the grid spacing is too large, the result is missing
whitespace.  The grid spacing can be specified using the  -fixed
and  -linespacing  options.  If one or both are not given on the
command line, pdftotext  will  attempt  to  compute  appropriate
value(s).

-raw   Keep the text in content stream order.  Depending on how the PDF
file was generated, this may or may not be useful.

-fixed number
Specify the character pitch (character width),  in  points,  for
physical  layout,  table, or line printer mode.  This is ignored
in all other modes.

-linespacing number
Specify the line spacing, in  points,  for  line  printer  mode.
This is ignored in all other modes.

-clip  Text which is hidden because of clipping is removed before doing
layout, and then added back in.  This can be helpful for  tables
where clipped (invisible) text would overlap the next column.

-nodiag
Diagonal text, i.e., text that is not close to one of the 0, 90,
180, or 270 degree axes, is discarded.  This is useful  to  skip
watermarks drawn on top of body text, etc.

-enc encoding-name
Sets  the  encoding  to  use for text output.  The encoding-name
must be defined with the  unicodeMap  command  (see  xpdfrc(5)).
The  encoding name is case-sensitive.  This defaults to "Latin1"
(which is a built-in encoding).  [config file: textEncoding]

-eol unix | dos | mac
Sets the end-of-line convention to use for text output.  [config
file: textEOL]

-nopgbrk
Don't  insert  a page breaks (form feed character) at the end of
each page.  [config file: textPageBreaks]

-bom   Insert a Unicode byte order marker (BOM) at  the  start  of  the
text output.

-marginl number
Specifies  the  left margin, in points.  Text in the left margin
(i.e., within that many points of the left edge of the page)  is
discarded.  The default value is zero.

-marginr number
Specifies the right margin, in points.  Text in the right margin
(i.e., within that many points of the right edge of the page) is
discarded.  The default value is zero.

-margint number
Specifies  the  top  margin,  in points.  Text in the top margin
(i.e., within that many points of the top edge of the  page)  is
discarded.  The default value is zero.

-marginb number
Specifies the bottom margin, in points.  Text in the bottom mar-
gin (i.e., within that many points of the  bottom  edge  of  the
page) is discarded.  The default value is zero.

-opw password
Specify  the  owner  password  for the PDF file.  Providing this
will bypass all security restrictions.

-upw password
Specify the user password for the PDF file.

-verbose
Print a status message (to stdout) before processing each  page.
[config file: printStatusInfo]

-q     Don't print any messages or errors.  [config file: errQuiet]

-cfg config-file
Read config-file in place of ~/.xpdfrc or the system-wide config
file.

-listencodings
List all available text output encodings, then exit.

-v     Print copyright and version information, then exit.

-h     Print usage information,  then  exit.   (-help  and  --help  are
equivalent.)

BUGS
Some  PDF  files contain fonts whose encodings have been mangled beyond
recognition.  There is no way (short of OCR) to extract text from these
files.

EXIT CODES
The Xpdf tools use the following exit codes:

0      No error.

1      Error opening a PDF file.

2      Error opening an output file.

3      Error related to PDF permissions.

99     Other error.

AUTHOR
The  pdftotext software and documentation are copyright 1996-2022 Glyph
& Cog, LLC.

SEE ALSO
xpdf(1),  pdftops(1),  pdftohtml(1),  pdfinfo(1),  pdffonts(1),  pdfde-
tach(1), pdftoppm(1), pdftopng(1), pdfimages(1), xpdfrc(5)
http://www.xpdfreader.com/



18 Apr 2022                     pdftotext(1)
