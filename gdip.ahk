; Classe GDIp extraido do script Snipper disponível em:
; https://www.autohotkey.com/boards/viewtopic.php?f=83&t=115622
; Update: 24.04.2023

; Com essa classe é possível salvar prints da tela e salvar em arquivo
; de forma automatizada. Salvei essa biblioteca separada para essa função

#Requires AutoHotkey v2.0

gdip.SavePrintScreen()

SavePrintScreen()
{
    GDIp.Startup()
    pBitmap := GDIp.BitmapFromScreen({ X: 0, Y: 0, W: A_ScreenWidth, H: A_ScreenHeight })
    GDIp.SaveBitmapToFile(pBitmap,  'Logs\' A_ScriptName '.png')
    GDIp.DisposeImage(pBitmap)
    GDIp.Shutdown()
}

#DllLoad 'GdiPlus'
Class GDIp
{
    Static Startup()
    {
        If (this.HasProp("Token"))
            Return
        input := Buffer((A_PtrSize = 8) ? 24 : 16, 0)
        NumPut("UInt", 1, input)
        DllCall("gdiplus\GdiplusStartup", "UPtr*", &pToken := 0, "UPtr", input.ptr, "UPtr", 0)
        this.Token := pToken
    }

    Static Shutdown()
    {
        If (this.HasProp("Token"))
            DllCall("Gdiplus\GdiplusShutdown", "UPtr", this.DeleteProp("Token"))
    }

    Static BitmapFromScreen(Area)
    {
        chdc := this.CreateCompatibleDC()
        hbm := this.CreateDIBSection(Area.W, Area.H, chdc)
        obm := this.SelectObject(chdc, hbm)
        hhdc := this.GetDC()
        this.BitBlt(chdc, 0, 0, Area.W, Area.H, hhdc, Area.X, Area.Y)
        this.ReleaseDC(hhdc)
        pBitmap := this.CreateBitmapFromHBITMAP(hbm)
        this.SelectObject(chdc, obm), this.DeleteObject(hbm), this.DeleteDC(hhdc), this.DeleteDC(chdc)
        Return pBitmap
    }

    Static SetBitmapToClipboard(pBitmap)
    {
        off1 := A_PtrSize = 8 ? 52 : 44
        off2 := A_PtrSize = 8 ? 32 : 24

        pid := DllCall("GetCurrentProcessId", "uint")
        hwnd := WinExist("ahk_pid " . pid)
        r1 := DllCall("OpenClipboard", "UPtr", hwnd)
        If !r1
            Return -1

        hBitmap := this.CreateHBITMAPFromBitmap(pBitmap, 0)
        If !hBitmap
        {
            DllCall("CloseClipboard")
            Return -3
        }

        r2 := DllCall("EmptyClipboard")
        If !r2
        {
            this.DeleteObject(hBitmap)
            DllCall("CloseClipboard")
            Return -2
        }

        oi := Buffer((A_PtrSize = 8) ? 104 : 84, 0)
        DllCall("GetObject", "UPtr", hBitmap, "int", oi.size, "UPtr", oi.ptr)
        hdib := DllCall("GlobalAlloc", "uint", 2, "UPtr", 40 + NumGet(oi, off1, "UInt"), "UPtr")
        pdib := DllCall("GlobalLock", "UPtr", hdib, "UPtr")
        DllCall("RtlMoveMemory", "UPtr", pdib, "UPtr", oi.ptr + off2, "UPtr", 40)
        DllCall("RtlMoveMemory", "UPtr", pdib + 40, "UPtr", NumGet(oi, off2 - A_PtrSize, "UPtr"), "UPtr", NumGet(oi, off1, "UInt"))
        DllCall("GlobalUnlock", "UPtr", hdib)
        this.DeleteObject(hBitmap)
        r3 := DllCall("SetClipboardData", "uint", 8, "UPtr", hdib) ; CF_DIB = 8
        DllCall("CloseClipboard")
        DllCall("GlobalFree", "UPtr", hdib)
        E := r3 ? 0 : -4    ; 0 - success
        Return E
    }

    Static CreateCompatibleDC(hdc := 0)
    {
        Return DllCall("CreateCompatibleDC", "UPtr", hdc)
    }

    Static CreateDIBSection(w, h, hdc := "", bpp := 32, &ppvBits := 0, Usage := 0, hSection := 0, Offset := 0)
    {
        hdc2 := hdc ? hdc : this.GetDC()
        bi := Buffer(40, 0)
        NumPut("UInt", 40, bi, 0)
        NumPut("UInt", w, bi, 4)
        NumPut("UInt", h, bi, 8)
        NumPut("UShort", 1, bi, 12)
        NumPut("UShort", bpp, bi, 14)
        NumPut("UInt", 0, bi, 16)

        hbm := DllCall("CreateDIBSection"
            , "UPtr", hdc2
            , "UPtr", bi.ptr    ; BITMAPINFO
            , "uint", Usage
            , "UPtr*", &ppvBits
            , "UPtr", hSection
            , "uint", Offset, "UPtr")

        If !hdc
            this.ReleaseDC(hdc2)
        Return hbm
    }

    Static SelectObject(hdc, hgdiobj)
    {
        Return DllCall("SelectObject", "UPtr", hdc, "UPtr", hgdiobj)
    }

    Static BitBlt(ddc, dx, dy, dw, dh, sdc, sx, sy, raster := "")
    {
        Return DllCall("gdi32\BitBlt"
            , "UPtr", ddc
            , "int", dx, "int", dy
            , "int", dw, "int", dh
            , "UPtr", sdc
            , "int", sx, "int", sy
            , "uint", raster ? raster : 0x00CC0020)
    }

    Static CreateBitmapFromHBITMAP(hBitmap, hPalette := 0)
    {
        DllCall("gdiplus\GdipCreateBitmapFromHBITMAP", "UPtr", hBitmap, "UPtr", hPalette, "UPtr*", &pBitmap := 0)
        Return pBitmap
    }

    Static CreateHBITMAPFromBitmap(pBitmap, Background := 0xffffffff)
    {
        DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "UPtr", pBitmap, "UPtr*", &hBitmap := 0, "int", Background)
        Return hBitmap
    }

    Static DeleteObject(hObject)
    {
        Return DllCall("DeleteObject", "UPtr", hObject)
    }

    Static ReleaseDC(hdc, hwnd := 0)
    {
        Return DllCall("ReleaseDC", "UPtr", hwnd, "UPtr", hdc)
    }

    Static DeleteDC(hdc)
    {
        Return DllCall("DeleteDC", "UPtr", hdc)
    }

    Static DisposeImage(pBitmap, noErr := 0)
    {
        If (StrLen(pBitmap) <= 2 && noErr = 1)
            Return 0

        r := DllCall("gdiplus\GdipDisposeImage", "UPtr", pBitmap)
        If (r = 2 || r = 1) && (noErr = 1)
            r := 0
        Return r
    }

    Static GetDC(hwnd := 0)
    {
        Return DllCall("GetDC", "UPtr", hwnd)
    }

    Static GetDCEx(hwnd, flags := 0, hrgnClip := 0)
    {
        Return DllCall("GetDCEx", "UPtr", hwnd, "UPtr", hrgnClip, "int", flags)
    }

    Static GetWindowRect(hwnd, &W, &H)
    {
        rect := Buffer(16, 0)
        er := DllCall("dwmapi\DwmGetWindowAttribute"
            , "UPtr", hwnd        ; HWND  hwnd
            , "UInt", 9           ; DWORD dwAttribute (DWMWA_EXTENDED_FRAME_BOUNDS)
            , "UPtr", rect.ptr    ; PVOID pvAttribute
            , "UInt", rect.size   ; DWORD cbAttribute
            , "UInt")             ; HRESULT

        If er
            DllCall("GetWindowRect", "UPtr", hwnd, "UPtr", rect.ptr, "UInt")

        r := {}
        r.x1 := NumGet(rect, 0, "Int"), r.y1 := NumGet(rect, 4, "Int")
        r.x2 := NumGet(rect, 8, "Int"), r.y2 := NumGet(rect, 12, "Int")
        r.w := Abs(Max(r.x1, r.x2) - Min(r.x1, r.x2))
        r.h := Abs(Max(r.y1, r.y2) - Min(r.y1, r.y2))
        W := r.w, H := r.h
        Return r
    }

    Static GraphicsFromHDC(hDC, hDevice := "", InterpolationMode := "", SmoothingMode := "", PageUnit := "", CompositingQuality := "")
    {
        If hDevice
            DllCall("Gdiplus\GdipCreateFromHDC2", "UPtr", hDC, "UPtr", hDevice, "UPtr*", &pGraphics := 0)
        Else
            DllCall("gdiplus\GdipCreateFromHDC", "UPtr", hDC, "UPtr*", &pGraphics := 0)

        If pGraphics
        {
            If (InterpolationMode != "")
                this.SetInterpolationMode(pGraphics, InterpolationMode)
            If (SmoothingMode != "")
                this.SetSmoothingMode(pGraphics, SmoothingMode)
            If (PageUnit != "")
                this.SetPageUnit(pGraphics, PageUnit)
            If (CompositingQuality != "")
                this.SetCompositingQuality(pGraphics, CompositingQuality)
        }

        Return pGraphics
    }

    Static SetInterpolationMode(pGraphics, InterpolationMode)
    {
        Return DllCall("gdiplus\GdipSetInterpolationMode", "UPtr", pGraphics, "int", InterpolationMode)
    }

    Static SetSmoothingMode(pGraphics, SmoothingMode)
    {
        Return DllCall("gdiplus\GdipSetSmoothingMode", "UPtr", pGraphics, "int", SmoothingMode)
    }

    Static SetPageUnit(pGraphics, Unit)
    {
        Return DllCall("gdiplus\GdipSetPageUnit", "UPtr", pGraphics, "int", Unit)
    }

    Static SetCompositingQuality(pGraphics, CompositionQuality)
    {
        Return DllCall("gdiplus\GdipSetCompositingQuality", "UPtr", pGraphics, "int", CompositionQuality)
    }

    Static DrawImage(pGraphics, pBitmap, dx := "", dy := "", dw := "", dh := "", sx := "", sy := "", sw := "", sh := "", Matrix := 1, Unit := 2, ImageAttr := 0)
    {
        usrImageAttr := 0
        If !ImageAttr
        {
            If !IsNumber(Matrix)
                ImageAttr := this.SetImageAttributesColorMatrix(Matrix)
            Else If (Matrix != 1)
                ImageAttr := this.SetImageAttributesColorMatrix("1|0|0|0|0|0|1|0|0|0|0|0|1|0|0|0|0|0|" Matrix "|0|0|0|0|0|1")
        } Else usrImageAttr := 1

        If (dx != "" && dy != "" && dw = "" && dh = "" && sx = "" && sy = "" && sw = "" && sh = "")
        {
            sx := sy := 0
            sw := dw := this.GetImageWidth(pBitmap)
            sh := dh := this.GetImageHeight(pBitmap)
        } Else If (sx = "" && sy = "" && sw = "" && sh = "")
        {
            If (dx = "" && dy = "" && dw = "" && dh = "")
            {
                sx := dx := 0, sy := dy := 0
                sw := dw := this.GetImageWidth(pBitmap)
                sh := dh := this.GetImageHeight(pBitmap)
            } Else
            {
                sx := sy := 0
                this.GetImageDimensions(pBitmap, &sw, &sh)
            }
        }

        _E := DllCall("gdiplus\GdipDrawImageRectRect"
            , "UPtr", pGraphics
            , "UPtr", pBitmap
            , "float", dx, "float", dy
            , "float", dw, "float", dh
            , "float", sx, "float", sy
            , "float", sw, "float", sh
            , "int", Unit
            , "UPtr", ImageAttr ? ImageAttr : 0
            , "UPtr", 0, "UPtr", 0)

        If (ImageAttr && usrImageAttr != 1)
            this.DisposeImageAttributes(ImageAttr)

        Return _E
    }

    Static CreateImageAttributes()
    {
        DllCall("gdiplus\GdipCreateImageAttributes", "UPtr*", &ImageAttr := 0)
        Return ImageAttr
    }

    Static DisposeImageAttributes(ImageAttr)
    {
        Return DllCall("gdiplus\GdipDisposeImageAttributes", "UPtr", ImageAttr)
    }

    Static SetImageAttributesColorMatrix(clrMatrix, ImageAttr := 0, grayMatrix := 0, ColorAdjustType := 1, fEnable := 1, ColorMatrixFlag := 0)
    {
        GrayscaleMatrix := 0

        If (StrLen(clrMatrix) < 5 && ImageAttr)
            Return -1

        If StrLen(clrMatrix) < 5
            Return

        ColourMatrix := Buffer(100, 0)
        Matrix := RegExReplace(RegExReplace(clrMatrix, "^[^\d-\.]+([\d\.])", "$1", , 1), "[^\d-\.]+", "|")
        Matrix := StrSplit(Matrix, "|")
        Loop 25
        {
            M := (Matrix[A_Index] != "") ? Matrix[A_Index] : Mod(A_Index - 1, 6) ? 0 : 1
            NumPut("Float", M, ColourMatrix, (A_Index - 1) * 4)
        }

        Matrix := ""
        Matrix := RegExReplace(RegExReplace(grayMatrix, "^[^\d-\.]+([\d\.])", "$1", , 1), "[^\d-\.]+", "|")
        Matrix := StrSplit(Matrix, "|")
        If (StrLen(Matrix) > 2 && ColorMatrixFlag = 2)
        {
            GrayscaleMatrix := Buffer(100, 0)
            Loop 25
            {
                M := (Matrix[A_Index] != "") ? Matrix[A_Index] : Mod(A_Index - 1, 6) ? 0 : 1
                NumPut("Float", M, GrayscaleMatrix, (A_Index - 1) * 4)
            }
        }

        If !ImageAttr
        {
            created := 1
            ImageAttr := this.CreateImageAttributes()
        }

        E := DllCall("gdiplus\GdipSetImageAttributesColorMatrix"
            , "UPtr", ImageAttr
            , "int", ColorAdjustType
            , "int", fEnable
            , "UPtr", ColourMatrix.ptr
            , "UPtr", GrayscaleMatrix ? GrayscaleMatrix.ptr : 0
            , "int", ColorMatrixFlag)

        E := created = 1 ? ImageAttr : E
        Return E
    }

    Static GetImageDimensions(pBitmap, &Width, &Height)
    {
        If StrLen(pBitmap) < 3
            Return -1

        Width := 0, Height := 0
        E := this.GetImageDimension(pBitmap, &Width, &Height)
        Width := Round(Width)
        Height := Round(Height)
        Return E
    }

    Static GetImageDimension(pBitmap, &w, &h)
    {
        Return DllCall("gdiplus\GdipGetImageDimension", "UPtr", pBitmap, "float*", &w := 0, "float*", &h := 0)
    }

    Static GetImageWidth(pBitmap)
    {
        DllCall("gdiplus\GdipGetImageWidth", "UPtr", pBitmap, "uint*", &Width := 0)
        Return Width
    }

    Static GetImageHeight(pBitmap)
    {
        DllCall("gdiplus\GdipGetImageHeight", "UPtr", pBitmap, "uint*", &Height := 0)
        Return Height
    }

    Static DrawRectangle(pGraphics, pPen, x, y, w, h)
    {
        Return DllCall("gdiplus\GdipDrawRectangle", "UPtr", pGraphics, "UPtr", pPen, "float", x, "float", y, "float", w, "float", h)
    }

    Static UpdateLayeredWindow(hwnd, hdc, x := "", y := "", w := "", h := "", Alpha := 255)
    {
        If ((x != "") && (y != ""))
            pt := Buffer(8, 0), NumPut("UInt", x, pt, 0), NumPut("UInt", y, pt, 4)

        If (w = "") || (h = "")
            this.GetWindowRect(hwnd, &w, &h)

        Return DllCall("UpdateLayeredWindow"
            , "UPtr", hwnd                                  ; layered window hwnd
            , "UPtr", 0                                     ; hdcDst (screen) - usually 0
            , "UPtr", ((x = "") && (y = "")) ? 0 : pt.ptr   ; POINT x,y of layered window
            , "int64*", w | h << 32                         ; SIZE w,h of layered window
            , "UPtr", hdc                                   ; hdcSrc - source bitmap to be drawn on to layered window - NULL if not changing
            , "int64*", 0                                   ; x,y offset of bitmap to be drawn
            , "uint", 0                                     ; crKey - bgcolor to use?  meaningless when using full alpha
            , "UInt*", Alpha << 16 | 1 << 24                ;
            , "uint", 2)
    }

    Static CreatePen(ARGB, w, Unit := 2)
    {
        E := DllCall("gdiplus\GdipCreatePen1", "UInt", ARGB, "float", w, "int", Unit, "UPtr*", &pPen := 0)
        Return pPen
    }

    Static DeletePen(pPen)
    {
        Return DllCall("gdiplus\GdipDeletePen", "UPtr", pPen)
    }

    Static DeleteGraphics(pGraphics)
    {
        Return DllCall("gdiplus\GdipDeleteGraphics", "UPtr", pGraphics)
    }

    Static SaveBitmapToFile(pBitmap, sOutput, Quality := 75, toBase64 := 0)
    {
        _p := 0

        SplitPath sOutput, , , &Extension
        If !RegExMatch(Extension, "^(?i:BMP|DIB|RLE|JPG|JPEG|JPE|JFIF|GIF|TIF|TIFF|PNG)$")
            Return -1

        Extension := "." Extension
        DllCall("gdiplus\GdipGetImageEncodersSize", "uint*", &nCount := 0, "uint*", &nSize := 0)
        ci := Buffer(nSize)
        DllCall("gdiplus\GdipGetImageEncoders", "uint", nCount, "uint", nSize, "UPtr", ci.ptr)
        If !(nCount && nSize)
            Return -2

        Static IsUnicode := StrLen(Chr(0xFFFF))
        If (IsUnicode)
        {
            StrGet_Name := "StrGet"
            Loop nCount
            {
                sString := %StrGet_Name%(NumGet(ci, (idx := (48 + 7 * A_PtrSize) * (A_Index - 1)) + 32 + 3 * A_PtrSize, "UPtr"), "UTF-16")
                If !InStr(sString, "*" Extension)
                    Continue

                pCodec := ci.ptr + idx
                Break
            }
        } Else
        {
            Loop nCount
            {
                Location := NumGet(ci, 76 * (A_Index - 1) + 44, "UPtr")
                nSize := DllCall("WideCharToMultiByte", "uint", 0, "uint", 0, "uint", Location, "int", -1, "uint", 0, "int", 0, "uint", 0, "uint", 0)
                sString := Buffer(nSize)
                DllCall("WideCharToMultiByte", "uint", 0, "uint", 0, "uint", Location, "int", -1, "str", sString, "int", nSize, "uint", 0, "uint", 0)
                If !InStr(sString, "*" Extension)
                    Continue

                pCodec := ci.ptr + 76 * (A_Index - 1)
                Break
            }
        }

        If !pCodec
            Return -3

        If (Quality != 75)
        {
            Quality := (Quality < 0) ? 0 : (Quality > 100) ? 100 : Quality
            If (Quality > 90 && toBase64 = 1)
                Quality := 90

            If RegExMatch(Extension, "^\.(?i:JPG|JPEG|JPE|JFIF)$")
            {
                DllCall("gdiplus\GdipGetEncoderParameterListSize", "UPtr", pBitmap, "UPtr", pCodec, "uint*", &nSize)
                EncoderParameters := Buffer(nSize, 0)
                DllCall("gdiplus\GdipGetEncoderParameterList", "UPtr", pBitmap, "UPtr", pCodec, "uint", nSize, "UPtr", EncoderParameters.ptr)
                nCount := NumGet(EncoderParameters, "UInt")
                Loop nCount
                {
                    elem := (24 + A_PtrSize) * (A_Index - 1) + 4 + (pad := (A_PtrSize = 8) ? 4 : 0)
                    If (NumGet(EncoderParameters, elem + 16, "UInt") = 1) && (NumGet(EncoderParameters, elem + 20, "UInt") = 6)
                    {
                        _p := elem + EncoderParameters.ptr - pad - 4
                        NumPut(Quality, NumGet(NumPut(4, NumPut(1, _p + 0, "UPtr") + 20, "UInt"), "UPtr"), "UInt")
                        Break
                    }
                }
            }
        }

        If (toBase64 = 1)
        {
            ; part of the function extracted from ImagePut by iseahound
            ; https://www.autohotkey.com/boards/viewtopic.php?f=6&t=76301&sid=bfb7c648736849c3c53f08ea6b0b1309
            DllCall("ole32\CreateStreamOnHGlobal", "UPtr", 0, "int", true, "UPtr*", &pStream := 0)
            _E := DllCall("gdiplus\GdipSaveImageToStream", "UPtr", pBitmap, "UPtr", pStream, "UPtr", pCodec, "uint", _p)
            If _E
                Return -6

            DllCall("ole32\GetHGlobalFromStream", "UPtr", pStream, "uint*", &hData)
            pData := DllCall("GlobalLock", "UPtr", hData, "UPtr")
            nSize := DllCall("GlobalSize", "uint", pData)

            bin := Buffer(nSize, 0)
            DllCall("RtlMoveMemory", "UPtr", bin.ptr, "UPtr", pData, "uptr", nSize)
            DllCall("GlobalUnlock", "UPtr", hData)
            ObjRelease(pStream)
            DllCall("GlobalFree", "UPtr", hData)

            ; Using CryptBinaryToStringA saves about 2MB in memory.
            DllCall("Crypt32.dll\CryptBinaryToStringA", "UPtr", bin.ptr, "uint", nSize, "uint", 0x40000001, "UPtr", 0, "uint*", &base64Length := 0)
            base64 := Buffer(base64Length, 0)
            _E := DllCall("Crypt32.dll\CryptBinaryToStringA", "UPtr", bin.ptr, "uint", nSize, "uint", 0x40000001, "UPtr", &base64, "uint*", base64Length)
            If !_E
                Return -7

            bin := Buffer(0)
            Return StrGet(base64, base64Length, "CP0")
        }

        _E := DllCall("gdiplus\GdipSaveImageToFile", "UPtr", pBitmap, "WStr", sOutput, "UPtr", pCodec, "uint", _p)
        Return _E ? -5 : 0
    }
    Static SavePrintScreen()
    {
        this.Startup()
        pBitmap := this.BitmapFromScreen({ X: 0, Y: 0, W: A_ScreenWidth, H: A_ScreenHeight })
        this.SaveBitmapToFile(pBitmap,  A_Desktop '\' A_ScriptName '.png')
        this.DisposeImage(pBitmap)
        this.Shutdown()
    }
}