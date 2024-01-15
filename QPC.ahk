#Requires AutoHotkey v2.0

QPC( R := 0 )  ;   By SKAN for ah2 on CT91/D497 @ goo.gl/nf7O4G
{ 
    static P := 0,  F := 0, Q := DllCall('kernel32.dll\QueryPerformanceFrequency', 'int64p', &F)
    return( DllCall('kernel32.dll\QueryPerformanceCounter','int64p',&Q) * 0 + ( R ? ( P := Q ) * 0 : ( Q - P ) / F) )
}