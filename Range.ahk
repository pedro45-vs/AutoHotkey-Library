#Requires AutoHotkey v2.0

; https://www.autohotkey.com/boards/viewtopic.php?f=96&t=88725

Range(Start, Stop, Step:=1) => (&n) => (n := Start, Start += Step, Step > 0 ? n <= Stop : n >= Stop)
