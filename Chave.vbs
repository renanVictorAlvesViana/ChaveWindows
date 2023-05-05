Set WshShell = CreateObject("WScript.Shell")

MsgBox ConvertToKey(WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"))

Function ConvertToKey(Key)

Const KeyOffset = 52

I = 28

Chars = "BCDFGHJKMPQRTVWXY2346789"

Do

Cur = 0

x = 14

Do

Cur = Cur * 256

Cur = Key(x + KeyOffset) + Cur

Key(x + KeyOffset) = (Cur \ 24) And 255

Cur = Cur Mod 24

x = x - 1

Loop While x >= 0

I = I - 1

KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput

If (((29 - I) Mod 6) = 0) And (I <> -1) Then

I = I - 1

KeyOutput = "-" & KeyOutput

End If

Loop While I >= 0

ConvertToKey = KeyOutput

End Function