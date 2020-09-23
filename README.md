<div align="center">

## Has\.\.\. String Functions


</div>

### Description

This is a series of Has.. Codes. It will search a textbox for the following: Lowercase letter, Uppercase letters, Numeric Characters, and Accented Characters. So, if it has lowercase letters or something, it will display a message box. this is a good example of how to use for..next loops, ASCII codes and the Instr Function, then again, this might be a lousy example, you be the judge.
 
### More Info
 
ASCII Codes, atleast add on to it, search for ASCII Gen and you will find one of my ASCII code generators.

Message Box if the Function returns true


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jared Collums](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jared-collums.md)
**Level**          |Intermediate
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jared-collums-has-string-functions__1-6089/archive/master.zip)





### Source Code

```
Public Function HasUppercase(TextBox As Object)
For i = 65 To 90 'i equals every letter from "A" to "Z"
If InStr(TextBox.Text, Chr$(i)) Then MsgBox "Has Uppercase"
'Searches for letters A to Z (i), and if i is present, Display a box.
End Function
Public Function HasLowercase(TextBox As Object)
For i = 97 To 122 'i equals every letter from "a" to "z"
If InStr(TextBox.Text, Chr$(i)) Then MsgBox "Has Lowercase"
'Searches for letters a to z (i), and if i is present, Display a box.
Next i
End Function
Public Function HasNumeric(TextBox As Object)
For i = 0 To 9 'i equals every number from "0" to "9"
If InStr(TextBox.Text, i) Then MsgBox "Has Numeric"
'Searches for numbers 0 to 9 (i), and if i is present, Display a box.
Next i
End Function
Public Function HasAccentchars(TextBox As Object)
For i = 128 To 223 'i equals every character from "&#8364;" to "ß"
If InStr(TextBox.Text, Chr$(i)) Then MsgBox "Has Accented Characters"
'Searches for accent characters &#8364; to ß (i), and if i is present, Display a box.
Next i
End Function
```

