<div align="center">

## bytes to KB, MB or GB


</div>

### Description

with my code you can input a number of bytes and it will tell you how many Kilobytes Megabytes or Giga bytes it is equal to.
 
### More Info
 
it pretty self explanatory


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Orenstein](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-orenstein.md)
**Level**          |Advanced
**User Rating**    |3.2 (16 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-orenstein-bytes-to-kb-mb-or-gb__1-7347/archive/master.zip)

### API Declarations

```
Public Enum BYTEVALUES
  KiloByte = 1024
  MegaByte = 1048576
  GigaByte = 107374182
End Enum
```


### Source Code

```
Public Function CutDecimal(Number As String, ByPlace As Byte) As String
  Dim Dec As Byte
  Dec = InStr(1, Number, ".", vbBinaryCompare) ' find the Decimal
  If Dec = 0 Then
    CutDecimal = Number 'if there is no decimal Then dont do anything
    Exit Function
  End If
  CutDecimal = Mid(Number, 1, Dec + ByPlace) 'How many places you want after the decimal point
End Function
Function GiveByteValues(Bytes As Double) As String
  If Bytes < BYTEVALUES.KiloByte Then
    GiveByteValues = Bytes & " Bytes"
  ElseIf Bytes >= BYTEVALUES.GigaByte Then
    GiveByteValues = CutDecimal(Bytes / BYTEVALUES.GigaByte, 2) & " Gigabytes"
  ElseIf Bytes >= BYTEVALUES.MegaByte Then
    GiveByteValues = CutDecimal(Bytes / BYTEVALUES.MegaByte, 2) & " Megabytes"
  ElseIf Bytes >= BYTEVALUES.KiloByte Then
    GiveByteValues = CutDecimal(Bytes / BYTEVALUES.KiloByte, 2) & " Kilobytes"
  End If
End Function
```

