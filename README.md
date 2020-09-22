<div align="center">

## Speed up code with a factor 1000 using string buffering\!


</div>

### Description

The purpose of this code is to prevent performance leak when repeatingly appending text to a string "the normal way" eg. string1 = string1 & "text". By using stringbuffering you allocate a large ammount of zero-bytes to your string, and then uses the mid function to insert the text you want to append. The biggest gain by using stringbuffering is that you speed ud your code by a factor 1000 but it's allso harder to get your code to crash with a "Out of memory" error (IF YOU LIKE THE CODE, PLEASE VOTE!)
 
### More Info
 
The string you want to append, and an index value, telling which buffer you want to append to.

Use:

AppendToBuffer MyString, iIndex

instead of:

String1 = String1 & Mystring

the entire string, being your the string plus the text you appended-

None as far as i know :-)


<span>             |<span>
---                |---
**Submitted On**   |2002-07-16 15:34:20
**By**             |[David Bue Pedersen](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-bue-pedersen.md)
**Level**          |Advanced
**User Rating**    |4.8 (72 globes from 15 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Speed\_up\_c1067697162002\.zip](https://github.com/Planet-Source-Code/david-bue-pedersen-speed-up-code-with-a-factor-1000-using-string-buffering__1-36945/archive/master.zip)








