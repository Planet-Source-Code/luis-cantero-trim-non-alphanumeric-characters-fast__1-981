<div align="center">

## Trim non alphanumeric characters FAST


</div>

### Description

It will erase any non-alphanumeric characters from a string rapidly. Usefull if you want to check strings for non-valid characters.

Strings such as email or web addresses, you can even make so that only numbers can be entered in, for example, a text box.
 
### More Info
 
Any string

The filtered string


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Luis Cantero](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/luis-cantero.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/luis-cantero-trim-non-alphanumeric-characters-fast__1-981/archive/master.zip)





### Source Code

```
Function TrimVoid(strWhat)
'*************************
'Usage: x = TrimVoid(String)
'*************************
'Example: Chunk = TrimVoid(Chunk)
'Filters all non-alphanumeric characters from string "Chunk".
'*************************
For i = 1 To Len(strWhat)
If Mid(strWhat, i, 1) Like "[a-zA-Z0-9]" Then strNew = strNew & Mid(strWhat, i, 1)
Next
TrimVoid = strNew
End Function
'NOTES - replace the above code with the lines below to get the wanted results.
'For trimming email addresses use this:
'Like "[a-zA-Z0-9._-]"
'For trimming web addresses use this:
'Like "[a-zA-Z0-9._/-]"
'To accept only numbers in a text box use this in the text box's Change Sub:
'Like "[0-9]"
```

