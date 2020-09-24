<div align="center">

## Trick VB into making small bit arrays by controlling its array management


</div>

### Description

Replace the array descriptor that VB uses with your own descriptor structure, and make it point to your data. Here is an example that forces VB to create an array of bytes that only takes 2 bytes per array element instead of 6, plus the descriptor size. Lots of comments. Please don't forget to vote!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Xavier Harel](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/xavier-harel.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/xavier-harel-trick-vb-into-making-small-bit-arrays-by-controlling-its-array-management__1-14905/archive/master.zip)





### Source Code

```
'Paste all this in the form code of a new project, run it, and step into it to follow the process:
' This is a redeclaration of VB's VarPtr that forces it to return the address of the array descriptor structure:
Private Declare Function GetArrayPtr Lib "msvbvm60.dll" Alias "VarPtr" ( _
 Ptr() As Any _
) As Long
' This is a translation into VB code of C++'s safearray descriptor structure:
Private Type SafeArrayBound
 lNumOfElements As Long
 lLowBound As Long
End Type
Private Type SafeArr
 iDimensions As Integer
 iFeatures As Integer
 lElementSize As Long
 lLocks As Long
 lDataPtr As Long
 saBound(0) As SafeArrayBound
End Type
Private Const FADF_AUTO = &H1		' Array is allocated on the stack.
Private Const FADF_FIXEDSIZE = &H10	' Array may not be resized or reallocated.
Private Sub FillMyBytesArray()
Dim Bytes() As Byte ' creates an array descriptor of type SafeArr pointing to no data
Dim sMyString As String ' will hold data that I'll use as the data that Bytes() is pointing to
Dim aMySAB(0) As SafeArrayBound
Dim aMySA As SafeArr
Dim sResult As String
Dim i As Integer
 sMyString = "This is a relatively short string"
 ' create the descriptor that will replace the Bytes() array descriptor declared above
 With aMySAB(0) ' Description of an array dimension (size and lbound)
 ' the string is stored as unicode, which means that the 1st word is stored as "T" + chr(0) + "h" + chr(0) + "i" + chr(0) + "s" + chr(0)
 ' so that there are really twice as many bytes stored as the length of the string:
 .lNumOfElements = 2 * Len(sMyString) ' number of elements in this array dimension
 .lLowBound = 0 ' specifies the array's Lbound value
 End With
 With aMySA
 ' this is a 1-dimension byte array:
 .iDimensions = 1
 .lElementSize = 1 ' size of each element
 .iFeatures = FADF_AUTO Or FADF_FIXEDSIZE ' Flags that enable array features.
 .lDataPtr = VarPtr(ByVal sMyString) ' make the descriptor point to the declared string data. ByVal is VERY important.
 .saBound(0) = aMySAB(0) ' describes each dimension of the array, in this case only one.
 End With
 ' move the memory contents of the descriptor to the address of the Bytes() array descriptor, the ByVal is VERY important if you don't want to overwrite memory and risk a crash!
 CopyMemory ByVal GetArrayPtr(Bytes), VarPtr(aMySA), 4
 ' Reattach all the bytes together to reconstruct sMyString. Notice that Bytes() now has data, and that there is no error calling Ubound(Bytes):
 For i = 0 To UBound(Bytes)
 sResult = sResult & Chr(Bytes(i))
 Next i
 ' Since we read the string data directly from memory, we have unicode, and we have to disregard all odd array indexes, or convert the result string from unicode:
 sResult = StrConv(sResult, vbFromUnicode)
 ' now sResult contains the same data as sMyString.
End Sub
Private Suv Form_Load()
 FillMyBytesArray ' call the sub above
End Sub
' I will post more on this topic if it becomes popular
```

