This little programm compares the code in Visual Basic and C# to output a value on an Excel sheet.

It also throws an error in the VB output code.
Seems like Cells.Value is no longer available when working with VB :(

Background:
I have a large VB project (created ~2015, couple thousand lines) which I try to migrate to .Net 6.0.
The migration worked well, only Cells.Value does not :(
