# VBABMP2HEX
Converts a spreadsheet from cell drawn BMP to base8 Hex

Example spreadsheet can be found here:
http://tiny.cc/2016Resume

----------------------------------FUNCTIONALITY----------------------------------

The purpose of this VBA script is to convert a BMP image drawn in excel cells to a binary, then to base 8 hexadecimal.

There is currently no function in excel that allows you to make a decision based on the color of a cell. Simple VBA was necessary to manipulate cells interior formatting to then make that into a number representation standard excel formulas will work on. 

Key Concepts:

VBA
cell.Interior.ColorIndex

Built in formulas
Decimal to Hexadecimal,
Number to Text
String Concatenation
