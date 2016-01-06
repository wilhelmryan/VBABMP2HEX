# VBABMP2HEX
Converts a spreadsheet from cell drawn BMP to base8 Hex

Example spreadsheet can be found here:
http://tiny.cc/2016Resume

Page 'Handwritten Binary'- Exactly as it sounds, manually inputting 0/1s... tedious and rigid, not automatically populating desired Hex codes

Page 'Computed Binary'- Again, exactly as it sounds. Utilizing VBA and excel formulas to streamline the process and make additions and modifications populate without handwriting binary. 

How to use: Select active page 'Computed Binary', Press 'Convert to Binary', done. 
Modifications can be made to the 5x7 images outlined in red by changing the background color to Black. Press 'Convert to Binary' again and it will automatically update the table on the right hand side under 'Hexadecimal'. 

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



-Jarrod Hiscock-Wagner Jan. 2016
