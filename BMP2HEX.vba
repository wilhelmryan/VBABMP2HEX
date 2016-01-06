Private Sub CommandButton1_Click()
' VBA function on button press to convert a 'drawn' image to binary format
' Author Jarrod Hiscock-Wagner Jan.2016
' uses visual cues to show where the work is being done; showing 0/1s on the images is extraneous

'declare our desired color and range
Dim color As Integer, rng As Range, cell As Range

'retreive color from keyspace
color = Sheet2.Range("E1").Interior.ColorIndex

'assert our desired range
Set rng = Range("A5:BC77")

'iterate through each cell and output '0' or '1'

For Each cell In rng
    'cell is black
    If cell.Interior.ColorIndex = color Then cell.Value = 1
    'cell is blank
    If cell.Interior.ColorIndex = -4142 Then cell.Value = 0
    Next cell

'declare variables for text formatting
Dim rowperBMP As Integer, colmnperBMP As Integer, charPerline As Integer, rowOffset As Integer, colmnOffset As Integer
'declare ranges to use
Dim startBinary As Range, startOutput As Range

'use with our offsets to properly insert text
Set startOutput = Range("BF6")
Set startBinary = Range("B6")

'set our offsets
rowOffset = 8
colmnOffset = 6

'used for computing new character
rowperBMP = 6
colmnperBMP = 4
charPerline = 9

'i= total char, j=current char, k=column, l=row
Dim i As Integer, j As Integer, k As Integer, l As Integer
Dim result As Integer

'total number of characters to convert
i = 72

'for each char we want to transform
For j = 0 To i

    'for each row in that char
    For k = 0 To rowperBMP
    
    'declare a new value for row total and assign it to zero
    Dim rowTotal As Integer
    rowTotal = 0
    
        'for each column in that row...
        For l = 0 To colmnperBMP
            'declare a variable for our current row/column
            Dim columnTotal As Integer, append As Integer
            columnTotal = 0
            
            'append our current binary power
            'MSB to LSB, 5 bits
            Select Case l
                Case 0 'magnitude of binary power 2^x
                    append = 16
                Case 1
                    append = 8
                Case 2
                    append = 4
                Case 3
                    append = 2
                Case 4
                    append = 1
            End Select
            
            'add our current cell to the column total
            columnTotal = startBinary.Offset(((j \ charPerline) * rowOffset) + k, ((j Mod charPerline) * colmnOffset) + l).Value * append
            
            'update our row total /
            rowTotal = rowTotal + columnTotal
            'next column
             Next l
             
             startOutput.Offset(j, k).Value = rowTotal 'compute binary to here
        'next row
        Next k
        
    'next character
    Next j
             
'clear our binary output on image for clarity/sanity sake
rng.ClearContents

End Sub
