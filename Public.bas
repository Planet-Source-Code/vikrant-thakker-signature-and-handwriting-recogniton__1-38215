Attribute VB_Name = "Public"
Public BuffArray(100) As String    ' Creates an array of 100 elements
Public strFile As String           ' Stores the complete line in the array
Public Distort As Single
Public Difference As Integer       ' Gets the difference of the previous and the current position of mouse
Public Score() As Single           ' After comparison with the array, each unit is increased when a matched direction is found
Public Highest As Integer          ' The alphabet that has the highest score
Public HighScore As Single         ' The amount of highest score

' WriteLet is a boolean...
' If "writeLet = True" then this means that the user has not
' completed drawing a character.
' And if "WriteLet=False" then this means that user has completed
' drawing a character.

Public Ans As String    ' This variable is used to store the User Response for a 'Quit' message
Public tm As Integer    ' This stores the time taken by user to draw one character
Public Type LetterType
    Direc(100) As Integer
End Type

'Dim Alphabet(25) As LetterType
Public Alphabet(250) As LetterType
Public WriteLet As Boolean  ' This is used to know whether the user has completed drawing a character or not
Public HoldX As Integer, HoldY As Integer ' Store the current position value of X and Y
Public LetterMovement(200) As Integer
'Dim Letter(-1 To 25) As String
Public Letter() As String
Public NumLet As Integer
Public WriteFile As Boolean ' if true then allow  Read and write from the database file, if false then close the database file



Public Function Direction(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer) As Integer
' This Function returns the value of Direction by calculating
' old and current X and Y  values

'x1 and y1 are the center points
ReDim Letter(-1 To rsChar.RecordCount) As String
Dim Slope As Single


' Here we get old and current X (x1,x2) and Y (y1,y2) values,
' And on using the formula we calculate the slope
' Based on the return value of slope, we assign a particular
' value to the direction

' x1 is the old X value
'x2 is the current X value
If X2 - X1 = 0 Then
    Slope = 50
Else
    Slope = -(Y2 - Y1) / (X2 - X1)
End If


If Slope <= 0 And Slope > -0.5 Then
    Direction = 0
ElseIf Slope <= -0.5 And Slope > -1 Then
    Direction = 1
ElseIf Slope <= -1 And Slope > -2 Then
    Direction = 2
ElseIf Slope < -2 Then
    Direction = 3
ElseIf Slope > 2 Then
    Direction = 4
ElseIf Slope <= 2 And Slope > 1 Then
    Direction = 5
ElseIf Slope <= 1 And Slope > 0.5 Then
    Direction = 6
ElseIf Slope <= 0.5 And Slope > 0 Then
    Direction = 7
End If

' y1 is the Old Y value,
' y2 is the current Y value
If Y2 > Y1 Then
    Direction = Direction + 8
End If

End Function


Public Sub LoadAll()
Dim strFileLine As String
Dim Count As Integer
Dim i As Integer
Dim Start As Integer
ReDim Letter(-1 To rsChar.RecordCount) As String

'Each alphabet is stored in an array named Letter(-1 to rsChar.RecordCount)
Letter(-1) = ""

If rsChar.RecordCount > 0 Then
rsChar.MoveFirst
For i = 0 To rsChar.RecordCount - 1 Step 1
    Letter(i) = rsChar!Char
    If rsChar.EOF = False Then rsChar.MoveNext
Next
End If


Dim a As Integer
' Each Letter has 100 movements, each of 2 digits,
' So this part is to read the movements

' Due to this we get the different percentage of different characters, if we delete this part each character will show the same probaility

Start = 2
If rsChar.RecordCount > 0 Then
rsChar.MoveFirst
Count = 0
For i = 0 To rsChar.RecordCount - 1 Step 1
    strFileLine = rsChar!String
    For a = Start To 200 Step 2
        Alphabet(Count).Direc(Int(a / 2)) = Val(Mid(strFileLine, a, 2))
    Next a
    Start = 1
    Count = Count + 1
If rsChar.EOF = False Then rsChar.MoveNext
Next
End If
End Sub

Public Sub MouseDown()
' Store the current X and Y values when the mouse button is clicked down
'Tmr.Enabled = True
tm = 0
WriteLet = True
HoldX = X
HoldY = Y
End Sub

