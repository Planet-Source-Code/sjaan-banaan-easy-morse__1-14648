Attribute VB_Name = "Module2"
Option Explicit
Const Good As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890.,:?-'() """
Type Morse
    lencode As Long
    code() As Bit
    space As Boolean
End Type

Enum Bit
    [lang]
    [kort]
End Enum
Function retrBits(Char) As Morse
Dim Msg As Morse
Msg.space = False
ReDim Msg.code(6)
Select Case Char
Case "A", "a": With Msg: .lencode = 2: .code(0) = [kort]: .code(1) = [lang]: End With
Case "B", "b": With Msg: .lencode = 4: .code(0) = [lang]: .code(1) = [kort]: .code(2) = [kort]: .code(3) = [kort]: End With
Case "C", "c": With Msg: .lencode = 4: .code(0) = [lang]: .code(1) = [kort]: .code(2) = [lang]: .code(3) = [kort]: End With
'Case "CH", "ch": With Msg: .lencode = 4: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [lang]: .code(3) = [lang]: End With
Case "D", "d": With Msg: .lencode = 3: .code(0) = [lang]: .code(1) = [kort]: .code(2) = [kort]: End With
Case "E", "e": With Msg: .lencode = 1: .code(0) = [kort]: End With
Case "F", "f": With Msg: .lencode = 4: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [lang]: .code(3) = [kort]: End With
Case "G", "g": With Msg: .lencode = 3: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [kort]: End With
Case "H", "h": With Msg: .lencode = 4: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [kort]: .code(3) = [kort]: End With
Case "I", "i": With Msg: .lencode = 2: .code(0) = [kort]: .code(1) = [kort]: End With
Case "J", "j": With Msg: .lencode = 4: .code(0) = [kort]: .code(1) = [lang]: .code(2) = [lang]: .code(3) = [lang]: End With
Case "K", "k": With Msg: .lencode = 3: .code(0) = [lang]: .code(1) = [kort]: .code(2) = [lang]: End With
Case "L", "l": With Msg: .lencode = 4: .code(0) = [kort]: .code(1) = [lang]: .code(2) = [kort]: .code(3) = [kort]: End With
Case "M", "m": With Msg: .lencode = 2: .code(0) = [lang]: .code(1) = [lang]: End With
Case "N", "n": With Msg: .lencode = 2: .code(0) = [lang]: .code(1) = [kort]: End With
Case "O", "o": With Msg: .lencode = 3: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [lang]: End With
Case "P", "p": With Msg: .lencode = 4: .code(0) = [kort]: .code(1) = [lang]: .code(2) = [lang]: .code(3) = [kort]: End With
Case "Q", "q": With Msg: .lencode = 4: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [kort]: .code(3) = [lang]: End With
Case "R", "r": With Msg: .lencode = 3: .code(0) = [kort]: .code(1) = [lang]: .code(2) = [kort]: End With
Case "S", "s": With Msg: .lencode = 3: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [kort]: End With
Case "T", "t": With Msg: .lencode = 1: .code(0) = [lang]: End With
Case "U", "u": With Msg: .lencode = 3: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [lang]: End With
Case "V", "v": With Msg: .lencode = 4: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [kort]: .code(3) = [lang]: End With
Case "W", "w": With Msg: .lencode = 3: .code(0) = [kort]: .code(1) = [lang]: .code(2) = [lang]: End With
Case "X", "x": With Msg: .lencode = 4: .code(0) = [lang]: .code(1) = [kort]: .code(2) = [kort]: .code(3) = [lang]: End With
Case "Y", "y": With Msg: .lencode = 4: .code(0) = [lang]: .code(1) = [kort]: .code(2) = [lang]: .code(3) = [lang]: End With
Case "Z", "z": With Msg: .lencode = 4: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [kort]: .code(3) = [kort]: End With
Case "Ä", "ä": With Msg: .lencode = 4: .code(0) = [kort]: .code(1) = [lang]: .code(2) = [kort]: .code(3) = [lang]: End With
Case "Ö", "ö": With Msg: .lencode = 4: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [lang]: .code(3) = [kort]: End With
Case "Ü", "ü": With Msg: .lencode = 4: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [lang]: .code(3) = [lang]: End With

Case "1": With Msg: .lencode = 5: .code(0) = [kort]: .code(1) = [lang]: .code(2) = [lang]: .code(3) = [lang]: .code(4) = [lang]: End With
Case "2": With Msg: .lencode = 5: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [lang]: .code(3) = [lang]: .code(4) = [lang]: End With
Case "3": With Msg: .lencode = 5: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [kort]: .code(3) = [lang]: .code(4) = [lang]: End With
Case "4": With Msg: .lencode = 5: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [kort]: .code(3) = [kort]: .code(4) = [lang]: End With
Case "5": With Msg: .lencode = 5: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [kort]: .code(3) = [kort]: .code(4) = [kort]: End With
Case "6": With Msg: .lencode = 5: .code(0) = [lang]: .code(1) = [kort]: .code(2) = [kort]: .code(3) = [kort]: .code(4) = [kort]: End With
Case "7": With Msg: .lencode = 5: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [kort]: .code(3) = [kort]: .code(4) = [kort]: End With
Case "8": With Msg: .lencode = 5: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [lang]: .code(3) = [kort]: .code(4) = [kort]: End With
Case "9": With Msg: .lencode = 5: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [lang]: .code(3) = [lang]: .code(4) = [kort]: End With
Case "0": With Msg: .lencode = 5: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [lang]: .code(3) = [lang]: .code(4) = [lang]: End With

Case ".": With Msg: .lencode = 6: .code(0) = [kort]: .code(1) = [lang]: .code(2) = [kort]: .code(3) = [lang]: .code(4) = [kort]: .code(5) = [lang]: End With
Case ",": With Msg: .lencode = 6: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [kort]: .code(3) = [kort]: .code(4) = [lang]: .code(5) = [lang]: End With
Case ":": With Msg: .lencode = 6: .code(0) = [lang]: .code(1) = [lang]: .code(2) = [lang]: .code(3) = [kort]: .code(4) = [kort]: .code(5) = [kort]: End With
Case "?": With Msg: .lencode = 6: .code(0) = [kort]: .code(1) = [kort]: .code(2) = [lang]: .code(3) = [lang]: .code(4) = [kort]: .code(5) = [kort]: End With
Case "-": With Msg: .lencode = 6: .code(0) = [lang]: .code(1) = [kort]: .code(2) = [kort]: .code(3) = [kort]: .code(4) = [kort]: .code(5) = [lang]: End With
Case "'": With Msg: .lencode = 6: .code(0) = [kort]: .code(1) = [lang]: .code(2) = [lang]: .code(3) = [lang]: .code(4) = [lang]: .code(5) = [kort]: End With
Case "(": With Msg: .lencode = 6: .code(0) = [lang]: .code(1) = [kort]: .code(2) = [lang]: .code(3) = [lang]: .code(4) = [kort]: .code(5) = [kort]: End With
Case ")": With Msg: .lencode = 6: .code(0) = [lang]: .code(1) = [kort]: .code(2) = [lang]: .code(3) = [lang]: .code(4) = [kort]: .code(5) = [lang]: End With
Case """": With Msg: .lencode = 6: .code(0) = [kort]: .code(1) = [lang]: .code(2) = [kort]: .code(3) = [kort]: .code(4) = [lang]: .code(5) = [kort]: End With
Case " ": Msg.space = True
End Select
retrBits = Msg
End Function
Function CheckChar(Char) As Boolean
Dim cha As String * 1
cha = UCase(Char)
CheckChar = (InStr(Good, cha) <> 0)
End Function

Function CreateItString(string1)
Dim I, X, Msg
For I = 1 To Len(string1)
X = UCase(Mid(string1, I, 1))
If CheckChar(X) Then Msg = Msg & X
Next
CreateItString = Msg
End Function

Function CreateIt(string1, Temp() As Morse)
Dim Y, X, I, kees As Morse, S
Y = CreateItString(string1)
ReDim Temp(Len(Y))
For I = 1 To Len(Y)
X = Mid(Y, I, 1)
kees = retrBits(X)
Temp(I - 1) = kees
Next
End Function

Function SendIt(morses() As Morse, AfstandMS)
Dim X, Y, I, S, Letter As Boolean
For I = LBound(morses) To UBound(morses)
    For S = 0 To morses(I).lencode - 1
    If morses(I).space = False Then
    X = morses(I).code(S)
    Y = IIf(X = [lang], AfstandMS * 3, AfstandMS)
    '1000 is just a freq, you can chose an other one
    BeepSound 1000, Y
    Wait AfstandMS
    Letter = True
    Else
    Wait AfstandMS * 6 'A space is the length of 6 points
    Letter = False
    End If
    Next
    If Letter And (I <> UBound(morses)) Then Wait AfstandMS * 3 ' wait 3 points between 2 letters
Next
End Function

Sub MorseIt(Text, AfstandMS)
Dim X() As Morse
Call CreateIt(Text, X)
Call SendIt(X, AfstandMS)
End Sub
