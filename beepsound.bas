Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function ReadPort Lib "io.dll" (ByVal Address As Long) As Byte
Private Declare Sub WritePort Lib "io.dll" (ByVal Address As Long, ByVal Value As Byte)
Private Declare Function GetTickCount Lib "kernel32" () As Long

Function BeepSound(Freq, MsDur)
    Dim Result&, Lo As Byte, Hi As Byte
    Dim Times As Long
    Result& = Freq
     If Result > 18 And Result < 20000 Then
      Result = 1193180 / Result
      Lo = Result And &HFF&
      Hi = Result \ &H100&
    
      Call WritePort(&H43, &HB6&)
      Call WritePort(&H42, Lo)
      Call WritePort(&H42, Hi)
      
      Result = ReadPort(&H61&)
      Call WritePort(&H61&, Result Or &H3&)
      BeepSound = 1
    Else
      BeepSound = -1
      Exit Function
    End If
    Times = GetTickCount
    Do Until GetTickCount - Times >= MsDur
    DoEvents
    Loop
    Result = ReadPort(&H61&)
    Call WritePort(&H61&, Result And &HFC&)
End Function

Function Wait(Ms)
Dim Times As Long
Times = GetTickCount
Do Until GetTickCount - Times >= Ms
DoEvents
Loop
End Function
