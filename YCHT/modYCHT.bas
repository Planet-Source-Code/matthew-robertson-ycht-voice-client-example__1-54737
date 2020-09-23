Attribute VB_Name = "modYCHT"
'modYCHT By: Matthew Robertson

Type typCookie
    Y       As String
    T       As String
End Type
Type typYCHT
    ID      As String
    PW      As String
    Serv    As String
    Room    As String
    VCAuth  As String
    RMSpace As String
End Type
Global YCookie As typCookie
Global YCHT As typYCHT

Function ChatSend(Text As String, Room As String)
ChatSend = Packet(41, Room & Chr(1) & Text)
End Function

Function CheckPassword(Data As String) As Boolean
If InStr(Data, "Invalid Password") Then
    MsgBox "Wrong Password!"
    CheckPassword = False
ElseIf InStr(Data, "This Yahoo! ID does not exist") Then
    MsgBox "The Yahoo! ID does not exist!"
    CheckPassword = False
Else
    CheckPassword = True
End If
End Function
Function Join(Room As String)
Join = Packet("11", Room)
End Function

Function Packet(PackType As String, Pack As String)
Packet = "YCHT" & String(2, 0) & Chr(1) & String(4, 0) & Chr("&h" & PackType) & String(2, 0) & Chr(Fix(Len(Pack) / 256)) & Chr(Len(Pack) Mod 256) & Pack
End Function

Function Login(ID As String, YCookie As String, TCookie As String)
Login = Packet("1", ID & Chr(1) & YCookie & Chr(32) & TCookie)
End Function

Sub GetCookies(Data As String)
Dim sptData() As String
sptData = Split(Data, "Set-Cookie: ")
For i = 1 To UBound(sptData)
    Data = sptData(i)
    Data = Left(Data, InStr(Data, ";"))
    If Left(Data, 1) = "Y" Then YCookie.Y = Data
    If Left(Data, 1) = "T" Then YCookie.T = Data
Next
End Sub

Function HexPassword(PW As String)
' turns a pass into a string of hex charcters.
' this servs 2 purposes, 1 it allows odd charcters in the password,
' and 2 it makes the password harder to notice to some1 watching ur packets.
Dim HPW As String
For i = 1 To Len(PW)
    HPW = HPW & "%" & Hex(Asc(Mid(PW, i, 1)))
Next
HexPassword = HPW
End Function
Function Ping()
Ping = Packet("64", "")
End Function


