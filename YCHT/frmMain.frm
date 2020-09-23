VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{2B323CCC-50E3-11D3-9466-00A0C9700498}#1.0#0"; "yacscom.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YCHT Voice Client Example By: Matthew Robertson"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstChatters 
      Height          =   3375
      Left            =   4920
      TabIndex        =   11
      Top             =   240
      Width           =   1935
   End
   Begin VB.CheckBox chkTalk 
      Caption         =   "Talk"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3700
      Width           =   615
   End
   Begin VB.CommandButton cmdEnable 
      Caption         =   "Enable Voice"
      Height          =   285
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3700
      Width           =   1215
   End
   Begin VB.CommandButton cmdDisconnect 
      Cancel          =   -1  'True
      Caption         =   "Disconnect"
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtIn 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1320
      Width           =   4695
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox txtMsg 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   5775
   End
   Begin VB.Timer tmrPing 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   0
      Top             =   240
   End
   Begin MSWinsockLib.Winsock wskYCHT 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox chkHandFree 
      Caption         =   "Hands Free"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.ComboBox cboServers 
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtRoom 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chatters:"
      Height          =   195
      Index           =   2
      Left            =   5040
      TabIndex        =   17
      Top             =   0
      Width           =   630
   End
   Begin VB.Label lblTalker 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   3720
      Width           =   3735
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Voice Server/Room:"
      Height          =   195
      Index           =   1
      Left            =   2640
      TabIndex        =   15
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Yahoo ID/Pass:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   0
      Width           =   1140
   End
   Begin YACSCOMLibCtl.YAcs YVoice 
      Left            =   120
      OleObjectBlob   =   "frmMain.frx":0000
      Top             =   0
   End
   Begin VB.Label lblStat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By: Matthew Robertson"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   1000
      Width           =   1665
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BadWords As String ' string of cusses yahoo sends to filter
Sub AddText(ByVal Txt As String)
With txtIn
    If Len(.Text) > 10000 Then .Text = Right(.Text, 2000)
    .SelStart = Len(.Text)
    If Not .Text = "" Then .SelText = vbCrLf
    Txt = FilterYahooText(Txt)
    Txt = FilterBadWords(Txt)
    .SelText = Txt
    .SelStart = Len(.Text)
End With
End Sub

Sub ChatList(ByVal Data As String)
Dim Spt() As String, Chatters As String
Data = MidStr(Data, "rmspace", "")
Spt = Split(MidStr(Data, "À€0À€", "À€0"), Chr(1))
For i = 0 To UBound(Spt)
    Spt(i) = Trim(MidStr(Spt(i), "", Chr(2)))
    If Not Spt(i) = "" Then
     ChatListRem Spt(i), False
     lstChatters.AddItem Spt(i)
     Chatters = Chatters & Spt(i) & " "
    End If
Next
AddText " " & Chatters & "joins the room."
End Sub

Sub ChatListRem(Chatter As String, Optional Display As Boolean = True)
With lstChatters
 For i = 0 To .ListCount - 1
  If LCase(.List(i)) = LCase(Chatter) Then
    .RemoveItem i
    If Display = True Then AddText " " & Chatter & " leaves the room."
  End If
 Next
End With
End Sub


Function FilterBadWords(ByVal Txt As String)
On Error GoTo Error
If BadWords = "" Then GoTo Error
Dim Words() As String
Words = Split(BadWords, ",")
For i = 0 To UBound(Words)
    Txt = Replace(Txt, Words(i), String(Len(Words(i)), "*"), , , vbTextCompare)
Next
Error:
FilterBadWords = Txt
End Function

Sub GetVoiceServers()
Dim Data As String, Spt() As String, Serv As String
Data = "http://vc.yahoo.com/"
Data = frmHTTP.OpenURL(Data)
Data = MidStr(Data, "<pre>", "</pre>")
Debug.Print Data
Spt = Split(Data, Chr(10))
For i = 1 To UBound(Spt) - 1
    Serv = Trim(Spt(i))
    cboServers.AddItem MidStr(Serv, "", " ")
Next
cboServers.Text = cboServers.List(0)
End Sub

Sub LoadInfo()
txtID = GetSetting("YCHT", "Login", "ID", "")
txtPW = GetSetting("YCHT", "Login", "PW", "")
txtRoom = GetSetting("YCHT", "Chat", "Room", "Programming:1")
End Sub

Function MidStr(ByVal allStr As String, preStr As String, pstStr As String)
Dim i As Integer
'On Error GoTo Error
i = InStr(1, allStr, preStr, vbTextCompare)
If Not i = 0 Or preStr = "" Then
 i = i + Len(preStr)
 allStr = Mid(allStr, i)
Else
 GoTo Error
End If
i = InStr(1, allStr, pstStr, vbTextCompare)
If Not i = 0 And Not pstStr = "" Then
 i = i - 1
 allStr = Left(allStr, i)
Else
 GoTo Error
End If
MidStr = Trim(allStr)
Exit Function
Error:
MidStr = allStr
End Function

Sub SaveInfo()
If Not YCHT.ID = "" Then SaveSetting "YCHT", "Login", "ID", YCHT.ID
If Not YCHT.PW = "" Then SaveSetting "YCHT", "Login", "PW", YCHT.PW ' u might wonna incode this somehow
If Not YCHT.Room = "" Then SaveSetting "YCHT", "Chat", "Room", YCHT.Room
End Sub

Sub VoiceChat()
Dim VCServ As String
VCServ = cboServers.Text
If VCServ = "" Then VCServ = "v1.vc.scd.yahoo.com"
With YVoice
    .HostName = VCServ
    .appInfo = "mc(6, 0, 0, 0000)&u=" & YCHT.ID & "&ia=us"
    .userName = YCHT.ID
    .loadSound YCHT.ID
    .confKey = YCHT.VCAuth
    .confName = "ch/" & YCHT.Room & "::" & YCHT.RMSpace
    .inputGain = 99
    .outputGain = 99
    .inputAGC = 99
    .inputSource = 99
    .createAndJoinConference
    .joinConference
End With
End Sub

Private Sub cboServers_Change()
YVoice.HostName = cboServers.Text
'VoiceChat
End Sub


Private Sub chkHandFree_Click()
If chkHandFree.Value = 1 Then
    chkTalk.Value = 1
    YVoice.startTransmit
Else
    chkTalk.Value = 0
    YVoice.stopTransmit
End If
End Sub



Private Sub chkTalk_Click()
'YVoice.stopTransmit
End Sub

Private Sub chkTalk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
YVoice.startTransmit
chkTalk.Value = 1
End Sub


Private Sub chkTalk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
YVoice.stopTransmit
chkHandFree.Value = 0
chkTalk.Value = 0
End Sub


Private Sub cmdDisconnect_Click()
wskYCHT.Close
cmdLogin.Caption = "Login"
lblStat = "Disconnected"
cmdEnable.Caption = "Enable Voice"
YVoice.leaveConference
lstChatters.Clear
AddText " Disconnected"
End Sub

Private Sub cmdEnable_Click()
If cmdEnable.Caption = "Enable Voice" Then
    VoiceChat
    cmdEnable.Caption = "Disable Voice"
ElseIf cmdEnable.Caption = "Disable Voice" Then
    YVoice.leaveConference
    cmdEnable.Caption = "Enable Voice"
End If
End Sub

Private Sub cmdLogin_Click()
If cmdLogin.Caption = "Login" Then
    With YCHT
     .ID = txtID
     .PW = txtPW
     .Room = txtRoom
     .Serv = "jcs.chat.dcn.yahoo.com"
    End With
    lblStat = "Connecting..."
    Dim Data As String
    Data = "http://login.yahoo.com/config?login=" & YCHT.ID & "&passwd=" & HexPassword(YCHT.PW)
    Data = frmHTTP.OpenURL(Data)
    If CheckPassword(Data) = True Then
     lblStat = "Logging in..."
     GetCookies Data
     wskYCHT.Connect YCHT.Serv, 8001
     GetVoiceServers
     tmrPing.Enabled = True
    Else
     lblStat = "Faulty Password!"
    End If
ElseIf cmdLogin.Caption = "Join" Then
    YVoice.leaveConference
    lstChatters.Clear
    YCHT.Room = txtRoom
    SendPack Join(YCHT.Room)
End If
End Sub

Function SendPack(Pack As String) As Boolean
On Error GoTo Error
 frmMain.wskYCHT.SendData Pack
 Debug.Print "  " & Pack
 SendPack = True
 Exit Function
Error:
 SendPack = False
End Function










Function FilterYahooText(ByVal Str As String)
On Error GoTo Error
Dim i As Integer, ii As Integer, Llp As Boolean
Llp = True
For lp = 1 To 11 ' better then a loop because this wont ever loop forever
 Llp = False
  i = InStr(Str, "<")
 If Not i = 0 Then
    ii = InStr(i, Str, ">")
    If Not ii = 0 Then
     Str = Left(Str, i - 1) & Right(Str, Len(Str) - ii)
     Llp = True
    End If
 End If
  i = InStr(Str, "[")
 If Not i = 0 Then
    ii = InStr(i, Str, "m")
    If Not ii = 0 Then
     Str = Left(Str, i - 1) & Right(Str, Len(Str) - ii)
     Llp = True
    End If
 End If
    DoEvents
If Llp = False Then Exit For
Next
Error:
FilterYahooText = Str
End Function


Private Sub cmdSend_Click()
Dim Msg As String
Msg = FilterBadWords(txtMsg)
txtMsg = ""
If LCase(Msg) = "/refresh" Then ' refresh list
    lstChatters.Clear
    SendPack Join(YCHT.Room)
ElseIf LCase(Left(Msg, 6)) = "/join " Then ' join room
    txtRoom = Mid(Msg, 7)
    Call cmdLogin_Click ' join
Else
    SendPack ChatSend(Msg, YCHT.Room) ' send msg
End If
End Sub



Private Sub Form_Load()
LoadInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
Hide
SaveInfo
For i = 0 To Forms.Count - 1
    Unload Forms(i)
Next
End
End Sub

Private Sub tmrPing_Timer()
'SendPack Ping
End Sub

Private Sub txtID_Change()
cmdLogin.Default = True
End Sub

Private Sub txtMsg_Change()
cmdSend.Default = True
End Sub

Private Sub txtPW_Change()
cmdLogin.Default = True
End Sub

Private Sub txtRoom_Change()
cmdLogin.Default = True
End Sub

Private Sub wskYCHT_Close()
cmdDisconnect_Click
End Sub

Private Sub wskYCHT_Connect()
SendPack Login(YCHT.ID, YCookie.Y, YCookie.T)
End Sub

Private Sub wskYCHT_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String, sptData() As String, PackType As Integer
wskYCHT.GetData Data
PackType = Asc(Mid(Data, 12, 1))
sptData = Split(Mid(Data, 17), "À€")
Debug.Print PackType & "- " & Mid(Data, 17)
Select Case PackType
 Case 65 ' chat text
    AddText sptData(1) & ": " & sptData(2)
 Case 18 ' chat depart
    ChatListRem sptData(1)
 Case 17 ' chat join
    If sptData(0) = "That room is full.  Try a similar room?" Then ' full room
     AddText "Room is full, trying similar rooms."
     SendPack Join(sptData(1))
    ElseIf Left(sptData(0), 3) = "***" Then ' error msg
     AddText " " & sptData(0)
    ElseIf InStr(Data, "&.vcauth=") Then
     If Not sptData(0) = "Join Failed" Then ' if no error or refreshing list
      If InStr(sptData(0), ":") Then YCHT.Room = sptData(0)
      AddText "Joined " & YCHT.Room & " - " & sptData(1)
      txtRoom = YCHT.Room
     End If
     YCHT.VCAuth = MidStr(Data, "&.vcauth=", "&")
     YCHT.RMSpace = MidStr(Data, "&.rmspace=", "À€")
     VoiceChat ' auto-voice connect
     lstChatters.Clear
    End If
    ChatList Data
 Case 2 ' disconnect
    If InStr(Data, "Logoff successful.") Then Call cmdDisconnect_Click
 Case 1 ' logged in
    cmdLogin.Caption = "Join"
    lblStat = sptData(0) & " logged in"
    AddText "Logged in " & sptData(0)
    BadWords = sptData(1)
End Select
End Sub




Private Sub YVoice_onAudioError(ByVal code As Long, ByVal message As String)
lblTalker = "Voice Error!"
End Sub

Private Sub YVoice_onConferenceNotReady()
lblTalker = "Voice Error!"
End Sub

Private Sub YVoice_onConferenceReady()
lblTalker = "Voice Ready!"
End Sub


Private Sub YVoice_onLocalOffAir()
lblTalker = "Off Air"
chkTalk.Value = 0
End Sub

Private Sub YVoice_onLocalOnAir()
lblTalker = "On Air"
End Sub


Private Sub YVoice_onRemoteSourceOffAir(ByVal sourceId As Long, ByVal sourceName As String)
lblTalker = ""
End Sub

Private Sub YVoice_onRemoteSourceOnAir(ByVal sourceId As Long, ByVal sourceName As String)
lblTalker = sourceName
End Sub


Private Sub YVoice_onSystemConnect()
lblTalker = "Voice Connected!"
chkHandFree.Enabled = True
chkTalk.Enabled = True
cmdEnable.Caption = "Disable Voice"
End Sub


Private Sub YVoice_onSystemConnectFailure(ByVal code As Long, ByVal message As String)
lblTalker = "Voice Error!"
cmdEnable.Caption = "Enable Voice"
End Sub


Private Sub YVoice_onSystemDisconnect()
lblTalker = "Voice Disconencted!"
chkHandFree.Enabled = False
chkTalk.Enabled = False
cmdEnable.Caption = "Enable Voice"
End Sub


