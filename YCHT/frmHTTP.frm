VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmHTTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmHTTP"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "http://www.yahoo.com/"
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test OpenURL()"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin MSWinsockLib.Winsock wskHTTP 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHTTP.frx":0000
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmHTTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmHTTP, By: Matthew Robertson

Dim Timeout     As Integer
Dim HTML        As String
Dim HTTP_Server As String
Dim HTTP_Page   As String
Sub CancelURL()
HTML = "Cancelled"
wskHTTP.Close
End Sub


Function HTTP(Page As String, Optional Host As String = "127.0.0.1")
HTTP = "GET /" & Page & " HTTP/1.1" & vbCrLf & _
"Host: " & Host & vbCrLf & _
"User-Agent: Mozilla/5.0 (Windows) frmHTTP" & vbCrLf & _
"Accept: text/html,*/*" & vbCrLf & _
"Accept -Language: en -ca" & vbCrLf & vbCrLf
End Function

Public Function OpenURL(ByVal URL As String, Optional Port As Integer = 80)
If Timeout = 0 Then Timeout = 10
On Error GoTo Error
CancelURL
Dim i As Integer
If LCase(Left(URL, 7)) = "http://" Then URL = Mid(URL, 8, Len(URL) - 7)
i = InStr(URL, "/")
If i < 1 Then
    URL = URL & "/"
    i = Len(URL)
End If
HTTP_Server = Left(URL, i - 1)
HTTP_Page = Right(URL, Len(URL) - i)
HTML = ""
wskHTTP.Connect HTTP_Server, Port
Dim Sec As Long
Sec = Timer + Timeout
Do Until Timer > Sec
    DoEvents
    If wskHTTP.State = 0 Then GoTo Done
Loop
wskHTTP.Close
'HTML = "Time out"
Done:
OpenURL = HTML
Exit Function
Error:
OpenURL = "Error!"
End Function



Private Sub cmdTest_Click()
Dim Test As String
Test = OpenURL(txtURL)
Debug.Print Test
MsgBox Test
End Sub


Private Sub Form_Load()
Timeout = 10 ' time to load a page (sec)
End Sub


Private Sub Form_Unload(Cancel As Integer)
CancelURL
End Sub

Private Sub wskHTTP_Close()
wskHTTP.Close
End Sub

Private Sub wskHTTP_Connect()
wskHTTP.SendData HTTP(HTTP_Page, HTTP_Server)
End Sub

Private Sub wskHTTP_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
wskHTTP.GetData Data
HTML = HTML & Data
End Sub


