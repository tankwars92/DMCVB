VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "DMconnect"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "DMconnectClient.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtLog 
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   240
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "dmconnect.hoho.ws"
      RemotePort      =   1111
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wmMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const EM_SCROLLCARET = &HB7

Private Sub Form_Resize()
    Dim pad As Integer
    pad = 120

    Dim bottomY As Integer
    bottomY = Me.ScaleHeight - txtInput.Height - pad

    txtInput.Move pad, bottomY, Me.ScaleWidth - cmdSend.Width - 3 * pad
    cmdSend.Move txtInput.Left + txtInput.Width + pad, bottomY

    txtLog.Move pad, pad, Me.ScaleWidth - 2 * pad, txtInput.Top - 2 * pad
End Sub





Private Sub txtLog_GotFocus()
    txtInput.SetFocus
End Sub

Private Sub ScrollToBottom(txt As TextBox)
    txt.SelStart = Len(txt.Text)
    SendMessage txt.hwnd, EM_SCROLLCARET, 0, 0
End Sub

Private Sub Form_Load()
    Winsock1.RemoteHost = "dmconnect.hoho.ws"
    Winsock1.RemotePort = 1112
    Winsock1.Connect
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        cmdSend_Click
    End If
End Sub

Private Sub Winsock1_Connect()
    txtLog.Text = txtLog.Text & "[Connected]"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim data() As String
    Winsock1.GetData msg, vbString
    
    txtLog.Text = txtLog.Text & vbCrLf & msg
    
    ScrollToBottom txtLog
End Sub

Private Sub cmdSend_Click()
    Winsock1.SendData txtInput.Text
    txtInput.Text = ""
End Sub



