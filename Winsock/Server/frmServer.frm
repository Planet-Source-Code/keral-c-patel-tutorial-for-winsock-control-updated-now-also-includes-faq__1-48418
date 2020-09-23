VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmServer.frx":0000
   ScaleHeight     =   3405
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   30
      Top             =   3300
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   2940
      Left            =   4455
      TabIndex        =   5
      Top             =   360
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   5186
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtSend 
      Height          =   330
      Left            =   150
      TabIndex        =   0
      Top             =   2745
      Width           =   3630
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3885
      TabIndex        =   1
      Top             =   2715
      Width           =   495
   End
   Begin VB.ListBox lstMessages 
      Height          =   2400
      Left            =   165
      TabIndex        =   2
      Top             =   150
      Width           =   4215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4620
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Server-Side Example for Winsock Tutorial By: Keral."
      Height          =   195
      Left            =   555
      TabIndex        =   4
      Top             =   3135
      Width           =   3660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   150
      Width           =   210
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a Small Example but with this You can do many Big things.
'If you liked this then Please vote for me.
'©2003 Keral.C.Patel.
'This API's are for making a form Movable
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'This API's are for changing a progressbar's forecolor
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)

'Progressbar
Private Const WM_USER = &H400
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)

Private Sub PBcolor(PB As ProgressBar, Backcolor As Long, Forecolor As Long)

    'Send a message, which window?, what type of message, message value
    SendMessage PB.hwnd, CCM_SETBKCOLOR, 0, ByVal Backcolor
    SendMessage PB.hwnd, PBM_SETBARCOLOR, 0, ByVal Forecolor

End Sub

Private Sub cmdSend_Click()

    On Error Resume Next
    'This data will be sent to the Client
    Winsock1.SendData "Server:-    " & txtSend.Text
    lstMessages.AddItem "Server:-    " & txtSend.Text
    txtSend.Text = ""
    txtSend.SetFocus

End Sub

Private Sub Form_Load()

    On Error Resume Next
    'If one Copy of Our Application is already running then don't load a new one

    If Not App.PrevInstance = True Then

        Winsock1.LocalPort = 1412 'This can be any Valid Port Number
        'Wait for Clients to Connect with Your Server.
        Winsock1.Listen

    End If

    'Now for Making our progressbar's Fore color to red
    PBcolor ProgressBar1, vbWhite, vbRed

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'for making a form Movable
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

End Sub

Private Sub Label1_Click()

    On Error Resume Next 'So that it will not raise an error after sending the data to the server
    'which is already disconnected

    Winsock1.SendData "Server is Disconnected!"
    'Here DoEvents gives time to perform the winsock operation before unloading it from memory
    DoEvents
    'Now Unload it
    Unload Me

End Sub

Private Sub Timer1_Timer()

    If ProgressBar1.Value < 100 Then

        ProgressBar1.Value = ProgressBar1.Value + 10

    Else

        Timer1.Enabled = False

    End If

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

    On Error Resume Next
    'First Check if the Winsock Control is Connected or not
    'If connected then Close it

    If Winsock1.State <> sckClosed Then Winsock1.Close

    'Now accept the Request
    Winsock1.Accept requestID

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    On Error Resume Next
    Dim str As String
    'Now we will store data that has came into this string
    Winsock1.GetData str
    lstMessages.AddItem str

    'Now for progressbar animation
    ProgressBar1.Value = 0
    Timer1.Enabled = True

End Sub

