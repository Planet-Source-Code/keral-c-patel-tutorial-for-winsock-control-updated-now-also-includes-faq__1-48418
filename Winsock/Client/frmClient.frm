VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClient 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmClient.frx":0000
   ScaleHeight     =   3915
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   225
      Top             =   3780
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4980
      Top             =   780
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstMessages 
      Height          =   2400
      Left            =   240
      TabIndex        =   4
      Top             =   630
      Width           =   4215
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   3195
      Width           =   495
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   75
      Width           =   1455
   End
   Begin VB.TextBox txtSend 
      Height          =   330
      Left            =   225
      TabIndex        =   2
      Top             =   3225
      Width           =   3630
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   135
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   3420
      Left            =   4620
      TabIndex        =   7
      Top             =   405
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   6033
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Client-Side Example for Winsock Tutorial By: Keral."
      Height          =   195
      Left            =   735
      TabIndex        =   6
      Top             =   3615
      Width           =   3660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   195
      Left            =   4605
      TabIndex        =   5
      Top             =   165
      Width           =   210
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'If you liked this then Please vote for me.
'©2003 Keral.C.Patel.
'For making the Form Movable
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'This API's are for changing a progressbar's forecolor
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
'For launching the explorer
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
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

Private Sub cmdConnect_Click()

    On Error Resume Next
    'This will connect to the Computer Specified By IP and the Port
    Winsock1.Connect txtIP.Text, "1412" 'Just remember this Port Number Should be Same on which our Server is Listening

End Sub

Private Sub cmdSend_Click()

    On Error Resume Next
    'The Following data will be sent to the Server Side
    Winsock1.SendData "Client:-    " & txtSend.Text
    lstMessages.AddItem "Client:-    " & txtSend.Text
    'Clear the TextBox & set the Focus
    txtSend.Text = ""
    txtSend.SetFocus

End Sub

Private Sub Form_Load()

    'Now for Making our progressbar's Fore color to red
    PBcolor ProgressBar1, vbWhite, vbBlue

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'For making the Form Movable
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim msgret As Byte
msgret = MsgBox("Do you want to vote for me?", vbYesNo, "Winsock-Tutorial")
If msgret = 6 Then ShellExecute Me.hwnd, vbNullString, "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=48418&lngWId=1", vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL
End Sub

Private Sub Label1_Click()

    On Error Resume Next
    'Letting server know that client has Disconnected.
    Winsock1.SendData "Client is Disconnected!"
    DoEvents
    Unload Me

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    On Error Resume Next
    Dim str As String
    Winsock1.GetData str
    lstMessages.AddItem str
    'To start the animation
    Timer1.Enabled = True
    ProgressBar1.Value = 0

End Sub

Private Sub Timer1_Timer()

    If ProgressBar1.Value < 100 Then

        ProgressBar1.Value = ProgressBar1.Value + 10

    Else

        Timer1.Enabled = False

    End If

End Sub

