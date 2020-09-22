VERSION 5.00
Object = "{86A5759D-9DDD-11D3-BF2F-00C0F025D341}#1.0#0"; "SYSTRAY.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Safe Proxy"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3405
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSckOut 
      Alignment       =   2  'Center
      BackColor       =   &H80000011&
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtSckIn 
      Alignment       =   2  'Center
      BackColor       =   &H80000011&
      Height          =   285
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   360
      Width           =   735
   End
   Begin VB.Timer tmrKillSock 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   500
      Left            =   1560
      Top             =   840
   End
   Begin MSWinsockLib.Winsock sckIn 
      Index           =   0
      Left            =   1920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   8180
   End
   Begin MSWinsockLib.Winsock sckOut 
      Index           =   0
      Left            =   2160
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin SystemTrayIcon.SysIcon sysTray 
      Left            =   2400
      Top             =   960
      _ExtentX        =   1720
      _ExtentY        =   1296
      NormalPicture   =   "frmMain.frx":030A
      AnimPicture     =   "frmMain.frx":0624
      IconText        =   ""
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "by David Fiala"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   1410
   End
   Begin VB.Label lblAbout 
      AutoSize        =   -1  'True
      Caption         =   "Safe Proxy"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   0
      Width           =   1065
   End
   Begin VB.Label lblSockOut 
      Caption         =   "SockOut:"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblSocks 
      Caption         =   "SockIn:"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "End"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Safe Proxy by David Fiala  - Released to planet-source-code on: May 29, 2001
'djf1010@aol.com
'This is one of my older programs I built and I probally wouldn't release it to the public the way it is
'but I don't plan on fixing it any time soon. Enjoy...

Option Explicit

Private Sub cmdEnd_Click()
    Dim intExit As String
    intExit = MsgBox("Do you really want to close Safe Proxy?", vbYesNo + vbSystemModal + vbDefaultButton2 + vbCritical, "Sure?")
    If intExit = vbYes Then Unload Me
End Sub

Private Sub cmdHide_Click()
    frmMain.Hide 'Hide it
End Sub

Private Sub cmdStart_Click()
    Load sckIn(1) 'load socket for listening
    sckIn(1).LocalPort = "8180" 'set port
    sckIn(1).Listen 'make it listen
    cmdStart.Enabled = False 'diable the command button
    cmdStart.Caption = "Started" 'change the caption to started
End Sub

Private Sub Form_Load()
    sysTray.ShowNormalIcon 'show a icon on systray
End Sub

Private Sub mnuEnd_Click()
    Dim intExit As String
    intExit = MsgBox("Do you really want to close Safe Proxy?", vbYesNo + vbSystemModal + vbDefaultButton2 + vbCritical, "Sure?")
    If intExit = vbYes Then Unload Me
End Sub

Private Sub mnuShow_Click()
    frmMain.Show
End Sub

Private Sub sckIn_Close(Index As Integer)
    On Error Resume Next 'don't let it crash
    txtSckIn.Text = txtSckIn.Text - 1
    Unload sckIn(Index) 'unload the socket so its available later
    Unload sckOut(Index) 'unload the socket so its available later
End Sub

Public Sub sckIn_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Load sckIn(requestID) 'load the socket for this session
    Load sckOut(requestID) 'load the socket for this session
    sckIn(requestID).Tag = "" 'clear tags just incase of previous session
    sckOut(requestID).Tag = "" 'clear tags just incase of previous session
    sckIn(requestID).Accept requestID 'accept the clients connection
    txtSckIn.Text = txtSckIn.Text + 1
End Sub

Private Sub sckIn_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String 'declare variable
    sckIn(Index).GetData strData 'get the incoming data
    If Mid(strData, 1, 4) = "POST" Then
        sckIn(Index).Tag = sckIn(Index).Tag & strData 'add it to .tag
        sckIn(Index).Tag = Replace(sckIn(Index).Tag, "Proxy-", "") 'replace Proxy- to nothing
        sckIn(Index).Tag = Replace(sckIn(Index).Tag, "Keep-Alive", "Close") 'replace Keep-Alive to Close
        sckOut(Index).RemoteHost = GetUrlServerPost(strData) 'set the remote host for the out socket
        sckIn(Index).Tag = NewPost(sckIn(Index).Tag) 'get rid of the http://server in the get command
        sckOut(Index).Connect 'make the out socket connect
    ElseIf Mid(strData, 1, 3) = "GET" Then
        sckIn(Index).Tag = sckIn(Index).Tag & strData 'add it to .tag
        sckIn(Index).Tag = Replace(sckIn(Index).Tag, "Proxy-", "") 'replace Proxy- to nothing
        sckIn(Index).Tag = Replace(sckIn(Index).Tag, "Keep-Alive", "Close") 'replace Keep-Alive to Close
        sckOut(Index).RemoteHost = GetUrlServer(strData) 'set the remote host for the out socket
        sckIn(Index).Tag = NewGet(sckIn(Index).Tag) 'get rid of the http://server in the get command
        sckOut(Index).Connect 'make the out socket connect
    End If 'If Mid(strData, 1, 4) = "POST" Then
End Sub

Private Sub sckIn_SendComplete(Index As Integer)
    On Error Resume Next 'as if this would never crash... yeah right
    txtSckIn.Text = txtSckIn.Text - 1
    Unload sckIn(Index) 'unload the socket so its available later
    Unload sckOut(Index) 'unload the socket so its available later
End Sub

Private Sub sckOut_Close(Index As Integer)
    On Error Resume Next 'make it so it don't crash
    txtSckOut.Text = txtSckOut.Text - 1
    sckOut(Index).Close 'close socket
    sckOut(Index).Tag = Replace(sckOut(Index).Tag, "Keep-Alive", "Close") 'change keep-alive to close
    sckOut(Index).Tag = Blocks(sckOut(Index).Tag) 'block out the restircted stuff (eg: type=file)
    sckIn(Index).SendData sckOut(Index).Tag 'send the data out
    Unload sckOut(Index) 'unload the socket so its available
End Sub

Private Sub sckOut_Connect(Index As Integer)
    txtSckOut.Text = txtSckOut.Text + 1
    sckOut(Index).SendData sckIn(Index).Tag 'send some data
End Sub

Private Sub sckOut_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String 'declare a variable
    sckOut(Index).GetData strData 'retrieve the data
    sckOut(Index).Tag = sckOut(Index).Tag & strData 'put the data on .tag
End Sub

Private Sub sysTray_IconLeftUp()
    PopupMenu mnuPopup
End Sub

Private Sub sysTray_IconRightUp()
    PopupMenu mnuPopup
End Sub

Private Sub tmrKillSock_Timer(Index As Integer)
    On Error Resume Next 'Just in case of error.
    Unload sckIn(Index) 'Unload the incoming socket
    Unload sckOut(Index) 'Unload outgoing the socket
    Unload tmrKillSock(Index) 'Unload this timer
End Sub

