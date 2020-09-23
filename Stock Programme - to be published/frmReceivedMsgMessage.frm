VERSION 5.00
Begin VB.Form frmReceivedMsgMessage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SCS Messenger"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5910
   Icon            =   "frmReceivedMsgMessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1620
      Left            =   120
      ScaleHeight     =   1620
      ScaleWidth      =   5655
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdOk 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2235
         MouseIcon       =   "frmReceivedMsgMessage.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5160
         Top             =   0
      End
      Begin VB.Image imgProgress 
         Height          =   195
         Left            =   300
         Picture         =   "frmReceivedMsgMessage.frx":015E
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   5055
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   120
         Picture         =   "frmReceivedMsgMessage.frx":1558
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You have received a Message From"
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   2565
      End
      Begin VB.Label lblSender 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   4
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message Type: "
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   480
         Width           =   1140
      End
      Begin VB.Label lblMsgType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmReceivedMsgMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Timer1.Enabled = False
Unload Me
End Sub

Private Sub Form_Load()
Beep
Me.Picture = frmStyle.Picture
lblSender = sender
lblMsgType = Msg_Type
timeformessage = 30
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim tcwp, tcwimg  As Integer
timeformessage = timeformessage - 1
'lblTime.Caption = timremain
tcwp = timeformessage / 30 * 100
tcwimg = tcwp / 100 * 5055
imgProgress.Width = tcwimg
If timeformessage = 0 Then
Call cmdOk_Click
End If
End Sub
