VERSION 5.00
Begin VB.Form frmUserSendingMessage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SCS Messenger"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5910
   Icon            =   "frmUserSendingMessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   5655
      TabIndex        =   0
      Top             =   100
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
         MouseIcon       =   "frmUserSendingMessage.frx":000C
         TabIndex        =   1
         Top             =   840
         Width           =   1215
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
         TabIndex        =   5
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   45
      End
      Begin VB.Label lblCurrentSenderReceiver 
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
         Left            =   1725
         TabIndex        =   3
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The User, "
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   705
         Left            =   120
         Picture         =   "frmUserSendingMessage.frx":015E
         Top             =   120
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmUserSendingMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
Beep
Me.Picture = frmStyle.Picture

'Me.Top = frmSendMessage.Top + (frmSendMessage.Height - Me.Height) / 2
'Me.Left = frmSendMessage.Left + (frmSendMessage.Width - Me.Width) / 2

If user_logged_out = True Then
    lblCurrentSenderReceiver = frmSendMessage.cmbLoggedUsers
    Label2.Caption = "has logged out."
Else
lblCurrentSenderReceiver = Sending_User
    Label2.Caption = "is sending a Message now. Please try again shortly."
End If
End Sub
