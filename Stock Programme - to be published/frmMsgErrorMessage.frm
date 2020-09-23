VERSION 5.00
Begin VB.Form frmMsgErrorMessage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SCS Messenger"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5040
   Icon            =   "frmMsgErrorMessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   100
      Width           =   4815
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
         Left            =   1800
         MouseIcon       =   "frmMsgErrorMessage.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You can't send messages to your Account itself."
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   3390
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   120
         Picture         =   "frmMsgErrorMessage.frx":015E
         Top             =   120
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMsgErrorMessage"
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
Me.Top = frmSendMessage.Top + (frmSendMessage.Height - Me.Height) / 2
Me.Left = frmSendMessage.Left + (frmSendMessage.Width - Me.Width) / 2
End Sub
