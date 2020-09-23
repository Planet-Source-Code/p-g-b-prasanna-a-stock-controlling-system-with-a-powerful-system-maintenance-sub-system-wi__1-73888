VERSION 5.00
Begin VB.Form frmSentMsgMessage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Sent"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5910
   Icon            =   "frmSentMsgMessage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
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
         MouseIcon       =   "frmSentMsgMessage.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   735
         Left            =   120
         Picture         =   "frmSentMsgMessage.frx":015E
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message to "
         Height          =   195
         Left            =   960
         TabIndex        =   4
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblReceiver 
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
         Left            =   1845
         TabIndex        =   3
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " sent succeeded."
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmSentMsgMessage"
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

If intsending_all = 1 Then
    lblReceiver = "All Logged Users"
Else
lblReceiver = frmSendMessage.cmbLoggedUsers
End If

End Sub

