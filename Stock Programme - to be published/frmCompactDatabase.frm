VERSION 5.00
Begin VB.Form frmCompactDatabase 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compact Database..."
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4980
   Icon            =   "frmCompactDatabase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3330
      Left            =   80
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2955
         Left            =   120
         Picture         =   "frmCompactDatabase.frx":000C
         ScaleHeight     =   2955
         ScaleWidth      =   4575
         TabIndex        =   1
         Top             =   250
         Width           =   4575
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "Cancel"
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
            Left            =   2880
            MouseIcon       =   "frmCompactDatabase.frx":1B06
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   2520
            Width           =   1575
         End
         Begin VB.CommandButton cmdCompactDatabase 
            Caption         =   "Compact Database"
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
            Left            =   840
            MouseIcon       =   "frmCompactDatabase.frx":1C58
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   2520
            Width           =   1935
         End
         Begin VB.Image Image5 
            Height          =   315
            Left            =   840
            Picture         =   "frmCompactDatabase.frx":1DAA
            Top             =   1620
            Width           =   345
         End
         Begin VB.Image Image4 
            Height          =   315
            Left            =   840
            Picture         =   "frmCompactDatabase.frx":23D4
            Top             =   1260
            Width           =   345
         End
         Begin VB.Image Image1 
            Height          =   675
            Left            =   0
            Picture         =   "frmCompactDatabase.frx":29FE
            Top             =   2280
            Width           =   750
         End
         Begin VB.Label lblActiveUserLogStatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monitor Active User Status..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   210
            Left            =   1320
            MouseIcon       =   "frmCompactDatabase.frx":44F8
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   1320
            Width           =   2670
         End
         Begin VB.Label lblSendMessagetoUsers 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Send Message to Users..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   210
            Left            =   1320
            MouseIcon       =   "frmCompactDatabase.frx":464A
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   1680
            Width           =   2415
         End
         Begin VB.Image Image3 
            Height          =   735
            Left            =   0
            Picture         =   "frmCompactDatabase.frx":479C
            Top             =   -10
            Width           =   735
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmCompactDatabase.frx":6432
            Height          =   975
            Left            =   840
            TabIndex        =   4
            Top             =   360
            Width           =   3615
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Important Tip:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   3
            Top             =   0
            Width           =   1335
         End
         Begin VB.Line Line1 
            X1              =   840
            X2              =   4440
            Y1              =   2280
            Y2              =   2280
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "frmCompactDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF8080
lblActiveUserLogStatus.ForeColor = &HFF8080

End Sub

Private Sub cmdCompactDatabase_Click()
inttask = 2
Compact_Database
End Sub

Private Sub cmdCompactDatabase_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF8080
lblActiveUserLogStatus.ForeColor = &HFF8080
End Sub

Private Sub Form_Load()
Me.Picture = frmStyle.Picture
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF8080
lblActiveUserLogStatus.ForeColor = &HFF8080
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF80FF
lblActiveUserLogStatus.ForeColor = &HFF80FF
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF8080
lblActiveUserLogStatus.ForeColor = &HFF8080
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF8080
lblActiveUserLogStatus.ForeColor = &HFF8080
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF8080
lblActiveUserLogStatus.ForeColor = &HFF8080
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF8080
lblActiveUserLogStatus.ForeColor = &HFF8080

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF8080
lblActiveUserLogStatus.ForeColor = &HFF8080
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF8080
lblActiveUserLogStatus.ForeColor = &HFF8080
End Sub

Private Sub lblActiveUserLogStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblActiveUserLogStatus.Left = lblActiveUserLogStatus.Left + 20
lblActiveUserLogStatus.Top = lblActiveUserLogStatus.Top + 20
End Sub

Private Sub lblActiveUserLogStatus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblActiveUserLogStatus.ForeColor = &HFF80FF
lblSendMessagetoUsers.ForeColor = &HFF8080
End Sub

Private Sub lblActiveUserLogStatus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblActiveUserLogStatus.Left = lblActiveUserLogStatus.Left - 20
lblActiveUserLogStatus.Top = lblActiveUserLogStatus.Top - 20
frmLoggedUserStatus.Show 1
End Sub
Private Sub lblSendMessagetoUsers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.Left = lblSendMessagetoUsers.Left + 20
lblSendMessagetoUsers.Top = lblSendMessagetoUsers.Top + 20
End Sub

Private Sub lblSendMessagetoUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF80FF
lblActiveUserLogStatus.ForeColor = &HFF8080
End Sub

Private Sub lblSendMessagetoUsers_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.Left = lblSendMessagetoUsers.Left - 20
lblSendMessagetoUsers.Top = lblSendMessagetoUsers.Top - 20
frmSendMessage.Show 1
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSendMessagetoUsers.ForeColor = &HFF8080
lblActiveUserLogStatus.ForeColor = &HFF8080
End Sub
