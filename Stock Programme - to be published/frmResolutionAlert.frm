VERSION 5.00
Begin VB.Form frmResolutionAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resolution Alert"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7590
   Icon            =   "frmResolutionAlert.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.CheckBox chkDontShow 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Do not show this message again."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         MouseIcon       =   "frmResolutionAlert.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1095
         Left            =   960
         TabIndex        =   12
         Top             =   120
         Width           =   6135
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To get a better resolution fit of Stock Controlling System, set the resolution setting to"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   5895
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1024 by 768 pixels or"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "800 by 600 pixels."
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Do you need to set the resolution setting manually?"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   3600
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1095
         Left            =   960
         TabIndex        =   7
         Top             =   120
         Width           =   6135
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Do you need to set the resolution setting manually?"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   3600
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1024 by 768 pixels."
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   1365
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To get a better resolution fit of Stock Controlling System, set the resolution setting to"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   5895
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Your current screen resolution is 800 by 600 pixels."
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   3585
         End
      End
      Begin VB.CommandButton cmdYes 
         Caption         =   "&Yes"
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
         Left            =   4560
         MouseIcon       =   "frmResolutionAlert.frx":015E
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdNo 
         Cancel          =   -1  'True
         Caption         =   "&No"
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
         Left            =   5880
         MouseIcon       =   "frmResolutionAlert.frx":02B0
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   705
         Left            =   120
         Picture         =   "frmResolutionAlert.frx":0402
         Top             =   240
         Width           =   750
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
         TabIndex        =   6
         Top             =   120
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   480
         Width           =   45
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
         TabIndex        =   4
         Top             =   480
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmResolutionAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkDontShow_Click()
On Error Resume Next
reg_obj.RegWrite (Resolution_Alert), "0"
intresolutionAlertenable = 0
End Sub

Private Sub cmdNo_Click()
Unload Me
End Sub

Private Sub cmdYes_Click()
On Error Resume Next
 Shell "Control desk.cpl,,3", vbNormalFocus
 Store_User_Logged_Status_Logout
 User_Log_Out
 End
 End Sub

Private Sub Form_Activate()
On Error Resume Next
cmdYes.SetFocus
End Sub

Private Sub Form_Load()
Me.Picture = frmStyle.Picture
If intresmsg = 0 Then
    Frame2.Visible = True
    Frame1.Visible = False
ElseIf intresmsg = 1 Then
    Frame1.Visible = True
    Frame2.Visible = False
End If
End Sub
