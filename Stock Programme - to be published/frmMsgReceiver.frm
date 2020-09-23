VERSION 5.00
Begin VB.Form frmMsgReceiver 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Message Receiver"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   240
      Picture         =   "frmMsgReceiver.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   5055
      TabIndex        =   17
      Top             =   6600
      Width           =   5050
      Begin VB.Image imgProgress 
         Height          =   195
         Left            =   0
         Picture         =   "frmMsgReceiver.frx":2F32
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5055
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -120
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6060
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5895
         Left            =   40
         ScaleHeight     =   5895
         ScaleWidth      =   5175
         TabIndex        =   4
         Top             =   120
         Width           =   5175
         Begin VB.CommandButton cmdOk 
            Cancel          =   -1  'True
            Caption         =   "OK"
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
            Left            =   3855
            MouseIcon       =   "frmMsgReceiver.frx":432C
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   5400
            Width           =   1215
         End
         Begin VB.CommandButton cmdReplytoSender 
            Caption         =   "Reply to Sender"
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
            Left            =   1680
            MouseIcon       =   "frmMsgReceiver.frx":447E
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   5400
            Width           =   2055
         End
         Begin VB.TextBox txtIncomming_Msg_Body 
            BackColor       =   &H00FBF4F4&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2565
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   1
            Top             =   2420
            Width           =   4935
         End
         Begin VB.Line Line6 
            X1              =   840
            X2              =   1560
            Y1              =   5760
            Y2              =   5760
         End
         Begin VB.Line Line5 
            X1              =   840
            X2              =   1560
            Y1              =   5640
            Y2              =   5640
         End
         Begin VB.Line Line4 
            X1              =   840
            X2              =   1560
            Y1              =   5520
            Y2              =   5520
         End
         Begin VB.Line Line3 
            X1              =   840
            X2              =   1560
            Y1              =   5400
            Y2              =   5400
         End
         Begin VB.Line Line2 
            X1              =   840
            X2              =   5040
            Y1              =   5280
            Y2              =   5280
         End
         Begin VB.Line Line1 
            X1              =   840
            X2              =   5040
            Y1              =   5160
            Y2              =   5160
         End
         Begin VB.Label lblDateSent 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1680
            TabIndex        =   16
            Top             =   1560
            Width           =   3375
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Sent:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label lblTimeSent 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1680
            TabIndex        =   14
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time Sent:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Image Image1 
            Height          =   765
            Left            =   120
            Picture         =   "frmMsgReceiver.frx":45D0
            Top             =   5100
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Message Type:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1440
         End
         Begin VB.Label lblMsgType 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1680
            TabIndex        =   8
            Top             =   600
            Width           =   3375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Message:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblMsgFrom 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1680
            TabIndex        =   6
            Top             =   120
            Width           =   3375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Message From:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   1470
         End
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Seconds."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4560
      TabIndex        =   12
      Top             =   6300
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Remaining"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3480
      TabIndex        =   11
      Top             =   6300
      Width           =   825
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4410
      TabIndex        =   10
      Top             =   6300
      Width           =   90
   End
End
Attribute VB_Name = "frmMsgReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
If blnsystem_log_out_notification = True Then
    Store_User_Logged_Status_Logout
    intsystemlogstatus = 3
    User_Log_Out
    'Unload Me
    'Unload frmMain
    End
ElseIf blnuser_log_off_notification = True Then
    Unload Me
    'For Each f In Forms
        'If f.Name <> "frmLogin" Then
            'On Error Resume Next
            'Unload f
        'End If
        'Next
    'Unload Me
    'Unload frmMain
    intsystemlogstatus = 4
    Form_Unload_Pro
    frmLogin.Show
Else
    Unload Me
End If
blnsystem_log_out_notification = False
blnuser_log_off_notification = False
End Sub

Private Sub cmdReplytoSender_Click()
Unload Me
CHECK_USER_SENDING_MSG
End Sub
Private Sub Form_Load()
Me.Picture = frmStyle.Picture
If blnsystem_log_out_notification = True Or blnuser_log_off_notification = True Then
    'blnsystem_log_out_notification = False
    timremain = 30
    lblTime.Caption = timremain
    Timer1.Enabled = True
    cmdReplytoSender.Enabled = False
End If
If user_send_msg_privilege = 0 Then
        cmdReplytoSender.Enabled = False
End If
'Label7.Width = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Cancel = 1
End Sub

Private Sub Timer1_Timer()
Dim tcwp, tcwimg  As Integer
timremain = timremain - 1
lblTime.Caption = timremain
tcwp = timremain / 30 * 100
tcwimg = tcwp / 100 * 5055
'c = b + b
imgProgress.Width = tcwimg
'c = b / 4815 * 100
c = timremain / 30 * 100
'Label8.Caption = c & " %"
'Label7.Refresh
'lblTime.Refresh
If timremain = 0 Then
Call cmdOk_Click
End If
End Sub
