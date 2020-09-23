VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu - Stock Controlling System"
   ClientHeight    =   10740
   ClientLeft      =   -1185
   ClientTop       =   885
   ClientWidth     =   15270
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   11487.91
   ScaleMode       =   0  'User
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9480
      Top             =   120
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   4560
      TabIndex        =   22
      Top             =   0
      Width           =   10695
      Begin VB.Timer Timer4 
         Interval        =   1000
         Left            =   3480
         Top             =   120
      End
      Begin VB.Timer Timer3 
         Interval        =   3000
         Left            =   3960
         Top             =   120
      End
      Begin VB.Timer Timer2 
         Interval        =   3000
         Left            =   4440
         Top             =   120
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         Picture         =   "frmMain.frx":617A
         ScaleHeight     =   735
         ScaleWidth      =   10335
         TabIndex        =   23
         Top             =   240
         Width           =   10335
         Begin VB.CommandButton cmdLogOffUser 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Log Off User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6480
            MouseIcon       =   "frmMain.frx":17CDC
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   0
            Width           =   1455
         End
         Begin VB.CommandButton cmdExit 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8760
            MouseIcon       =   "frmMain.frx":17E2E
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   0
            Width           =   1455
         End
         Begin VB.Image Image13 
            Height          =   780
            Left            =   5640
            Picture         =   "frmMain.frx":17F80
            Top             =   0
            Width           =   660
         End
         Begin VB.Image Image10 
            Height          =   705
            Left            =   8040
            Picture         =   "frmMain.frx":19A92
            Top             =   0
            Width           =   585
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9375
      Left            =   4560
      TabIndex        =   21
      Top             =   1080
      Width           =   10695
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   10028
         Left            =   0
         ScaleHeight     =   10035
         ScaleWidth      =   10755
         TabIndex        =   24
         Top             =   0
         Width           =   10758
         Begin VB.Image Image15 
            Height          =   630
            Left            =   7080
            Picture         =   "frmMain.frx":1B0DC
            Top             =   360
            Width           =   3240
         End
         Begin VB.Image Image9 
            Height          =   4335
            Left            =   1080
            Picture         =   "frmMain.frx":21B6E
            Top             =   2360
            Width           =   8760
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   4335
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   9855
         Left            =   120
         ScaleHeight     =   9855
         ScaleWidth      =   3975
         TabIndex        =   12
         Top             =   240
         Width           =   3975
         Begin VB.CommandButton cmdBookIssue 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Books Issue"
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
            Height          =   615
            Left            =   1800
            MouseIcon       =   "frmMain.frx":9D588
            MousePointer    =   99  'Custom
            TabIndex        =   1
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton cmdBooksReceipt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Books Receipt"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1800
            MouseIcon       =   "frmMain.frx":9D6DA
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   1440
            Width           =   1455
         End
         Begin VB.CommandButton cmdCurStock 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Current Stock"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1800
            MouseIcon       =   "frmMain.frx":9D82C
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   3840
            Width           =   1455
         End
         Begin VB.CommandButton cmdBooksDetails 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Books Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1800
            MouseIcon       =   "frmMain.frx":9D97E
            MousePointer    =   99  'Custom
            TabIndex        =   5
            Top             =   4920
            Width           =   1455
         End
         Begin VB.CommandButton cmdReports 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reports"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1800
            MouseIcon       =   "frmMain.frx":9DAD0
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   2640
            Width           =   1455
         End
         Begin VB.CommandButton cmdMasterInfo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Master Info"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1800
            MouseIcon       =   "frmMain.frx":9DC22
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   6120
            Width           =   1455
         End
         Begin VB.CommandButton cmdPassword 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1800
            MouseIcon       =   "frmMain.frx":9DD74
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   7320
            Width           =   1455
         End
         Begin VB.CommandButton cmdSendMessage 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Send Message"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1800
            MouseIcon       =   "frmMain.frx":9DEC6
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   8640
            Width           =   1455
         End
         Begin VB.Image Image11 
            Height          =   795
            Left            =   240
            Picture         =   "frmMain.frx":9E018
            Top             =   8640
            Width           =   735
         End
         Begin VB.Image Image12 
            Height          =   705
            Left            =   240
            Picture         =   "frmMain.frx":9FEFE
            Top             =   300
            Width           =   720
         End
         Begin VB.Image Image7 
            Height          =   675
            Left            =   240
            Picture         =   "frmMain.frx":A19B0
            Top             =   7320
            Width           =   585
         End
         Begin VB.Image Image6 
            Height          =   645
            Left            =   240
            Picture         =   "frmMain.frx":A2F0A
            Top             =   6120
            Width           =   705
         End
         Begin VB.Image Image5 
            Height          =   705
            Left            =   240
            Picture         =   "frmMain.frx":A477C
            Top             =   4920
            Width           =   600
         End
         Begin VB.Image Image4 
            Height          =   660
            Left            =   240
            Picture         =   "frmMain.frx":A5DC6
            Top             =   3840
            Width           =   630
         End
         Begin VB.Image Image3 
            Height          =   675
            Left            =   120
            Picture         =   "frmMain.frx":A7408
            Top             =   2640
            Width           =   690
         End
         Begin VB.Image Image2 
            Height          =   525
            Left            =   240
            Picture         =   "frmMain.frx":A8CE6
            Top             =   1560
            Width           =   720
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Send Message"
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
            Left            =   120
            TabIndex        =   20
            Top             =   8160
            Width           =   1260
         End
         Begin VB.Line Line8 
            X1              =   1560
            X2              =   3700
            Y1              =   8280
            Y2              =   8280
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password Settings"
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
            Left            =   120
            TabIndex        =   19
            Top             =   6960
            Width           =   1575
         End
         Begin VB.Line Line7 
            X1              =   1920
            X2              =   3720
            Y1              =   7080
            Y2              =   7080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Master Information"
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
            Left            =   120
            TabIndex        =   18
            Top             =   5760
            Width           =   1590
         End
         Begin VB.Line Line6 
            X1              =   1920
            X2              =   3720
            Y1              =   5880
            Y2              =   5880
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Books Details"
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
            Left            =   120
            TabIndex        =   17
            Top             =   4560
            Width           =   1185
         End
         Begin VB.Line Line5 
            X1              =   1560
            X2              =   3720
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Stock"
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
            Left            =   120
            TabIndex        =   16
            Top             =   3480
            Width           =   1185
         End
         Begin VB.Line Line4 
            X1              =   1440
            X2              =   3720
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reports"
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
            Left            =   120
            TabIndex        =   15
            Top             =   2280
            Width           =   675
         End
         Begin VB.Line Line3 
            X1              =   960
            X2              =   3720
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Books Receipt"
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
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   1260
         End
         Begin VB.Line Line2 
            X1              =   1560
            X2              =   3720
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Books Issue"
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
            Left            =   120
            TabIndex        =   13
            Top             =   0
            Width           =   1050
         End
         Begin VB.Line Line1 
            X1              =   1440
            X2              =   3720
            Y1              =   120
            Y2              =   120
         End
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   10485
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuDataEntry 
      Caption         =   "Data &Entry"
      Begin VB.Menu mnuBooksIssue 
         Caption         =   "Books &Issue"
      End
      Begin VB.Menu mnuBooksReceipt 
         Caption         =   "Books &Receipt"
      End
      Begin VB.Menu mnul1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCurrentStock 
         Caption         =   "Current &Stock"
      End
      Begin VB.Menu mnul18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBooksDetails 
         Caption         =   "Books &Details"
      End
      Begin VB.Menu mnuMasterInformation 
         Caption         =   "Master &Information"
      End
      Begin VB.Menu mnul2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBooksnotinuse 
         Caption         =   "Books Not In &Use"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportingStockControllingSystem 
         Caption         =   "Reporting for &Stock Controlling System"
      End
      Begin VB.Menu mnuBooksnotinusereport 
         Caption         =   "Books Not In Use Report"
      End
      Begin VB.Menu mnul17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserStatusReports 
         Caption         =   "User Log Status Reports"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "S&ettings"
      Begin VB.Menu mnuUserPassword 
         Caption         =   "User/Password/Privileges Settings..."
      End
      Begin VB.Menu mnul3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuThemeSettings 
         Caption         =   "Theme &Settings..."
      End
      Begin VB.Menu mnul19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPerformanceoptimizer 
         Caption         =   "Optimize Performance for Common Tasks..."
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Send &Message..."
      End
      Begin VB.Menu mnul4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBypassmsgsender 
         Caption         =   "Bypass Message &Sender  - for Administrators"
      End
      Begin VB.Menu mnul12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebugHandler 
         Caption         =   "Debug Handler..."
         Begin VB.Menu mnuClearMessageSender 
            Caption         =   "Clear Message S&ender - for Administrators"
         End
      End
      Begin VB.Menu mnul5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteData 
         Caption         =   "&De&lete Database Data..."
      End
      Begin VB.Menu mnuClearDatabaseLocation 
         Caption         =   "Clear Database &Location..."
      End
      Begin VB.Menu mnul10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackupDatabase 
         Caption         =   "&Backup/Restore Database..."
      End
      Begin VB.Menu mnuCompactDatabase 
         Caption         =   "&Compact Database..."
      End
      Begin VB.Menu mnul7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActiveUserStatus 
         Caption         =   "Active User &Status..."
      End
      Begin VB.Menu mnul11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTasks 
         Caption         =   "T&asks"
         Begin VB.Menu mnuCalculator 
            Caption         =   "Calc&ulator"
         End
         Begin VB.Menu mnuNotepad 
            Caption         =   "N&otepad"
         End
      End
      Begin VB.Menu mnul9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRunSCSatSystemStatup 
         Caption         =   "Run &Stock Controlling System at System Startup"
      End
      Begin VB.Menu mnul16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnableResolutionAlert 
         Caption         =   "Enable Resolution Alert..."
      End
      Begin VB.Menu mnuChangeresolution 
         Caption         =   "Change Resolution..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnul8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutSystem 
         Caption         =   "About S&ystem"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
      Begin VB.Menu mnuLogOffCurrentUser 
         Caption         =   "Log &Off Current User"
      End
      Begin VB.Menu mnul15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExittoWindows 
         Caption         =   "Exit to &Windows"
      End
   End
   Begin VB.Menu mnuPopupActiveUser 
      Caption         =   "ActiveUser"
      Visible         =   0   'False
      Begin VB.Menu mnuSendaMessage 
         Caption         =   "Send a &Message"
      End
      Begin VB.Menu mnul13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInstantSystemLogOut 
         Caption         =   "User Instant System &Log Out (Force)"
      End
      Begin VB.Menu mnuInstantUserLogOff 
         Caption         =   "User Instant Log O&ff (Force)"
      End
      Begin VB.Menu mnul14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearActiveUser 
         Caption         =   "Clear Active &User - for Administrators"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const Start_Up = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run\SCS"
Dim blnbackupprompt As Boolean
Private Sub cmdBookIssue_Click()
On Error Resume Next
frmBookIssue.Show 1
End Sub

Private Sub cmdBooksDetails_Click()
frmBookdetails.Show 1
End Sub

Private Sub cmdBooksReceipt_Click()
frmBookReceipt.Show 1
End Sub

Private Sub cmdCurStock_Click()
frmCurrentStock.Show 1
End Sub

Private Sub cmdExit_Click()
Unload frmStyle
Store_User_Logged_Status_Logout
intsystemlogstatus = 1
Unload Me
End Sub

Private Sub cmdLogOffUser_Click()
If MsgBox("Are you sure you need to Log Off From " & User & " ?", vbQuestion + vbYesNo) = vbYes Then
intsystemlogstatus = 2
Unload Me
open_status = False
frmLogin.Show
End If
End Sub

Private Sub cmdMasterInfo_Click()
frmMasterInformation.Show 1
End Sub

Private Sub cmdPassword_Click()
frmPassword.Show 1
End Sub

Private Sub cmdRecordsDelete_Click()
'frmDeleteAllRecords.Show 1
End Sub

Private Sub cmdReports_Click()
frmReports.Show 1
End Sub

Public Sub cmdSendMessage_Click()
If Check_For_Privilege(5) = True Then Exit Sub
CHECK_USER_SENDING_MSG
'frmSendMessage.Show 1
End Sub


Private Sub Form_Activate()
If blnbackupprompt = True Then
    blnautobackup = False
End If
If blnautobackup = True Then
    blnbackupprompt = False
End If
If blnbackupprompt = True Then
    If blnautobackup = False Then
        blnbackupprompt = False
        frmBackupdatabase.Show 1
    Exit Sub
    End If
End If
If blnautobackup = True Then
    If blnbackupprompt = False Then
        'blnautobackup = False
        frmBackupdatabase.Show 1
    End If
End If
End Sub

Private Sub Form_Load()
frmMain.Height = Screen.Height
frmMain.Width = Screen.Width
Me.Picture = frmStyle.Picture
stbMain.Panels(1) = Format(Now, "dd/mm/yyyy")
stbMain.Panels(2) = Time
stbMain.Panels(5).Text = "Login As: " & User & " " & Time
Me.Caption = "Main Menu - Stock Controlling System                       --- " & "Login As: " & User & " " & Time & " ---"
open_status = True
sender_bypass = False
sender_clear = False
If intaccount_type = 0 Then
    mnuBypassmsgsender.Enabled = False
    mnuClearMessageSender.Enabled = False
    mnuDeleteData.Enabled = False
    mnuBackupDatabase.Enabled = False
    mnuDebugHandler.Enabled = False
    mnuUserStatusReports.Enabled = False
    mnuCompactDatabase.Enabled = False
    
    '-------------------------------------------------------------
    If Screen.Width = 15360 And Screen.Height = 11520 Then
        mnuEnableResolutionAlert.Enabled = False
    Else
        mnuEnableResolutionAlert.Enabled = True
    End If
ElseIf intaccount_type = 1 Then
       mnuBypassmsgsender.Enabled = True
       mnuClearMessageSender.Enabled = True
       mnuDeleteData.Enabled = True
       mnuBackupDatabase.Enabled = True
       mnuDebugHandler.Enabled = True
       mnuUserStatusReports.Enabled = True
       mnuCompactDatabase.Enabled = True
       
       '-------------------------------------------------------------
       If Screen.Width = 15360 And Screen.Height = 11520 Then
            mnuEnableResolutionAlert.Enabled = False
       Else
            mnuEnableResolutionAlert.Enabled = True
       End If

End If
Resolution_Set
Check_for_Startup
Prompt_Backup_Dialog
Prompt_for_AutoBackup
If intresolutionAlertenable = 1 Then
    mnuEnableResolutionAlert.Checked = True
ElseIf intresolutionAlertenable = 0 Then
    mnuEnableResolutionAlert.Checked = False
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload frmStyle
Store_User_Logged_Status_Logout
User_Log_Out
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
End Sub


Public Sub Resolution_Set()
On Error Resume Next
If Screen.Width = 15360 And Screen.Height = 11520 Then
   Me.Width = 15360
   Me.Height = 11520
   Theme_Handle 2
ElseIf Screen.Width > 15360 Then
           If blntheme_apply = False Then
                intresmsg = 0
            frmResolutionAlert.Show 1
            
            'If MsgBox("To get a better resolution fit of Stock Controlling System, set the resolution setting to " & vbCrLf & "1024 by 768 pixels or" & vbCrLf & "800 by 600 pixels." & vbCrLf & vbCrLf & "Do you need to set the resolution setting manually?", vbYesNo + vbExclamation) = vbYes Then
                'Shell "Control desk.cpl,,3", vbNormalFocus
                'Store_User_Logged_Status_Logout
                'User_Log_Out
                'End
            'End If
           End If
            'Else
            Me.WindowState = 0
            Me.Width = 15360
            Me.Height = 11520
           'MsgBox Me.Width
            'Frame2.Width = 15000
            Me.Width = Me.Width + 100
          Theme_Handle 2
ElseIf Screen.Width = 12000 And Screen.Height = 9000 Then
    If open_status = True Then
        open_status = False
        If blntheme_apply = False Then
            intresmsg = 1
            frmResolutionAlert.Show 1
            'If MsgBox("Your current screen resolution is 800 by 600 pixels. To get a better resolution fit of Stock Controlling System, set the resolution setting to " & vbCrLf & "1024 by 768 pixels." & vbCrLf & "Anyway you can use your current resolution as well." & vbCrLf & vbCrLf & "Do you need to set the resolution setting manually?", vbYesNo + vbExclamation) = vbYes Then
                'Shell "Control desk.cpl,,3", vbNormalFocus
                'Store_User_Logged_Status_Logout
                'User_Log_Out
                'End
            'End If
        End If
            Me.Width = 12000
            Me.Height = 9000
            Res_set_for_800_600
        'End If
    Else
        Me.Width = 12000
        Me.Height = 9000
        Res_set_for_800_600
    End If
Else
    If open_status = True Then
        open_status = False
        If blntheme_apply = False Then
            intresmsg = 0
            frmResolutionAlert.Show 1
            frmResolutionAlert.Show 1
            'If MsgBox("To get a better resolution fit of Stock Controlling System, set the resolution setting to " & vbCrLf & "1024 by 768 pixels or" & vbCrLf & "800 by 600 pixels." & vbCrLf & vbCrLf & "Do you need to set the resolution setting manually?", vbYesNo + vbExclamation) = vbYes Then
                'Shell "Control desk.cpl,,3", vbNormalFocus
                'Store_User_Logged_Status_Logout
                'User_Log_Out
                'End
            'End If
        End If
            Theme_Handle 3
        'End If
     Else
         Theme_Handle 3
    End If
    
End If
End Sub

Public Sub Res_set_for_800_600()

Frame1.Height = 8787.272: Frame1.Width = 4367.758
Frame1.Top = -130.909: Frame1.Left = 120.907

Picture1.Width = 3975: Picture1.Height = 7815
Picture1.Top = 180: Picture1.Left = 120

Label1.Top = 0: Label2.Top = 960
Label3.Top = 1800: Label4.Top = 2760
Label5.Top = 3840: Label6.Top = 4800
Label7.Top = 5880: Label8.Top = 6840

Image12.Top = 240: Image2.Top = 1200
Image3.Top = 2040: Image4.Top = 3120
Image5.Top = 4080: Image6.Top = 5160
Image7.Top = 6120: Image11.Top = 7040

Line1.X1 = 1440: Line1.X2 = 3720
Line1.Y1 = 120: Line1.Y2 = 120

Line2.X1 = 1560: Line2.X2 = 3720
Line2.Y1 = 1080: Line2.Y2 = 1080

Line3.X1 = 960: Line3.X2 = 3720
Line3.Y1 = 1920: Line3.Y2 = 1920

Line4.X1 = 1440: Line4.X2 = 3720
Line4.Y1 = 2880: Line4.Y2 = 2880

Line5.X1 = 1440: Line5.X2 = 3720
Line5.Y1 = 3960: Line5.Y2 = 3960

Line6.X1 = 1800: Line6.X2 = 3720
Line6.Y1 = 4920: Line6.Y2 = 4920

Line7.X1 = 1800: Line7.X2 = 3720
Line7.Y1 = 6000: Line7.Y2 = 6000

Line8.X1 = 1320: Line8.X2 = 3720
Line8.Y1 = 6960: Line8.Y2 = 6960

cmdBookIssue.Top = 240: cmdBooksReceipt.Top = 1200
cmdReports.Top = 2040: cmdCurStock.Top = 3000
cmdBooksDetails.Top = 4080: cmdMasterInfo.Top = 5050
cmdPassword.Top = 6120: cmdSendMessage.Top = 7080

Frame3.Top = 0: Frame3.Left = 4594.458
Frame3.Width = 7390.428: Frame3.Height = 1194.545

'Picture2.Left = 360
'Picture2.Top = 1680
'Picture2.Width = 6855
'Picture2.Height = 4455

'Picture3.Left = 3720
'Picture3.Top = 240

'Image15.Left = 10
'Image

Picture4.Top = 120: Picture4.Left = 120
Picture4.Width = 7095: Picture4.Height = 735

Image13.Top = -80: Image13.Left = 2400

Image10.Top = 0: Image10.Left = 4800

cmdExit.Top = 20: cmdExit.Left = 5520

cmdLogOffUser.Top = 20: cmdLogOffUser.Left = 3120

Frame2.Top = 1178.182: Frame2.Left = 4594.458
Frame2.Width = 7390.428: Frame2.Height = 7478.182

Image15.Left = 3960: Image15.Top = 240

Image9.Top = 1560: Image9.Left = 240
Theme_Handle 3: Image9.Picture = frmStyle.Image3
Image9.Top = 1750: Image9.Left = 350
Image15.Left = 3850: Image15.Top = 100
End Sub


Private Sub mnuAboutSystem_Click()
frmAbout.Show 1
End Sub

Private Sub mnuActiveUserStatus_Click()
On Error Resume Next
intlogstatusview = 0
frmLoggedUserStatus.Show 1
End Sub

Private Sub mnuBackupDatabase_Click()
frmBackupdatabase.Show 1
End Sub

Private Sub mnuBooksDetails_Click()
frmBookdetails.Show 1
End Sub

Private Sub mnuBooksIssue_Click()
frmBookIssue.Show 1
End Sub

Private Sub mnuBooksnotinuse_Click()
frmBooksNotInUseEntry.Show 1
End Sub

Private Sub mnuBooksnotinusereport_Click()
If Check_For_Privilege(4) = True Then: Exit Sub
Get_Report_for_Books_not_in_stock
End Sub

Private Sub mnuBooksReceipt_Click()
frmBookReceipt.Show 1
End Sub

Private Sub mnuBypassmsgsender_Click()
If user_send_msg_privilege = 0 Then
    MsgBox "You do not have permission for this action...!" & vbCrLf & "Action aborted." & vbCrLf & "Contact Administrator.", vbCritical
    Exit Sub
End If
sender_bypass = True
frmAccessCode.Show 1
End Sub

Private Sub mnuCalculator_Click()
On Error Resume Next
Shell "calc", vbNormalFocus
End Sub

Private Sub mnuChangeresolution_Click()
On Error Resume Next
If MsgBox("Do you want to change the resolution manulally ?", vbYesNo + vbQuestion) = vbYes Then
    Shell "Control desk.cpl,,3", vbNormalFocus
    Store_User_Logged_Status_Logout
    intsystemlogstatus = 7
    Unload Me
    open_status = False
    frmLogin.Show
End If
End Sub

Private Sub mnuClearActiveUser_Click()
On Error Resume Next
blnclearuser = True
frmAccessCode.Show 1
If clearuserok = False Then
    Exit Sub
End If

If strclearuser = User Then
    MsgBox "Action not allowed with current active user.", vbCritical
    Exit Sub
End If


    Set rstclearuser = New ADODB.Recordset
        rstclearuser.CursorLocation = adUseClient
        rstclearuser.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & strclearuser & "'", dbcon, adOpenStatic, adLockOptimistic
            rstclearuser("LOGGED_STATUS") = "0"
            rstclearuser.Update
            MsgBox "Log status of " & strclearuser & " cleared.", vbInformation
    rstclearuser.Close
    Set rstclearuser = Nothing
    'frmLoggedUserStatus.cmdRefresh_Click
    'frmLoggedUserStatus.Timer1_Timer
End Sub

Private Sub mnuClearDatabaseLocation_Click()
frmClearDatabaseLocation.Show 1
End Sub

Private Sub mnuClearMessageSender_Click()
If user_send_msg_privilege = 0 Then
    MsgBox "You do not have permission for this action...!" & vbCrLf & "Action aborted." & vbCrLf & "Contact Administrator.", vbCritical
    Exit Sub
End If
sender_clear = True
frmAccessCode.Show 1
End Sub

Private Sub mnuCompactDatabase_Click()
frmCompactDatabase.Show 1
End Sub

Private Sub mnuContents_Click()
ShellExecute Me.hwnd, "open", App.Path & "\HELP SCS.htm", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub mnuCurrentStock_Click()
frmCurrentStock.Show 1
End Sub

Private Sub mnuDeleteData_Click()
frmDeleteAllRecords.Show 1
End Sub

Private Sub mnuEnableResolutionAlert_Click()
On Error Resume Next
    If mnuEnableResolutionAlert.Checked = True Then
        MsgBox "Resolution Alert already Enabled.", vbInformation
    Else
      reg_obj.RegWrite (Resolution_Alert), "1"
      mnuEnableResolutionAlert.Checked = True
      MsgBox "Resolution Alert Enabled.", vbInformation
    End If
End Sub

Private Sub mnuExittoWindows_Click()
cmdExit_Click
End Sub

Private Sub mnuInstantSystemLogOut_Click()
On Error Resume Next
forcelogout = True
frmAccessCode.Show 1
If proceedforce = False Then
    Exit Sub
End If

If strinstantuser = User Then
    MsgBox "Action not allowed with current active user.", vbCritical
    Exit Sub
End If


    Set rstinstantreceiverlogoff = New ADODB.Recordset
        rstinstantreceiverlogoff.CursorLocation = adUseClient
        rstinstantreceiverlogoff.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & strinstantuser & "'", dbcon, adOpenStatic, adLockOptimistic
            If rstinstantreceiverlogoff("LOGGED_STATUS") = "0" Then
                MsgBox "User has already Logged out the system.", vbExclamation
                Exit Sub
            ElseIf rstinstantreceiverlogoff("LOGGED_STATUS") = "1" Then
                rstinstantreceiverlogoff("INSTANT_LOG_OUT") = "1"
                rstinstantreceiverlogoff.Update
            End If
   'frmLoggedUserStatus.cmdRefresh_Click
   'frmLoggedUserStatus.Timer1_Timer
End Sub

Private Sub mnuInstantUserLogOff_Click()
 On Error Resume Next
 forcelogoff = True
 frmAccessCode.Show 1
 If proceedforce = False Then
    Exit Sub
 End If
If strinstantuser = User Then
    MsgBox "Action not allowed with current active user.", vbCritical
    Exit Sub
End If
    Set rstinstantreceiverlogoff = New ADODB.Recordset
        rstinstantreceiverlogoff.CursorLocation = adUseClient
        rstinstantreceiverlogoff.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & strinstantuser & "'", dbcon, adOpenStatic, adLockOptimistic
            If rstinstantreceiverlogoff("LOGGED_STATUS") = "0" Then
                MsgBox "User has already Logged out the system.", vbExclamation
                Exit Sub
            ElseIf rstinstantreceiverlogoff("LOGGED_STATUS") = "1" Then
                rstinstantreceiverlogoff("INSTANT_LOG_OFF") = "1"
                rstinstantreceiverlogoff.Update
            End If
'frmLoggedUserStatus.cmdRefresh_Click
'frmLoggedUserStatus.Timer1_Timer
End Sub

Private Sub mnuLogOffCurrentUser_Click()
cmdLogOffUser_Click
End Sub

Private Sub mnuMasterInformation_Click()
frmMasterInformation.Show 1
End Sub

Private Sub mnuNotepad_Click()
On Error Resume Next
Shell "notepad", vbNormalFocus
End Sub

Private Sub mnuPerformanceoptimizer_Click()
frmPerformanceoptimizer.Show 1
End Sub

Private Sub mnuReportingStockControllingSystem_Click()
frmReports.Show 1
End Sub

Private Sub mnuRunSCSatSystemStatup_Click()
Dim Msg As String
Msg = MsgBox("Do you want to start Stock Controlling System at Windows start?", vbQuestion + vbYesNoCancel)
If Msg = vbYes Then
Add_to_Reg_Startup
mnuRunSCSatSystemStatup.Checked = True
ElseIf Msg = vbNo Then
Delete_from_Reg_Startup
mnuRunSCSatSystemStatup.Checked = False
End If
End Sub

Private Sub mnuSendaMessage_Click()
mnuSendMessage_Click
End Sub

Private Sub mnuSendMessage_Click()
If user_send_msg_privilege = 0 Then
    MsgBox "You do not have permission to send messages...!" & vbCrLf & "Action aborted." & vbCrLf & "Contact Administrator.", vbCritical
    Exit Sub
End If
CHECK_USER_SENDING_MSG
End Sub

Private Sub mnuThemeSettings_Click()
frmTheme.Show 1
End Sub

Private Sub mnuUserPassword_Click()
On Error Resume Next
frmPassword.Show 1
End Sub

Private Sub mnuUserStatusReports_Click()
frmUserLogStatusReports.Show 1
End Sub

Private Sub Timer1_Timer()
stbMain.Panels(2) = Time
End Sub

Private Sub Timer2_Timer()
Call Check_User_Exist_Acc_Type_Change
End Sub

Private Sub Timer3_Timer()
Check_Message
End Sub

Private Sub Timer4_Timer()
intsetlog = 0
Store_User_Logged_Status_Login
End Sub
Public Sub Add_to_Reg_Startup()
On Error Resume Next
Dim apppath As String
If Right(App.Path, 1) = "\" Then
apppath = App.Path
Else
apppath = App.Path & "\"
End If
reg_obj.RegWrite (Start_Up), apppath & App.EXEName & ".exe"
End Sub
Public Sub Delete_from_Reg_Startup()
On Error Resume Next
reg_obj.RegDelete (Start_Up)
End Sub

Public Sub Check_for_Startup()
On Error GoTo Not_at_Startup
Dim strpath_exist As String
strpath_exist = reg_obj.RegRead(Start_Up)
    If strpath_exist <> "" Then
        If Dir(strpath_exist) <> "" Then
            mnuRunSCSatSystemStatup.Checked = True
        Else
            reg_obj.RegDelete (Start_Up)
            mnuRunSCSatSystemStatup.Checked = False
    End If
Else
    mnuRunSCSatSystemStatup.Checked = False
End If
Exit Sub
Not_at_Startup:
mnuRunSCSatSystemStatup.Checked = False
End Sub

Public Sub Prompt_Backup_Dialog()
On Error GoTo Err
If reg_obj.RegRead(BackupDialog) = 1 Then
    If intaccount_type = 1 Then
        blnbackupprompt = True
    End If
ElseIf reg_obj.RegRead(BackupDialog) = 0 Then
        blnbackupprompt = False
End If
Exit Sub
Err:
blnbackupprompt = False
End Sub

Public Sub Prompt_for_AutoBackup()
On Error GoTo Err
If reg_obj.RegRead(AutoBackup) = 1 Then
        If intaccount_type = 1 Then
            blnautobackup = True
        End If
ElseIf reg_obj.RegRead(AutoBackup) = 0 Then
         blnautobackup = False
End If
Exit Sub
Err:
blnautobackup = False
End Sub
