VERSION 5.00
Begin VB.Form frmBackupdatabase 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup/Restore Database..."
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBackupdatabase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "C&lose"
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
      Left            =   7250
      MouseIcon       =   "frmBackupdatabase.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Database Backup/Restore Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8595
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7515
         Left            =   40
         ScaleHeight     =   7515
         ScaleWidth      =   8460
         TabIndex        =   15
         Top             =   240
         Width           =   8460
         Begin VB.CheckBox chkAutoBackup 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Create a Backup && Prompt me Automatically..."
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
            Left            =   4080
            MouseIcon       =   "frmBackupdatabase.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   2200
            Width           =   4335
         End
         Begin VB.CheckBox chkPromptmeatload 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Prompt me when programme loads..."
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
            Left            =   840
            MouseIcon       =   "frmBackupdatabase.frx":02B0
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   7080
            Width           =   3735
         End
         Begin VB.CommandButton cmdDeleteBackupFile 
            Caption         =   "Delete Backup File"
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
            Left            =   6000
            MouseIcon       =   "frmBackupdatabase.frx":0402
            MousePointer    =   99  'Custom
            TabIndex        =   8
            ToolTipText     =   "Delete Backup File."
            Top             =   3720
            Width           =   2415
         End
         Begin VB.ComboBox cmbSelectBackupFile 
            BackColor       =   &H00EDFAED&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   3720
            Width           =   4965
         End
         Begin VB.TextBox txtDbLocation 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBF4F4&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   720
            Width           =   6855
         End
         Begin VB.TextBox txtBackupLocationBackup 
            Appearance      =   0  'Flat
            BackColor       =   &H00EFE7FC&
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   1320
            Width           =   6855
         End
         Begin VB.CommandButton cmdBackupSave 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7800
            MouseIcon       =   "frmBackupdatabase.frx":0554
            MousePointer    =   99  'Custom
            TabIndex        =   2
            ToolTipText     =   "Select Backup Saving Location."
            Top             =   1320
            Width           =   615
         End
         Begin VB.CommandButton cmdCreateBackup 
            Caption         =   "&Create Database Backup..."
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
            MouseIcon       =   "frmBackupdatabase.frx":06A6
            MousePointer    =   99  'Custom
            TabIndex        =   3
            ToolTipText     =   "Create Database Backup."
            Top             =   1750
            Width           =   3855
         End
         Begin VB.CommandButton cmdRestoreBackup 
            Caption         =   "Restore Database Backup..."
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
            MouseIcon       =   "frmBackupdatabase.frx":07F8
            MousePointer    =   99  'Custom
            TabIndex        =   9
            ToolTipText     =   "Restore Database Backup."
            Top             =   4920
            Width           =   3855
         End
         Begin VB.TextBox txtDbSourceFile 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBF4F4&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   4440
            Width           =   6855
         End
         Begin VB.CommandButton cmdRestoreBackupFile 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7800
            MouseIcon       =   "frmBackupdatabase.frx":094A
            MousePointer    =   99  'Custom
            TabIndex        =   6
            ToolTipText     =   "Select Backup File Location."
            Top             =   3120
            Width           =   615
         End
         Begin VB.TextBox txtDbBackupLocationRestore 
            Appearance      =   0  'Flat
            BackColor       =   &H00EFE7FC&
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   3120
            Width           =   6855
         End
         Begin VB.CommandButton cmdMessageSend 
            Caption         =   "&Send Message to log out the users"
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
            MouseIcon       =   "frmBackupdatabase.frx":0A9C
            MousePointer    =   99  'Custom
            TabIndex        =   10
            ToolTipText     =   "Send Message."
            Top             =   5520
            Width           =   3855
         End
         Begin VB.CommandButton cmdCheckUserLogStatus 
            Caption         =   "&Monitor Active User Status..."
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
            Left            =   5280
            MouseIcon       =   "frmBackupdatabase.frx":0BEE
            MousePointer    =   99  'Custom
            TabIndex        =   11
            ToolTipText     =   "Monitor Active User Status."
            Top             =   6960
            Width           =   3135
         End
         Begin VB.Image Image8 
            Height          =   645
            Left            =   7800
            Picture         =   "frmBackupdatabase.frx":0D40
            Top             =   4200
            Width           =   645
         End
         Begin VB.Image Image7 
            Height          =   645
            Left            =   7800
            Picture         =   "frmBackupdatabase.frx":23AE
            Top             =   480
            Width           =   645
         End
         Begin VB.Label lblBackupLocationRestore 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Go to Location..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   180
            Left            =   6450
            MouseIcon       =   "frmBackupdatabase.frx":3A1C
            MousePointer    =   99  'Custom
            TabIndex        =   39
            ToolTipText     =   "Go to Backup File Location."
            Top             =   2880
            Width           =   1230
         End
         Begin VB.Label lblBackupLocationBackup 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Go to Location..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   180
            Left            =   6450
            MouseIcon       =   "frmBackupdatabase.frx":3B6E
            MousePointer    =   99  'Custom
            TabIndex        =   14
            ToolTipText     =   "Go to Backup Saving Location."
            Top             =   1080
            Width           =   1230
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select the Backup File to Restore:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   840
            TabIndex        =   38
            Top             =   3480
            Width           =   3060
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Database File Location on the Server/Shared Location/Local Machine"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   840
            TabIndex        =   37
            Top             =   4200
            Width           =   5940
         End
         Begin VB.Image Image6 
            Height          =   465
            Left            =   3480
            Picture         =   "frmBackupdatabase.frx":3CC0
            Top             =   5160
            Width           =   450
         End
         Begin VB.Image Image5 
            Height          =   465
            Left            =   3480
            Picture         =   "frmBackupdatabase.frx":4826
            Top             =   1920
            Width           =   450
         End
         Begin VB.Label lblRestoreTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1320
            TabIndex        =   36
            Top             =   5280
            Width           =   465
         End
         Begin VB.Label lblRestoreDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1320
            TabIndex        =   35
            Top             =   5040
            Width           =   465
         End
         Begin VB.Label lblRestoreUser 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   2760
            TabIndex        =   34
            Top             =   4800
            Width           =   465
         End
         Begin VB.Label lblBackupTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1320
            TabIndex        =   33
            Top             =   2160
            Width           =   465
         End
         Begin VB.Label lblBackupDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1320
            TabIndex        =   32
            Top             =   1920
            Width           =   465
         End
         Begin VB.Label lblBakupUser 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   2520
            TabIndex        =   31
            Top             =   1680
            Width           =   465
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   840
            TabIndex        =   30
            Top             =   5280
            Width           =   390
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   840
            TabIndex        =   29
            Top             =   5040
            Width           =   390
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Restoration Done by:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   840
            TabIndex        =   28
            Top             =   4800
            Width           =   1845
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   840
            TabIndex        =   27
            Top             =   2160
            Width           =   390
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   840
            TabIndex        =   26
            Top             =   1920
            Width           =   390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Last Backup Done by:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   840
            TabIndex        =   25
            Top             =   1680
            Width           =   1590
         End
         Begin VB.Image Image4 
            Height          =   645
            Left            =   4680
            Picture         =   "frmBackupdatabase.frx":538C
            Top             =   6840
            Width           =   435
         End
         Begin VB.Image Image3 
            Height          =   735
            Left            =   45
            Picture         =   "frmBackupdatabase.frx":6296
            Top             =   6120
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Database File Location on the Server/Shared Location/Local Machine"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   840
            TabIndex        =   24
            Top             =   480
            Width           =   5940
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Locate the Database Backup Saving Path:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   120
            Left            =   840
            TabIndex        =   23
            Top             =   1080
            Width           =   3510
         End
         Begin VB.Image Image1 
            Height          =   705
            Left            =   40
            Picture         =   "frmBackupdatabase.frx":7F2C
            Top             =   720
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Backup Database"
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
            TabIndex        =   22
            Top             =   120
            Width           =   1530
         End
         Begin VB.Line Line5 
            X1              =   1800
            X2              =   8400
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Image Image2 
            Height          =   720
            Left            =   45
            Picture         =   "frmBackupdatabase.frx":99DE
            Top             =   3120
            Width           =   675
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Locate the Backup File Location:"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   840
            TabIndex        =   21
            Top             =   2880
            Width           =   3375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Restore Database"
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
            Top             =   2520
            Width           =   1545
         End
         Begin VB.Line Line6 
            X1              =   1800
            X2              =   8400
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmBackupdatabase.frx":B3A0
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   840
            TabIndex        =   19
            Top             =   6120
            Width           =   7575
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
            TabIndex        =   18
            Top             =   5760
            Width           =   1335
         End
      End
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   7080
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   7080
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   7080
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7080
      Y1              =   8160
      Y2              =   8160
   End
End
Attribute VB_Name = "frmBackupdatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dbBakcon As ADODB.Connection
Private rstBackup As ADODB.Recordset
Private rstRestore As ADODB.Recordset
Private dbBaklog_path As String

Private Sub chkAutoBackup_Click()
On Error Resume Next
If chkAutoBackup.Value = 1 Then
    If chkPromptmeatload.Value = 1 Then
        chkPromptmeatload.Value = 0
    End If
End If
If chkAutoBackup.Value = 1 Then
    reg_obj.RegWrite (AutoBackup), "1"
ElseIf chkAutoBackup.Value = 0 Then
    reg_obj.RegWrite (AutoBackup), "0"
End If
blnautobackup = False
End Sub

Private Sub chkPromptmeatload_Click()
On Error Resume Next
If chkPromptmeatload.Value = 1 Then
    If chkAutoBackup.Value = 1 Then
        chkAutoBackup.Value = 0
    End If
End If
If chkPromptmeatload.Value = 1 Then
    reg_obj.RegWrite (BackupDialog), "1"
ElseIf chkPromptmeatload.Value = 0 Then
    reg_obj.RegWrite (BackupDialog), "0"
End If
End Sub

Private Sub cmbSelectBackupFile_Click()
If cmbSelectBackupFile = "" Then
    cmdDeleteBackupFile.Enabled = False
    cmdRestoreBackup.Enabled = False
Else
    cmdDeleteBackupFile.Enabled = True
    cmdRestoreBackup.Enabled = True
End If
End Sub

Private Sub cmdBackupSave_Click()
On Error GoTo Err
frmBF.Show 1
Exit Sub
Err:
End Sub

Private Sub cmdCheckUserLogStatus_Click()
On Error Resume Next
intlogstatusview = 1
frmLoggedUserStatus.Show 1
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCreateBackup_Click()
On Error GoTo Err
Dim strgetdate, strgettime As String
If Dir(txtBackupLocationBackup, vbDirectory) = "" Then
    MsgBox "Please select the Database Backup Saving Location.", vbExclamation
    cmdBackupSave_Click
    Exit Sub
ElseIf Dir(txtBackupLocationBackup, vbDirectory) <> "" Then
    strgetdate = Format(Date, "dd-MM-yyyy")
    strgettime = Format(Time, "hh.mm.ss AMPM")
    Me.Caption = "Backup/Restore Database...          - Backing up Database...,Please wait..."
    CopyFile txtDbLocation, txtBackupLocationBackup & strgetdate & " - " & strgettime & ".dbf", 1
    Set_Backup_Details
    Set_Details
    Get_Backup_Files
    Me.Caption = "Backup/Restore Database..."
    MsgBox "Database backup created as " & strgetdate & " - " & strgettime & ".dbf.", vbInformation
End If
Exit Sub
Err:
MsgBox "Error occurred while processing." & vbCrLf & "Check the file locations.", vbCritical
End Sub

'Private Sub cmdLocateforBackup_Click()
'On Error GoTo Err
'With frmStyle.cdDatabaseselect
'.'CancelError = True
'ReturnOpen:
'.DialogTitle = "Please Locate the Database to Backup..."
'.Filter = "Database (*.mdb) |*.mdb"
'.ShowOpen
'    If Right(.FileName, 12) <> "db_Stock.mdb" Then
'        MsgBox "The Database must be db_Stock.mdb.", vbExclamation
'        GoTo ReturnOpen
'    Else
'        On Error Resume Next
'        reg_obj.RegWrite (Db_Location), .FileName
'        txtDbLocation = .FileName
'    End If
'End With
'Exit Sub
'Err:
'End Sub

Private Sub cmdDeleteBackupFile_Click()
On Error GoTo Err
Dim strlocationfordelete, strgetfullfilepathfordelete As String
strlocationfordelete = txtDbBackupLocationRestore
If Right(strlocationfordelete, 1) <> "\" Then
    strlocationfordelete = strlocationfordelete & "\"
    strgetfullfilepathfordelete = strlocationfordelete & cmbSelectBackupFile
    
ElseIf Right(strlocationfordelete, 1) = "\" Then
    strgetfullfilepathfordelete = strlocationfordelete & cmbSelectBackupFile
End If

'MsgBox strgetfullfilepathfordelete
 If Dir(strgetfullfilepathfordelete) = "" Then
     MsgBox "Please check the Database Backup File.", vbExclamation
     cmbSelectBackupFile.Clear
     Get_Backup_Files
     Exit Sub
 Else
    If MsgBox("Are you sure you want to delete the Backup File: " & cmbSelectBackupFile & "?", vbYesNo + vbQuestion) = vbYes Then
        Kill strgetfullfilepathfordelete
        Get_Backup_Files
        MsgBox "Selected File deleted successfully.", vbInformation
    End If
End If
Exit Sub
Err:
MsgBox Err.Description & "_" & Err.Number, vbCritical
End Sub

Private Sub cmdMessageSend_Click()
Call cmdSendMessage_Click
End Sub

Private Sub cmdRestoreBackup_Click()
On Error GoTo Err
Dim strlocation, strgetfullfilepath As String
inttask = 1
Check_For_Users_Exist
If intdbok = 0 Then
    Exit Sub
End If
strlocation = txtDbBackupLocationRestore
If Right(strlocation, 1) <> "\" Then
    strlocation = strlocation & "\"
    strgetfullfilepath = strlocation & cmbSelectBackupFile
    
ElseIf Right(strlocation, 1) = "\" Then
        strgetfullfilepath = strlocation & cmbSelectBackupFile
End If

'MsgBox strgetfullfilepath
If cmbSelectBackupFile = "" Then
    MsgBox "Please check the Database Backup Location/File.", vbExclamation
    cmbSelectBackupFile.Clear
    Get_Backup_Files
    cmdRestoreBackupFile_Click
    Exit Sub
End If

If Dir(strgetfullfilepath) = "" Then
     MsgBox "Please check the Database Backup Location/File.", vbExclamation
     cmbSelectBackupFile.Clear
     Get_Backup_Files
     cmdRestoreBackupFile_Click
    Exit Sub
Else

   If MsgBox("Are you sure you need to Restore the Database?" & vbCrLf & vbCrLf & "Tip:" & vbCrLf & _
             "----" & vbCrLf & "If you click Yes your Database connectivity will be closed and restored the Database." & vbCrLf & _
             "No rollback operation can be done." & vbCrLf & "Current Database will be permanently deleted and " & vbCrLf & _
             "replace with the selected backup." & vbCrLf & vbCrLf & "It is recommended that you backup the current database first.", vbYesNo + vbInformation) = vbYes Then
        On Error Resume Next
        Me.Caption = "Backup/Restore Database...          - Restoring Database...,Please wait..."
        frmMain.Timer1.Enabled = False: frmMain.Timer2.Enabled = False
        frmMain.Timer3.Enabled = False: frmMain.Timer4.Enabled = False
        'As this is a Database Restoration No Log out log will be created."
        'Store_User_Logged_Status_Logout
        'intsystemlogstatus = 9
        'User_Log_Out
        dbcon.Close
        Set dbcon = Nothing
        deGroupedreports.cnGroupedreports.Close
        Kill txtDbSourceFile
        CopyFile strgetfullfilepath, txtDbSourceFile, 1
        Set_Restore_Details
        Me.Caption = "Backup/Restore Database..."
        MsgBox "Database Restoration succeeded.", vbInformation
        End
   End If
End If
Exit Sub
Err:
MsgBox "Error occurred while processing." & vbCrLf & "Check the file locations.", vbCritical
End Sub

Private Sub cmdRestoreBackupFile_Click()
On Error GoTo Err
cmdBackupSave_Click
Exit Sub
Err:
End Sub
Private Sub cmdSendMessage_Click()
frmMain.cmdSendMessage_Click
End Sub

Private Sub Form_Activate()
On Error GoTo Err
cmdCreateBackup.SetFocus
Exit Sub
Err:
    cmdBackupSave.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim strdbbackuplocation As String
Me.Picture = frmStyle.Picture
BakLog_Details_connection
Set_Details
txtDbLocation = Database_Path
txtDbSourceFile = Database_Path
strdbbackuplocation = reg_obj.RegRead(Db_Backup_Location)
If strdbbackuplocation <> "" Then
    If Dir(strdbbackuplocation, vbDirectory) = "" Then
        txtBackupLocationBackup = "": txtDbBackupLocationRestore = ""
        cmdRestoreBackup.Enabled = False: cmdCreateBackup.Enabled = False
        lblBackupLocationBackup.Enabled = False: lblBackupLocationRestore.Enabled = False
    Else
        txtBackupLocationBackup = strdbbackuplocation: txtDbBackupLocationRestore = strdbbackuplocation
        cmdRestoreBackup.Enabled = True: cmdCreateBackup.Enabled = True
        lblBackupLocationBackup.Enabled = True: lblBackupLocationRestore.Enabled = True
    End If
ElseIf strdbbackuplocation = "" Then
    txtBackupLocationBackup = "": txtDbBackupLocationRestore = ""
    cmdRestoreBackup.Enabled = False: cmdCreateBackup.Enabled = False
    lblBackupLocationBackup.Enabled = False: lblBackupLocationRestore.Enabled = False
End If
Get_Backup_Files
On Error GoTo Err
If reg_obj.RegRead(AutoBackup) = 1 Then
    If blnautobackup = True Then
        blnautobackup = False
        cmdCreateBackup_Click
    End If
End If
chkAutoBackup.Value = reg_obj.RegRead(AutoBackup)
chkPromptmeatload.Value = reg_obj.RegRead(BackupDialog)
Exit Sub
Err:
End Sub
Public Sub BakLog_Details_connection()
On Error Resume Next
dbBaklog_path = Mid(Database_Path, 1, Len(Database_Path) - 12) & "baklog.sav"
Set dbBakcon = New ADODB.Connection
dbBakcon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbBaklog_path & ";Persist Security Info=False;Jet OLEDB:Database Password=" & pword_db
dbBakcon.Open
End Sub

Public Sub Set_Details()
On Error Resume Next
Set rstBackup = New ADODB.Recordset
     rstBackup.CursorLocation = adUseClient
     rstBackup.Open "SELECT * FROM DB_BACKUP ORDER BY ID", dbBakcon, adOpenStatic, adLockOptimistic
Set rstRestore = New ADODB.Recordset
     rstRestore.CursorLocation = adUseClient
     rstRestore.Open "SELECT * FROM DB_RESTORE ORDER BY ID", dbBakcon, adOpenStatic, adLockOptimistic
'set data
If rstBackup.RecordCount = 0 Then
    rstBackup.AddNew
    rstBackup("USER") = "None": rstBackup("DATE") = "None": rstBackup("TIME") = "None"
    rstBackup.Update
End If
If rstRestore.RecordCount = 0 Then
    rstRestore.AddNew
    rstRestore("USER") = "None"
    rstRestore("DATE") = "None"
    rstRestore("TIME") = "None"
    rstRestore.Update
End If
lblBakupUser.Caption = rstBackup("USER")
lblBackupDate.Caption = rstBackup("DATE")
lblBackupTime.Caption = rstBackup("TIME")

lblRestoreUser.Caption = rstRestore("USER")
lblRestoreDate.Caption = rstRestore("DATE")
lblRestoreTime.Caption = rstRestore("TIME")
End Sub

Public Sub Set_Backup_Details()
On Error Resume Next
 rstBackup("USER") = User
 rstBackup("DATE") = Format(Date, "dd/MM/yyyy")
 rstBackup("TIME") = Time
 rstBackup.Update
End Sub

Public Sub Set_Restore_Details()
On Error Resume Next
rstRestore("USER") = User
rstRestore("DATE") = Format(Date, "dd/MM/yyyy")
rstRestore("TIME") = Time
rstRestore.Update
End Sub

Public Sub Get_Backup_Files()
On Error Resume Next
cmbSelectBackupFile.Clear
    Dim f, f1, fc, s
    Set f = fs.GetFolder(txtDbBackupLocationRestore)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
        If Right(s, 4) = ".dbf" Then
            cmbSelectBackupFile.AddItem s
        End If
        Next
cmbSelectBackupFile = cmbSelectBackupFile.List(0)
If cmbSelectBackupFile.List(0) = "" Then
    cmdDeleteBackupFile.Enabled = False
    cmdRestoreBackup.Enabled = False
Else
    cmdDeleteBackupFile.Enabled = True
    cmdRestoreBackup.Enabled = True
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
rstBackup.Close
Set rstBackup = Nothing

rstRestore.Close
Set rstRestore = Nothing

dbBakcon.Close
Set dbBakcon = Nothing

End Sub

Private Sub lblBackupLocationBackup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBackupLocationBackup.Left = lblBackupLocationBackup.Left + 20
lblBackupLocationBackup.Top = lblBackupLocationBackup.Top + 20
End Sub

Private Sub lblBackupLocationBackup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblBackupLocationBackup.ForeColor = &HFF8080 Then
    lblBackupLocationBackup.ForeColor = &HFF80FF
ElseIf lblBackupLocationBackup.ForeColor = &HFF80FF Then
    lblBackupLocationBackup.ForeColor = &HFF8080
End If
End Sub

Private Sub lblBackupLocationBackup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBackupLocationBackup.Left = lblBackupLocationBackup.Left - 20
lblBackupLocationBackup.Top = lblBackupLocationBackup.Top - 20
If Dir(txtBackupLocationBackup, vbDirectory) <> "" Then
    Shell "explorer " & txtBackupLocationBackup, vbNormalFocus
ElseIf Dir(txtBackupLocationBackup, vbDirectory) = "" Then
    MsgBox "Invalid Path.", vbCritical
End If
End Sub

Private Sub lblBackupLocationRestore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBackupLocationRestore.Left = lblBackupLocationRestore.Left + 20
lblBackupLocationRestore.Top = lblBackupLocationRestore.Top + 20
End Sub

Private Sub lblBackupLocationRestore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblBackupLocationRestore.ForeColor = &HFF8080 Then
    lblBackupLocationRestore.ForeColor = &HFF80FF
ElseIf lblBackupLocationRestore.ForeColor = &HFF80FF Then
    lblBackupLocationRestore.ForeColor = &HFF8080
End If
End Sub

Private Sub lblBackupLocationRestore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Dir(txtDbBackupLocationRestore, vbDirectory) <> "" Then
    Shell "explorer " & txtDbBackupLocationRestore, vbNormalFocus
ElseIf Dir(txtDbBackupLocationRestore, vbDirectory) = "" Then
    MsgBox "Invalid Path.", vbCritical
End If
lblBackupLocationRestore.Left = lblBackupLocationRestore.Left - 20
lblBackupLocationRestore.Top = lblBackupLocationRestore.Top - 20
End Sub

Private Sub txtBackupLocationBackup_Change()
If txtBackupLocationBackup <> "" Then
    cmdCreateBackup.Enabled = True
    lblBackupLocationBackup.Enabled = True
ElseIf txtBackupLocationBackup = "" Then
    cmdCreateBackup.Enabled = False
    lblBackupLocationBackup.Enabled = False
End If
End Sub

Private Sub txtDbBackupLocationRestore_Change()
If txtDbBackupLocationRestore <> "" Then
    cmdRestoreBackup.Enabled = True
    lblBackupLocationRestore.Enabled = True
ElseIf txtDbBackupLocationRestore = "" Then
    cmdRestoreBackup.Enabled = False
    lblBackupLocationRestore.Enabled = False
End If
End Sub
