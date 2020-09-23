VERSION 5.00
Begin VB.Form frmDatabaseSelectionMsg 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Controlling System - Database Configuration"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   ControlBox      =   0   'False
   Icon            =   "frmDatabaseSelectionMsg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5535
      ScaleWidth      =   7935
      TabIndex        =   0
      Top             =   120
      Width           =   7935
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
         Left            =   6360
         MouseIcon       =   "frmDatabaseSelectionMsg.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5175
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   6975
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   4980
            ScaleHeight     =   495
            ScaleWidth      =   1815
            TabIndex        =   25
            Top             =   360
            Width           =   1815
            Begin VB.Image Image3 
               Height          =   375
               Left            =   80
               Picture         =   "frmDatabaseSelectionMsg.frx":015E
               Top             =   0
               Width           =   1605
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            Height          =   2895
            Left            =   120
            TabIndex        =   20
            Top             =   1560
            Width           =   6735
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   2535
               Left            =   120
               ScaleHeight     =   2535
               ScaleWidth      =   6495
               TabIndex        =   21
               Top             =   240
               Width           =   6495
               Begin VB.CommandButton cmdConvert 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   5160
                  MouseIcon       =   "frmDatabaseSelectionMsg.frx":2144
                  MousePointer    =   99  'Custom
                  TabIndex        =   9
                  Top             =   2160
                  Width           =   1215
               End
               Begin VB.OptionButton optConvertBackup 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Convert an existing Database Backup File (*.dbf) as the Database"
                  Height          =   255
                  Left            =   120
                  MouseIcon       =   "frmDatabaseSelectionMsg.frx":2296
                  MousePointer    =   99  'Custom
                  TabIndex        =   8
                  Top             =   2160
                  Width           =   5000
               End
               Begin VB.OptionButton optLocal 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Database Locates on Local Computer"
                  Height          =   375
                  Left            =   120
                  MouseIcon       =   "frmDatabaseSelectionMsg.frx":23E8
                  MousePointer    =   99  'Custom
                  TabIndex        =   1
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   3135
               End
               Begin VB.CommandButton cmdSelectLocal 
                  Caption         =   "..."
                  Default         =   -1  'True
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   5160
                  MouseIcon       =   "frmDatabaseSelectionMsg.frx":253A
                  MousePointer    =   99  'Custom
                  TabIndex        =   2
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.OptionButton optMappedDrive 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Database Locates on a File Server mapped by a Network Drive"
                  Height          =   375
                  Left            =   120
                  MouseIcon       =   "frmDatabaseSelectionMsg.frx":268C
                  MousePointer    =   99  'Custom
                  TabIndex        =   3
                  Top             =   600
                  Width           =   4935
               End
               Begin VB.OptionButton optShared 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Database is in a Shared Location on a File Server"
                  Height          =   375
                  Left            =   120
                  MouseIcon       =   "frmDatabaseSelectionMsg.frx":27DE
                  MousePointer    =   99  'Custom
                  TabIndex        =   5
                  Top             =   1080
                  Width           =   4215
               End
               Begin VB.TextBox txtShared 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FBF4F4&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   360
                  TabIndex        =   7
                  Top             =   1680
                  Width           =   6015
               End
               Begin VB.CommandButton cmdSelectMapped 
                  Caption         =   "..."
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   5160
                  MouseIcon       =   "frmDatabaseSelectionMsg.frx":2930
                  MousePointer    =   99  'Custom
                  TabIndex        =   4
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.CommandButton cmdConfigure 
                  Caption         =   "C&onfigure..."
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
                  Left            =   5160
                  MouseIcon       =   "frmDatabaseSelectionMsg.frx":2A82
                  MousePointer    =   99  'Custom
                  TabIndex        =   6
                  Top             =   1080
                  Width           =   1215
               End
               Begin VB.Label Label9 
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "(Example: \\server\dbshared folder\db_Stock.mdb)"
                  ForeColor       =   &H000000C0&
                  Height          =   255
                  Left            =   360
                  TabIndex        =   22
                  Top             =   1440
                  Width           =   3855
               End
            End
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "077 - 9728092"
            Height          =   195
            Left            =   2900
            MouseIcon       =   "frmDatabaseSelectionMsg.frx":2BD4
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   4680
            Width           =   1035
         End
         Begin VB.Label lblEmail 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "pgbsoft@gmail.com"
            Height          =   195
            Left            =   600
            MouseIcon       =   "frmDatabaseSelectionMsg.frx":2D26
            MousePointer    =   99  'Custom
            TabIndex        =   27
            Top             =   4875
            Width           =   1395
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "For any technical issue contact me on:"
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   4680
            Width           =   2730
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "This Database Programme was specially developed for "
            Height          =   195
            Left            =   2800
            TabIndex        =   24
            Top             =   120
            Width           =   3915
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email: "
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   4875
            Width           =   465
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Database name must be db_Stock.mdb."
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
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   3495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-----"
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
            Top             =   1080
            Width           =   315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tip:"
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
            Top             =   960
            Width           =   345
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*Please, Select the Database.*"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   600
            Width           =   2190
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Database not found."
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
            TabIndex        =   12
            Top             =   120
            Width           =   1755
         End
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
         TabIndex        =   18
         Top             =   480
         Width           =   75
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   960
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   120
         Width           =   75
      End
      Begin VB.Image Image1 
         Height          =   705
         Left            =   120
         Picture         =   "frmDatabaseSelectionMsg.frx":2E78
         Top             =   240
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmDatabaseSelectionMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub cmdConfigure_Click()
    If Right(txtShared, 12) <> "db_Stock.mdb" Then
        MsgBox "The Database must be db_Stock.mdb.", vbExclamation
       txtShared.SetFocus
       SendKeys "{HOME}+{END}"
       Exit Sub
    End If
On Error Resume Next
   Located_Database = txtShared
   reg_obj.RegWrite (Database_Path_Store), Located_Database
   Database_Path = Located_Database
   MsgBox "Database Configured successfully.", vbInformation
Unload Me
End Sub

Private Sub cmdConfigure_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub cmdConvert_Click()
intbackupconvert = 1
Unload Me
Select_Database
End Sub

Private Sub cmdConvert_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub cmdSelectLocal_Click()
Unload Me
Select_Database
End Sub

Private Sub cmdSelectLocal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub cmdSelectMapped_Click()
Unload Me
Select_Database
End Sub

Private Sub cmdSelectMapped_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Form_Activate()
On Error Resume Next
cmdOk.SetFocus
End Sub

Private Sub Form_Load()
cmdConfigure.Enabled = False
Me.Picture = frmStyle.Picture
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub lblEmail_Click()
ShellExecute hwnd, "Open", "mailto:pgbsoft@gmail.com", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = True
End Sub

Private Sub lblLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub optConvertBackup_Click()
cmdConvert.Enabled = True
txtShared = ""
End Sub

Private Sub optConvertBackup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub optLocal_Click()
cmdSelectLocal.Enabled = True
cmdSelectMapped.Enabled = False
cmdConvert.Enabled = False
txtShared = ""
txtShared.Enabled = False
End Sub

Private Sub optLocal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub optMappedDrive_Click()
cmdSelectMapped.Enabled = True
cmdSelectLocal.Enabled = False
cmdConvert.Enabled = False
txtShared = ""
txtShared.Enabled = False
End Sub

Private Sub optMappedDrive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub optShared_Click()
cmdSelectLocal.Enabled = False
cmdSelectMapped.Enabled = False
cmdConvert.Enabled = False
txtShared.Enabled = True
txtShare = ""
txtShared.SetFocus
End Sub

Private Sub optShared_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub

Private Sub txtShared_Change()
If txtShared <> "" Then
    cmdConfigure.Enabled = True
Else
    cmdConfigure.Enabled = False
End If
End Sub

Private Sub txtShared_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblEmail.FontUnderline = False
End Sub
