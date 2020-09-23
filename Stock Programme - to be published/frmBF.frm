VERSION 5.00
Begin VB.Form frmBF 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse for folders"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   Icon            =   "frmBF.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3900
      Left            =   100
      TabIndex        =   5
      Top             =   360
      Width           =   4770
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3650
         Left            =   40
         ScaleHeight     =   3645
         ScaleWidth      =   4680
         TabIndex        =   6
         Top             =   120
         Width           =   4680
         Begin VB.CommandButton cmdOk 
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
            Left            =   3600
            MouseIcon       =   "frmBF.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   3240
            Width           =   975
         End
         Begin VB.CommandButton cmdNewFolder 
            Caption         =   "&New Folder"
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
            Left            =   2160
            MouseIcon       =   "frmBF.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   3240
            Width           =   1335
         End
         Begin VB.DirListBox Dir1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2340
            Left            =   40
            MouseIcon       =   "frmBF.frx":02B0
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   520
            Width           =   4575
         End
         Begin VB.DriveListBox Drive1 
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
            Height          =   315
            Left            =   40
            MouseIcon       =   "frmBF.frx":0402
            MousePointer    =   99  'Custom
            TabIndex        =   1
            Top             =   40
            Width           =   4575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            Height          =   15
            Left            =   45
            TabIndex        =   7
            Top             =   3075
            Width           =   4575
         End
         Begin VB.Line Line1 
            X1              =   45
            X2              =   2045
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line2 
            X1              =   45
            X2              =   2045
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line3 
            X1              =   45
            X2              =   2045
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Line Line4 
            X1              =   45
            X2              =   2045
            Y1              =   3600
            Y2              =   3600
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the Location:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmBF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ret As String
Public int_click As Boolean
Private Sub cmdNewFolder_Click()
frmFN.Show 1
End Sub

Private Sub cmdNewFolder_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub
Private Sub cmdOk_Click()
Dim retfs As String
On Error Resume Next
If Right(Dir1.Path, 1) = "\" Then
    ret = Dir1.Path
    retfs = UCase(Left(ret, 1)) + Mid(ret, 2, Len(ret) - 1)
    reg_obj.RegWrite (Db_Backup_Location), retfs
    frmBackupdatabase.txtBackupLocationBackup = retfs
    frmBackupdatabase.txtDbBackupLocationRestore = retfs
    frmBackupdatabase.cmbSelectBackupFile.Clear
    frmBackupdatabase.Get_Backup_Files
    Unload Me
Else
    ret = Dir1.Path & "\"
    retfs = UCase(Left(ret, 1)) + Mid(ret, 2, Len(ret) - 1)
    reg_obj.RegWrite (Db_Backup_Location), retfs
    frmBackupdatabase.txtBackupLocationBackup = retfs
    frmBackupdatabase.txtDbBackupLocationRestore = retfs
    frmBackupdatabase.cmbSelectBackupFile.Clear
    frmBackupdatabase.Get_Backup_Files
    Unload Me
End If
End Sub

Private Sub cmdOk_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub
Private Sub Dir1_Click()
Dir1.Path = Dir1.Path
End Sub

Private Sub Dir1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Drive1_Change()
On Error GoTo Err_Check
Dir1.Path = Drive1.Drive
Exit Sub
Err_Check:
MsgBox "Device is not ready..!", vbCritical
Drive1.Refresh
End Sub

Private Sub Drive1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Form_Activate()
cmdOk.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then: Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Err
Dim drvletter, fullpath As String

Me.Picture = frmStyle.Picture
Me.Top = frmBackupdatabase.Top + (frmBackupdatabase.Height - Me.Height) / 2
Me.Left = frmBackupdatabase.Left + (frmBackupdatabase.Width - Me.Width) / 2

drvletter = Left(reg_obj.RegRead(Db_Backup_Location), 1)
fullpath = reg_obj.RegRead(Db_Backup_Location)
Drive1.Drive = drvletter
Dir1.Path = fullpath
    If Dir(fullpath, vbDirectory) = "" Then
        Err_Path
    End If
Exit Sub
Err:
Err_Path
End Sub
Public Sub Err_Path()
Drive1.Drive = UCase(Left(App.Path, 1))
Dir1.Path = App.Path
End Sub

