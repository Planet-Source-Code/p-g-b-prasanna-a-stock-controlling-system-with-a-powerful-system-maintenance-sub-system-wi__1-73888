VERSION 5.00
Begin VB.Form frmFN 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Define New Folder Name"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3465
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   895
      Left            =   100
      ScaleHeight     =   900
      ScaleWidth      =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   3240
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
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
         Left            =   1920
         MouseIcon       =   "frmFN.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   430
         Width           =   1215
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
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
         Left            =   600
         MouseIcon       =   "frmFN.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   430
         Width           =   1215
      End
      Begin VB.TextBox txtFName 
         BackColor       =   &H00FBF4F4&
         Height          =   285
         Left            =   40
         MouseIcon       =   "frmFN.frx":02A4
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   40
         Width           =   3160
      End
   End
End
Attribute VB_Name = "frmFN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
On Error GoTo Err_Check
MkDir frmBF.Dir1.Path & "\" & txtFName.Text
If Right(frmBF.Dir1.Path, 1) = "\" Then
frmBF.Dir1.Path = frmBF.Drive1.Drive & "\" & txtFName.Text
Unload Me
frmBF.Dir1.Refresh
Exit Sub
Else
frmBF.Dir1.Path = frmBF.Dir1.Path & "\" & txtFName.Text
Unload Me
frmBF.Dir1.Refresh
End If
Exit Sub
Err_Check:
If Err.Number = 76 Then
MsgBox "Invalid folder name..", vbCritical
txtFName.Text = ""
txtFName.SetFocus
ElseIf Err.Number = 75 Then
MsgBox "Folder already exist..", vbExclamation
txtFName.SetFocus
SendKeys "{Home}+{End}"
Else
MsgBox Err.Number, vbOKOnly
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtFName.SetFocus
End Sub

Private Sub Form_Load()
Me.Picture = frmStyle.Picture
cmdOk.Enabled = False
Me.Top = frmBF.Top + (frmBF.Height - Me.Height) / 2
Me.Left = frmBF.Left + (frmBF.Width - Me.Width) / 2

End Sub

Private Sub txtFName_Change()
If txtFName.Text <> "" Then
cmdOk.Enabled = True
Else: cmdOk.Enabled = False
End If
End Sub


