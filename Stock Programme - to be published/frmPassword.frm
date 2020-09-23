VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Account/Password/Privileges..."
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   4215
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCreateAccount 
      Caption         =   "Create User Account"
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
      Left            =   1200
      MouseIcon       =   "frmPassword.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "Change Password/Privileges"
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
      Left            =   1200
      MouseIcon       =   "frmPassword.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   3720
      Width           =   2655
   End
   Begin VB.CommandButton cmdDeleteAccount 
      Caption         =   "Delete User Account"
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
      Left            =   1200
      MouseIcon       =   "frmPassword.frx":02B0
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
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
      Left            =   2640
      MouseIcon       =   "frmPassword.frx":0402
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Account/Password Settings"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4035
         Left            =   120
         ScaleHeight     =   4035
         ScaleWidth      =   3735
         TabIndex        =   6
         Top             =   240
         Width           =   3735
         Begin VB.ComboBox cmbUsernames 
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
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   600
            Width           =   2655
         End
         Begin VB.Image Image2 
            Height          =   705
            Left            =   0
            Picture         =   "frmPassword.frx":0554
            Top             =   3240
            Width           =   765
         End
         Begin VB.Image Image1 
            Height          =   720
            Index           =   0
            Left            =   0
            Picture         =   "frmPassword.frx":223A
            Top             =   1920
            Width           =   780
         End
         Begin VB.Image Image4 
            Height          =   525
            Left            =   0
            Picture         =   "frmPassword.frx":3FBC
            Top             =   720
            Width           =   555
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Delete Account"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Width           =   1335
         End
         Begin VB.Line Line2 
            X1              =   1440
            X2              =   3600
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line1 
            X1              =   1440
            X2              =   3600
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Create Account"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   1680
            Width           =   1455
         End
         Begin VB.Line Line3 
            X1              =   2520
            X2              =   3600
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Change Password/Privileges"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Top             =   2880
            Width           =   2535
         End
      End
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   2520
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   2520
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   2520
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   2520
      Y1              =   4560
      Y2              =   4560
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstusernames As ADODB.Recordset
Private Sub cmdChangePassword_Click()
frmChangePassword.Show 1
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdCreateAccount_Click()
frmCreateAccount.Show 1
End Sub

Private Sub cmdDeleteAccount_Click()
If cmbUsernames = User Then
    MsgBox "You can't delete the current account." & vbCrLf & "Please Log In with another Administrator Account.", vbCritical
    Exit Sub
End If
 Set rstusernamefordelete = New ADODB.Recordset
     rstusernamefordelete.CursorLocation = adUseClient
      rstusernamefordelete.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames.Text & "'", dbcon, adOpenStatic, adLockOptimistic
If MsgBox("Are you sure you need to delete the account " & cmbUsernames & " ?", vbYesNo + vbQuestion) = vbYes Then
                If rstusernamefordelete.RecordCount > 0 Then
                    If UCase(rstusernamefordelete("USER_NAME")) = UCase("Administrator") Then
                        MsgBox "The Account " & cmbUsernames & " is a built-in account." & vbCrLf & "It cannot be deleted.", vbCritical
                     Exit Sub
                    End If
                   On Error GoTo Err
                   rstusernamefordelete.Delete
                   MsgBox "User Account " & cmbUsernames & " successfully deleted.", vbInformation
                   cmbUsernames.Clear
                   Form_Load
                End If
End If
rstusernamefordelete.Close
Set rstusernamefordelete = Nothing
Exit Sub
Err:
    MsgBox Err.Description & " _ " & Err.Number & "."
    Unload Me
End Sub

Public Sub Form_Load()
On Error GoTo Err
Set rstusernames = New ADODB.Recordset
    rstusernames.CursorLocation = adUseClient
    rstusernames.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
Add_User_Names_to_Combo
On Error Resume Next
cmbUsernames.Text = cmbUsernames.List(0)
Account_type_initialize
Me.Picture = frmStyle.Picture
Me.Top = frmMain.Top + 2500
Me.Left = frmMain.Left + (frmMain.Width - Me.Width) / 2
Exit Sub
Err:
MsgBox Err.Description & " _" & Err.Number, vbCritical
Unload Me
End Sub
Public Sub Add_User_Names_to_Combo()
If rstusernames.RecordCount > 0 Then
    Do While Not rstusernames.EOF
        cmbUsernames.AddItem rstusernames("USER_NAME")
    rstusernames.MoveNext
    Loop
End If
End Sub

Public Sub Account_type_initialize()
If intaccount_type = 0 Then
    cmdCreateAccount.Enabled = False
    cmdDeleteAccount.Enabled = False
    cmbUsernames.Enabled = False
    End If
cmbUsernames = User
End Sub
