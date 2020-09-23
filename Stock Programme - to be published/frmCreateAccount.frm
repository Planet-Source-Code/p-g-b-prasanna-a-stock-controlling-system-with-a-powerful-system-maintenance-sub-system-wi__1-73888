VERSION 5.00
Begin VB.Form frmCreateAccount 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create User Account"
   ClientHeight    =   4065
   ClientLeft      =   7245
   ClientTop       =   6435
   ClientWidth     =   5745
   Icon            =   "frmCreateAccount.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3915
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   5520
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   40
         ScaleHeight     =   3735
         ScaleWidth      =   5415
         TabIndex        =   9
         Top             =   120
         Width           =   5415
         Begin VB.CheckBox chkShowPassword 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Show password"
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
            TabIndex        =   4
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txtNewUserName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2520
            TabIndex        =   1
            Top             =   480
            Width           =   2775
         End
         Begin VB.TextBox txtNewPassword 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2520
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox txtConfirmPassword 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2520
            PasswordChar    =   "*"
            TabIndex        =   3
            Top             =   1200
            Width           =   2775
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
            Left            =   4320
            MouseIcon       =   "frmCreateAccount.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   3240
            Width           =   975
         End
         Begin VB.CommandButton cmdCreate 
            Caption         =   "Create"
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
            Left            =   2640
            MouseIcon       =   "frmCreateAccount.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   3240
            Width           =   1575
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   255
            Left            =   1080
            ScaleHeight     =   255
            ScaleWidth      =   3855
            TabIndex        =   10
            Top             =   2520
            Width           =   3855
            Begin VB.OptionButton optLimited 
               BackColor       =   &H00FFFFFF&
               Caption         =   "User Limited"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2280
               TabIndex        =   6
               Top             =   0
               Width           =   1695
            End
            Begin VB.OptionButton optAdmin 
               BackColor       =   &H00FFFFFF&
               Caption         =   "User Administrator"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   0
               TabIndex        =   5
               Top             =   0
               Width           =   2055
            End
         End
         Begin VB.Line Line7 
            X1              =   120
            X2              =   2520
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   2520
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Line Line5 
            X1              =   120
            X2              =   2520
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   2520
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line3 
            X1              =   1920
            X2              =   5280
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Line Line2 
            X1              =   1800
            X2              =   5280
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Information"
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
            Top             =   120
            Width           =   1410
         End
         Begin VB.Image Image2 
            Height          =   645
            Left            =   120
            Picture         =   "frmCreateAccount.frx":02B0
            Top             =   480
            Width           =   705
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   5280
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Image Image1 
            Height          =   705
            Left            =   120
            Picture         =   "frmCreateAccount.frx":1B22
            Top             =   2280
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter User Name:"
            Height          =   195
            Left            =   1080
            TabIndex        =   14
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            Height          =   195
            Left            =   1080
            TabIndex        =   13
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm Password"
            Height          =   195
            Left            =   1080
            TabIndex        =   12
            Top             =   1200
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "User Account Type"
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
            TabIndex        =   11
            Top             =   1920
            Width           =   1650
         End
      End
   End
End
Attribute VB_Name = "frmCreateAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstusernames As ADODB.Recordset
Dim blnuserexist As Boolean
Dim blntextvalidation As Boolean
Dim blnmaxuser As Boolean

Private Sub chkShowPassword_Click()
If chkShowPassword.Value = 1 Then
    txtNewPassword.PasswordChar = ""
    txtConfirmPassword.PasswordChar = ""
ElseIf chkShowPassword = 0 Then
    txtNewPassword.PasswordChar = "*"
    txtConfirmPassword.PasswordChar = "*"
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCreate_Click()
Check_validation
If blntextvalidation = False Then
    Exit Sub
ElseIf blntextvalidation = True Then
    Check_Username
End If

'-----------------------------------------------------
On Err GoTo Err
If blnmaxuser = True Then
    blnmaxuser = False
    MsgBox "You have reached the maximum number of users." & vbCrLf & "No more User Accounts can be created...!", vbExclamation
    txtConfirmPassword = ""
    txtNewPassword = ""
    txtNewUserName = ""
    chkShowPassword.Value = 0
    txtConfirmPassword.Enabled = False
    txtNewPassword.Enabled = False
    txtNewUserName.Enabled = False
    chkShowPassword.Enabled = False
    optAdmin.Enabled = False
    optLimited.Enabled = False
    cmdCreate.Enabled = False
    Exit Sub
End If
If blnuserexist = True Then
    MsgBox "User name is already exist!", vbExclamation
    txtNewUserName.SetFocus
    SendKeys "{Home}+{End}"
  Exit Sub
ElseIf blnuserexist = False Then
rstusernames.AddNew
    rstusernames("USER_NAME") = txtNewUserName.Text
    rstusernames("PASSWORD") = txtConfirmPassword
    If optAdmin.Value = True Then
        rstusernames("TYPE") = "1"
        rstusernames("USER_REC_DELETE") = "1"
        rstusernames("USER_REC_ADD") = "1"
        rstusernames("USER_REC_EDIT") = "1"
        rstusernames("USER_REPORT_VIEW") = "1"
        rstusernames("USER_SEND_MSG") = "1"
    ElseIf optLimited.Value = True Then
        rstusernames("TYPE") = "0"
        rstusernames("USER_REC_DELETE") = "0"
        rstusernames("USER_REC_ADD") = "0"
        rstusernames("USER_REC_EDIT") = "0"
        rstusernames("USER_REPORT_VIEW") = "1"
        rstusernames("USER_SEND_MSG") = "1"
    End If
rstusernames.Update
MsgBox "User Account successfully created.", vbInformation
End If
frmPassword.cmbUsernames.Clear
frmPassword.Form_Load
txtNewUserName.Text = ""
txtNewPassword.Text = ""
txtConfirmPassword.Text = ""
txtNewUserName.SetFocus
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
Form_Load
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtNewUserName.SetFocus
End Sub

Private Sub Form_Load()
Set rstusernames = New ADODB.Recordset
    rstusernames.CursorLocation = adUseClient
    rstusernames.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
    optAdmin.Value = True
    Me.Picture = frmStyle.Picture
    'MsgBox rstusernames.RecordCount
If rstusernames.RecordCount = 8 Or rstusernames.RecordCount > 8 Then
    MsgBox "You have reached the maximum number of users." & vbCrLf & "No more User Accounts cannot be created...!", vbExclamation
    txtConfirmPassword.Enabled = False
    txtNewPassword.Enabled = False
    txtNewUserName.Enabled = False
    chkShowPassword.Enabled = False
    optAdmin.Enabled = False
    optLimited.Enabled = False
    cmdCreate.Enabled = False
End If
    Me.Left = frmPassword.Left + 1400
    Me.Top = frmPassword.Top + (frmPassword.Height - Me.Height) / 2
End Sub

Public Sub Check_Username()
 Set rstcheckusername = New ADODB.Recordset
     rstcheckusername.CursorLocation = adUseClient
      rstcheckusername.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & txtNewUserName & "'", dbcon, adOpenStatic, adLockReadOnly
                    'MsgBox rstusernames.RecordCount
                If rstusernames.RecordCount = 8 Then
                    blnmaxuser = True
                    Exit Sub
                End If
                
                If rstcheckusername.RecordCount > 0 Then
                    blnuserexist = True
                Else
                    blnuserexist = False
                End If
                
rstcheckusername.Close
Set rstcheckusername = Nothing
End Sub

Public Sub Check_validation()
   blntextvalidation = True
If UCase(txtNewUserName) = "ADMINISTRATOR" Or UCase(txtNewUserName) = "USER" Then
    MsgBox "The name you typed is a system built-in name." & vbCrLf & "Please type a different name.", vbExclamation
    txtNewUserName.SetFocus
    SendKeys "{Home}+{End}"
    blntextvalidation = False
   Exit Sub
End If
If txtNewUserName = "" Then
    MsgBox "Please type a User name!", vbExclamation
    txtNewUserName.SetFocus
    SendKeys "{Home}+{End}"
    blntextvalidation = False
   Exit Sub
End If
If txtNewPassword = "" Then
    MsgBox "Password is required!", vbExclamation
    txtNewPassword.SetFocus
    blntextvalidation = False
Exit Sub
End If
If txtConfirmPassword.Text = "" Then
    MsgBox "Confirm password is required!", vbExclamation
    txtConfirmPassword.SetFocus
    blntextvalidation = False
Exit Sub
End If
If txtNewPassword <> txtConfirmPassword Then
   MsgBox "Password confirmation failed." & vbCrLf & " Please enter passwords again.", vbCritical
   txtNewPassword.Text = ""
   txtConfirmPassword.Text = ""
   txtNewPassword.SetFocus
   blntextvalidation = False
   Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
rstusernames.Close
Set rstusernames = Nothing
End Sub
