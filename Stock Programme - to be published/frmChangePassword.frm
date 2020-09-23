VERSION 5.00
Begin VB.Form frmChangePassword 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password/Privileges"
   ClientHeight    =   9480
   ClientLeft      =   5175
   ClientTop       =   2820
   ClientWidth     =   5775
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
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
      MouseIcon       =   "frmChangePassword.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   9000
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   120
      TabIndex        =   20
      Top             =   720
      Width           =   5535
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7890
         Left            =   40
         ScaleHeight     =   7890
         ScaleWidth      =   5325
         TabIndex        =   21
         Top             =   200
         Width           =   5325
         Begin VB.ComboBox cmbUserAccountStatus 
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
            TabIndex        =   11
            Top             =   4440
            Width           =   1575
         End
         Begin VB.CommandButton cmdChangeAccountStatus 
            Appearance      =   0  'Flat
            Caption         =   "Change Account Status"
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
            Left            =   2715
            MouseIcon       =   "frmChangePassword.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   4440
            Width           =   2535
         End
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
            Left            =   840
            TabIndex        =   5
            Top             =   1600
            Width           =   1695
         End
         Begin VB.CheckBox chkSendMessages 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Allow Sending Messages"
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
            Left            =   1200
            TabIndex        =   17
            Top             =   6960
            Width           =   4095
         End
         Begin VB.CommandButton cmdChangePrivileges 
            Caption         =   "Change Privileges"
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
            Left            =   2760
            MouseIcon       =   "frmChangePassword.frx":02B0
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   7440
            Width           =   2535
         End
         Begin VB.CheckBox chkViewReports 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Allow Viewing Reports"
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
            Left            =   1200
            TabIndex        =   16
            Top             =   6600
            Width           =   4095
         End
         Begin VB.CheckBox chkRecEditing 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Allow System-wide Record Editing"
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
            Left            =   1200
            TabIndex        =   15
            Top             =   6240
            Width           =   4095
         End
         Begin VB.CheckBox chkRecAdding 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Allow System-wide Record Adding"
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
            Left            =   1200
            TabIndex        =   14
            Top             =   5880
            Width           =   4095
         End
         Begin VB.CheckBox chkRecDeletion 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Allow System-wide Record Deleting"
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
            Left            =   1200
            TabIndex        =   13
            Top             =   5520
            Width           =   4095
         End
         Begin VB.CommandButton cmdChangeUserType 
            Caption         =   "Change User Type"
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
            Left            =   2710
            MouseIcon       =   "frmChangePassword.frx":0402
            MousePointer    =   99  'Custom
            TabIndex        =   10
            Top             =   3360
            Width           =   2535
         End
         Begin VB.OptionButton optLimited 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Limited"
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
            Left            =   3960
            TabIndex        =   9
            Top             =   2880
            Width           =   1335
         End
         Begin VB.OptionButton optAdmin 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Administrator"
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
            TabIndex        =   8
            Top             =   2880
            Width           =   1935
         End
         Begin VB.TextBox txtOldPassword 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2595
            PasswordChar    =   "•"
            TabIndex        =   2
            Top             =   480
            Width           =   2655
         End
         Begin VB.TextBox txtNewPassword 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2595
            PasswordChar    =   "•"
            TabIndex        =   3
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtConfirmPassword 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   2595
            PasswordChar    =   "•"
            TabIndex        =   4
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CommandButton cmdChangePassword 
            Caption         =   "Change Password"
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
            Left            =   2715
            MouseIcon       =   "frmChangePassword.frx":0554
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Line Line9 
            X1              =   2760
            X2              =   5640
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Change User Account Status"
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
            TabIndex        =   28
            Top             =   3840
            Width           =   2655
         End
         Begin VB.Image Image4 
            Height          =   750
            Left            =   120
            Picture         =   "frmChangePassword.frx":06A6
            Top             =   4200
            Width           =   690
         End
         Begin VB.Label lblResetPassword 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reset User Password..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   3000
            MouseIcon       =   "frmChangePassword.frx":2240
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   2160
            Width           =   2220
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "User Action Privileges"
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
            Left            =   120
            TabIndex        =   27
            Top             =   5160
            Width           =   2055
         End
         Begin VB.Line Line3 
            X1              =   2160
            X2              =   5280
            Y1              =   5280
            Y2              =   5280
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   120
            Picture         =   "frmChangePassword.frx":2392
            Top             =   5640
            Width           =   750
         End
         Begin VB.Image Image3 
            Height          =   750
            Left            =   120
            Picture         =   "frmChangePassword.frx":4054
            Top             =   2880
            Width           =   780
         End
         Begin VB.Image Image2 
            Height          =   990
            Left            =   0
            Picture         =   "frmChangePassword.frx":5F0E
            Top             =   720
            Width           =   735
         End
         Begin VB.Line Line2 
            X1              =   1920
            X2              =   5280
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Change User Type"
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
            TabIndex        =   26
            Top             =   2520
            Width           =   1695
         End
         Begin VB.Line Line1 
            X1              =   1920
            X2              =   5280
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Change Password"
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
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Old Password"
            Height          =   195
            Left            =   840
            TabIndex        =   24
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "New Password"
            Height          =   195
            Left            =   840
            TabIndex        =   23
            Top             =   840
            Width           =   1065
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Confirm New Password"
            Height          =   195
            Left            =   840
            TabIndex        =   22
            Top             =   1200
            Width           =   1635
         End
      End
   End
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   4200
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   4200
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   4200
      Y1              =   9120
      Y2              =   9120
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   4200
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select User Name"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1545
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstusernames As ADODB.Recordset
Dim oldpass As String

Private Sub chkShowPassword_Click()
If chkShowPassword.Value = 1 Then
    txtOldPassword.PasswordChar = ""
    txtNewPassword.PasswordChar = ""
    txtConfirmPassword.PasswordChar = ""
ElseIf chkShowPassword.Value = 0 Then
    txtOldPassword.PasswordChar = "*"
    txtNewPassword.PasswordChar = "*"
    txtConfirmPassword.PasswordChar = "*"
End If
End Sub
Private Sub cmbUsernames_Click()
Commands_Set
End Sub

Private Sub cmdChangeAccountStatus_Click()
On Error GoTo Err
 Set rstaccountstatus = New ADODB.Recordset
     rstaccountstatus.CursorLocation = adUseClient
     rstaccountstatus.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockOptimistic
     If cmbUserAccountStatus.Text = "Enabled" Then
        rstaccountstatus("ENABLED") = "1"
        rstaccountstatus.Update
        MsgBox "The user " & cmbUsernames & " Enabled.", vbInformation
     ElseIf cmbUserAccountStatus = "Disabled" Then
        rstaccountstatus("ENABLED") = "0"
        rstaccountstatus.Update
        MsgBox "The user " & cmbUsernames & " Disabled.", vbInformation
     End If
 rstaccountstatus.Close
 Set rstaccountstatus = Nothing
Exit Sub
Err:
    MsgBox Err.Description & " _ " & Err.Number & "."
     rstaccountstatus.Close
     Set rstaccountstatus = Nothing

End Sub

Private Sub cmdChangePassword_Click()
Password_Change_Pro
End Sub

Private Sub cmdChangePrivileges_Click()
User_Privileges_Change_Pro
End Sub

Private Sub cmdChangeUserType_Click()
User_Type_Change_Pro
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_KeyPress(KeyAscii As Integer)
If KeyAscii = 162 Then
    PWORD_INFO
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtOldPassword.SetFocus
If cmbUsernames <> User Then
    lblResetPassword.Enabled = True
Else
    lblResetPassword.Enabled = False
End If

'MsgBox " left " & frmChangePassword.Left & vbCrLf & "top - " & frmChangePassword.Top
End Sub

Private Sub Form_Load()
Set rstusernames = New ADODB.Recordset
    rstusernames.CursorLocation = adUseClient
    rstusernames.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
    
Add_User_Names_to_Combo

cmbUserAccountStatus.Clear
cmbUserAccountStatus.AddItem "Enabled"
cmbUserAccountStatus.AddItem "Disabled"

On Error Resume Next
cmbUsernames.Text = cmbUsernames.List(0)
Account_type_initialize
Me.Picture = frmStyle.Picture
'If Screen.Width = 15360 And Screen.Height = 11520 Then
    'Me.Left = 6810
    'Me.Top = 3585
'ElseIf Screen.Width = 12000 And Screen.Height = 9000 Then
    'Me.Left = 5130
    'Me.Top = 2340
'End If
    Me.Left = frmPassword.Left + 1300
    Me.Top = frmPassword.Top - 1450
End Sub
Public Sub Add_User_Names_to_Combo()
If rstusernames.RecordCount > 0 Then
    cmbUsernames.Clear
    Do While Not rstusernames.EOF
        cmbUsernames.AddItem rstusernames("USER_NAME")
    rstusernames.MoveNext
    Loop
End If
rstusernames.Close
Set rstusernames = Nothing
End Sub

Public Sub Password_Change_Pro()
 Set rstgetusername = New ADODB.Recordset
     rstgetusername.CursorLocation = adUseClient
      rstgetusername.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockOptimistic
      oldpass = rstgetusername("PASSWORD")
      
     If txtOldPassword <> oldpass Then
        MsgBox "Old password is wrong.", vbCritical
        txtOldPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
     End If
     If txtNewPassword.Text = "" Then
        MsgBox "Password is required.", vbExclamation
        txtNewPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
     End If
     If txtConfirmPassword.Text = "" Then
        MsgBox "Confirm password is required.", vbExclamation
        txtConfirmPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
     End If
     If txtNewPassword <> txtConfirmPassword Then
        MsgBox "Password confirmation failed." & vbCrLf & "Please enter passwords again.", vbCritical
        txtNewPassword.Text = ""
        txtConfirmPassword.Text = ""
        txtNewPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
     ElseIf txtNewPassword = txtConfirmPassword Then
        rstgetusername("PASSWORD") = txtConfirmPassword
        rstgetusername.Update
        MsgBox "Password successfully changed." & vbCrLf & "Log in again for the changes.", vbInformation
        txtOldPassword.Text = ""
        txtNewPassword.Text = ""
        txtConfirmPassword.Text = ""
        txtOldPassword.SetFocus
     End If
rstgetusername.Close
Set rstgetusername = Nothing
End Sub

Public Sub User_Type_Change_Pro()
On Error GoTo Err
 Set rstgetusername = New ADODB.Recordset
     rstgetusername.CursorLocation = adUseClient
      rstgetusername.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockOptimistic
        If optAdmin.Value = True Then
            rstgetusername("TYPE") = "1"
            rstgetusername.Update
        ElseIf optLimited.Value = True Then
            rstgetusername("TYPE") = "0"
            rstgetusername.Update
        End If
MsgBox "User Type successfully changed." & vbCrLf & "Log In again for the changes." & _
vbCrLf & vbCrLf & "Tip" & vbCrLf & "---" & vbCrLf & "You may need to change the Privileges as well.", vbInformation
rstgetusername.Close
Set rstgetusername = Nothing
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
Form_Load
End Sub
Public Sub User_Privileges_Change_Pro()
On Error GoTo Err
 Set rstuserforprivileges = New ADODB.Recordset
     rstuserforprivileges.CursorLocation = adUseClient
      rstuserforprivileges.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockOptimistic
      
        rstuserforprivileges("USER_REC_DELETE") = chkRecDeletion.Value
        rstuserforprivileges("USER_REC_ADD") = chkRecAdding.Value
        rstuserforprivileges("USER_REC_EDIT") = chkRecEditing.Value
        rstuserforprivileges("USER_REPORT_VIEW") = chkViewReports.Value
        rstuserforprivileges("USER_SEND_MSG") = chkSendMessages.Value
        rstuserforprivileges.Update
        MsgBox "User Privileges successfully changed.", vbInformation
        
    rstuserforprivileges.Close
    Set rstuserforprivileges = Nothing
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
Form_Load
End Sub
Public Sub Account_type_initialize()
If intaccount_type = 0 Then
    On Error Resume Next
    cmbUsernames.Enabled = False
    optAdmin.Enabled = False
    optLimited.Enabled = False
    cmdChangeUserType.Enabled = False
    cmdChangeAccountStatus.Enabled = False
    cmbUserAccountStatus.Enabled = False
End If
cmbUsernames = User

End Sub

Private Sub lblResetPassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblResetPassword.Left = lblResetPassword.Left + 20
lblResetPassword.Top = lblResetPassword.Top + 20
End Sub

Private Sub lblResetPassword_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblResetPassword.ForeColor = &HC0&
End Sub

Private Sub lblResetPassword_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblResetPassword.Left = lblResetPassword.Left - 20
lblResetPassword.Top = lblResetPassword.Top - 20

On Error GoTo Err
 Set rstresetpassword = New ADODB.Recordset
     rstresetpassword.CursorLocation = adUseClient
     rstresetpassword.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockOptimistic
If MsgBox("Are you sure you want to reset the password for " & cmbUsernames & " ?", vbQuestion + vbYesNo) = vbYes Then
     rstresetpassword("PASSWORD") = "password"
     rstresetpassword.Update
     MsgBox "The password of " & cmbUsernames & " has been reset as 'password' ", vbInformation
End If

    rstresetpassword.Close
    Set rstresetpassword = Nothing
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'unload Me
Form_Load
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblResetPassword.ForeColor = &HC00000
End Sub

Public Sub Commands_Set()
On Error Resume Next
 Set rstchecktype = New ADODB.Recordset
     rstchecktype.CursorLocation = adUseClient
      rstchecktype.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbUsernames & "'", dbcon, adOpenStatic, adLockReadOnly
      If rstchecktype("TYPE") = "1" Then
        optAdmin.Value = True
      ElseIf rstchecktype("TYPE") = "0" Then
        optLimited.Value = True
      End If
      If rstchecktype("USER_REC_DELETE") = "1" Then
         chkRecDeletion.Value = 1
      Else
        chkRecDeletion.Value = 0
      End If
      If rstchecktype("USER_REC_ADD") = "1" Then
         chkRecAdding.Value = 1
      Else
         chkRecAdding.Value = 0
      End If
      If rstchecktype("USER_REC_EDIT") = "1" Then
         chkRecEditing.Value = 1
      Else
         chkRecEditing.Value = 0
      End If
      If rstchecktype("USER_REPORT_VIEW") = "1" Then
         chkViewReports.Value = 1
      Else
         chkViewReports.Value = 0
      End If
      If rstchecktype("USER_SEND_MSG") = "1" Then
         chkSendMessages.Value = 1
      Else
         chkSendMessages.Value = 0
      End If
      
      If IsNull(rstchecktype("ENABLED")) Then
         cmbUserAccountStatus = "Enabled"
            
      ElseIf rstchecktype("ENABLED") = "1" Then
         cmbUserAccountStatus = "Enabled"
         
      ElseIf rstchecktype("ENABLED") = "0" Then
         cmbUserAccountStatus = "Disabled"
      End If
      
If cmbUsernames = "Administrator" Then
    'optAdmin.Enabled = False
    optLimited.Enabled = False
    cmdChangeUserType.Enabled = False
    lblResetPassword.Enabled = False
    
    cmbUserAccountStatus.Enabled = False
    cmdChangeAccountStatus.Enabled = False
    
    chkRecAdding.Enabled = False
    chkRecDeletion.Enabled = False
    chkRecEditing.Enabled = False
    chkViewReports.Enabled = False
    chkSendMessages.Enabled = False
    cmdChangePrivileges.Enabled = False
    
        If User <> "Administrator" Then
            txtOldPassword.Enabled = False
            txtNewPassword.Enabled = False
            txtConfirmPassword.Enabled = False
            chkShowPassword.Enabled = False
            chkShowPassword.Value = 0
            cmdChangePassword.Enabled = False
        ElseIf User = "Administrator" Then
            txtOldPassword.Enabled = True
            txtNewPassword.Enabled = True
            txtConfirmPassword.Enabled = True
            chkShowPassword.Value = 0
            chkShowPassword.Enabled = True
            cmdChangePassword.Enabled = True
        End If
Else
    If cmbUsernames = User Then
        txtOldPassword.Enabled = True
        txtNewPassword.Enabled = True
        txtConfirmPassword.Enabled = True
        chkShowPassword.Value = 0
        chkShowPassword.Enabled = True
        cmdChangePassword.Enabled = True
        
        lblResetPassword.Enabled = False
            
        optAdmin.Enabled = False
        optLimited.Enabled = False
        cmdChangeUserType.Enabled = False
        
        cmbUserAccountStatus.Enabled = False
        cmdChangeAccountStatus.Enabled = False
        
        chkRecDeletion.Enabled = False
        chkRecAdding.Enabled = False
        chkRecEditing.Enabled = False
        chkViewReports.Enabled = False
        chkSendMessages.Enabled = False
        cmdChangePrivileges.Enabled = False
    Else
        txtOldPassword.Enabled = False
        txtNewPassword.Enabled = False
        txtConfirmPassword.Enabled = False
        chkShowPassword.Value = 0
        chkShowPassword.Enabled = False
        cmdChangePassword.Enabled = False
        
        lblResetPassword.Enabled = True
        
        optAdmin.Enabled = True
        optLimited.Enabled = True
        cmdChangeUserType.Enabled = True
        
        cmbUserAccountStatus.Enabled = True
        cmdChangeAccountStatus.Enabled = True
        
        chkRecDeletion.Enabled = True
        chkRecAdding.Enabled = True
        chkRecEditing.Enabled = True
        chkViewReports.Enabled = True
        chkSendMessages.Enabled = True
        cmdChangePrivileges.Enabled = True
    End If
End If
rstchecktype.Close
Set rstchecktype = Nothing
End Sub
