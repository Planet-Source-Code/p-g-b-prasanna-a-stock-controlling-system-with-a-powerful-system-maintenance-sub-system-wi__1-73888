VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Login"
   ClientHeight    =   1875
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4065
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1107.812
   ScaleMode       =   0  'User
   ScaleWidth      =   3816.815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   68
      Width           =   3855
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   40
         ScaleHeight     =   1455
         ScaleWidth      =   3735
         TabIndex        =   5
         Top             =   120
         Width           =   3735
         Begin VB.TextBox txtPassword 
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
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   1305
            PasswordChar    =   "•"
            TabIndex        =   2
            Top             =   525
            Width           =   2325
         End
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
            Height          =   390
            Left            =   2500
            MouseIcon       =   "frmLogin.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   1020
            Width           =   1120
         End
         Begin VB.CommandButton cmdOK 
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
            Height          =   390
            Left            =   1300
            MouseIcon       =   "frmLogin.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   1020
            Width           =   1140
         End
         Begin VB.ComboBox cmbSelectUser 
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
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   120
            Width           =   2340
         End
         Begin VB.Image Image1 
            Height          =   660
            Left            =   360
            Picture         =   "frmLogin.frx":02B0
            Top             =   800
            Width           =   660
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "&Password:"
            Height          =   270
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   540
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            BackStyle       =   0  'Transparent
            Caption         =   "&User Name:"
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   150
            Width           =   1080
         End
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstUsernames As ADODB.Recordset
Dim rstExpire As ADODB.Recordset
'Dim rstuserloggedstatus As ADODB.Recordset
Dim password As String
'Dim intresolutionAlertenable As Integer


Private Sub cmbSelectUser_Click()
    Set rstthemewithuser = New ADODB.Recordset
        rstthemewithuser.CursorLocation = adUseClient
        On Error GoTo Err_
        rstthemewithuser.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbSelectUser & "'", dbcon, adOpenStatic, adLockReadOnly
            If rstthemewithuser.RecordCount > 0 Then
                If Not IsNull(rstthemewithuser("THEME_SET")) Then
                    inttheme = rstthemewithuser("THEME_SET")
                Else
                    inttheme = 1
                End If
            End If
            
    frmStyle.Form_Load
    Me.Picture = frmStyle.Picture
    Me.Refresh
    frmSplash.Picture = frmStyle.Picture
    frmSplash.Refresh
    rstthemewithuser.Close
    Set rstthemewithuser = Nothing
    txtPassword = ""
    'txtPassword.SetFocus
    Exit Sub
   
Err_:
    MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
End Sub
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdCancel_KeyPress(KeyAscii As Integer)
If KeyAscii = 162 Then
    PWORD_INFO
ElseIf KeyAscii = 214 Then
    Make_My_Pro_Expire1
ElseIf KeyAscii = 220 Then
    Make_My_Pro_Expire0
End If
End Sub

Private Sub cmdOk_Click()
intsetlog = 1
Check_Password
End Sub
Private Sub Form_Activate()
txtPassword.SetFocus
End Sub

Private Sub Form_Load()
Set rstExpire = New ADODB.Recordset
    rstExpire.CursorLocation = adUseClient
    On Error GoTo db_Error
    rstExpire.Open "SELECT * FROM EXPIRE ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
    intexpire = rstExpire("EXPIRE")
    
Set rstUsernames = New ADODB.Recordset
    rstUsernames.CursorLocation = adUseClient
    rstUsernames.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly

Add_User_Names_to_Combo
Retrieve_User
Me.Picture = frmStyle.Picture
On Error GoTo Err
If ret_user <> "" Then
cmbSelectUser.Text = ret_user
Else
cmbSelectUser.Text = cmbSelectUser.List(0)
End If
Exit Sub
Err:
cmbSelectUser.Text = cmbSelectUser.List(0)
Exit Sub
db_Error:
MsgBox "Database Error_" & Err.Number & "." & vbCrLf & "Stock Contorlling System cannot continue." & vbCrLf & "Replace the database with a backup.", vbCritical
End
End Sub
Public Sub Add_User_Names_to_Combo()
If rstUsernames.RecordCount > 0 Then
    Do While Not rstUsernames.EOF
        cmbSelectUser.AddItem rstUsernames("USER_NAME")
        rstUsernames.MoveNext
    Loop
    End If
End Sub
Public Sub Check_Password()
If intexpire = 1 Then
    MsgBox "Stock Controling System has expired." & vbCrLf & "Contact Mr. Bandula." & vbCrLf & "pgbsoft@gmail.com.", vbCritical
  End
End If
 Set rstgetpassword = New ADODB.Recordset
     rstgetpassword.CursorLocation = adUseClient
      rstgetpassword.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbSelectUser & "'", dbcon, adOpenStatic, adLockOptimistic
      password = rstgetpassword("PASSWORD")
      intaccount_type = rstgetpassword("TYPE")
      
      If Not IsNull(rstgetpassword("USER_REC_DELETE")) Then
         user_record_delete_privilege = rstgetpassword("USER_REC_DELETE")
      Else
         user_record_delete_privilege = 0
      End If
      If Not IsNull(rstgetpassword("USER_REC_ADD")) Then
        user_record_add_privilege = rstgetpassword("USER_REC_ADD")
      Else
        user_record_add_privilege = 0
      End If
      If Not IsNull(rstgetpassword("USER_REC_EDIT")) Then
        
        user_record_edit_privilege = rstgetpassword("USER_REC_EDIT")
        'MsgBox user_record_edit_privilege
      Else
        user_record_edit_privilege = 0
      End If
      If Not IsNull(rstgetpassword("USER_REPORT_VIEW")) Then
        user_view_report_privilege = rstgetpassword("USER_REPORT_VIEW")
      Else
        user_view_report_privilege = 0
      End If
      If Not IsNull(rstgetpassword("USER_SEND_MSG")) Then
        user_send_msg_privilege = rstgetpassword("USER_SEND_MSG")
      Else
        user_send_msg_privilege = 0
      End If
      If Not IsNull(rstgetpassword("BOOKS_ISSUE_GRID")) Then
        enabledataviewbooksissue = rstgetpassword("BOOKS_ISSUE_GRID")
      Else
        enabledataviewbooksissue = 1
      End If
      If Not IsNull(rstgetpassword("BOOKS_RECEIPT_GRID")) Then
        enabledataviewbooksreceipt = rstgetpassword("BOOKS_RECEIPT_GRID")
      Else
        enabledataviewbooksreceipt = 1
      End If
      User = cmbSelectUser
        If txtPassword = password Then
            If Not IsNull(rstgetpassword("ENABLED")) Then
                If rstgetpassword("ENABLED") = "0" Then
                    MsgBox "Your Account has been disabled." & vbCrLf & "Please, contact Administrator.", vbCritical
                        txtPassword = ""
                        txtPassword.SetFocus
                        Exit Sub
                End If
            End If
        rstgetpassword.Close
        Set rstgetpassword = Nothing
        User_Log_In
        Unload frmSplash
        Unload Me
        Store_User
        Store_User_Logged_Status_Login
        Check_Resolution_Alert
            If intresolutionAlertenable = 0 Then
                blntheme_apply = True
            ElseIf intresolutionAlertenable = 1 Then
                blntheme_apply = False
            End If
        frmMain.Show
      Else
        MsgBox "Invalid Password, try again!", vbCritical, "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
      End If
'-------------------------------------------------------------------
End Sub

Public Sub Make_My_Pro_Expire1()
On Error Resume Next
 rstExpire("EXPIRE") = "1"
 rstExpire.Update
 End
End Sub
Public Sub Make_My_Pro_Expire0()
On Error Resume Next
 rstExpire("EXPIRE") = "0"
 rstExpire.Update
 End
End Sub
Public Sub Store_User()
On Error Resume Next
reg_obj.RegWrite (loggeduser), User
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload frmStyle
Unload frmSplash

rstUsernames.Close
Set rstExpire = Nothing

rstExpire.Close
Set rstExpire = Nothing
End Sub

