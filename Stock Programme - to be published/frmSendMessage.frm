VERSION 5.00
Begin VB.Form frmSendMessage 
   BackColor       =   &H00FEFCFC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stock Controlling System Messaging System..."
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   5535
   Icon            =   "frmSendMessage.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6750
         Left            =   40
         ScaleHeight     =   6750
         ScaleWidth      =   5175
         TabIndex        =   10
         Top             =   240
         Width           =   5175
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
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
            Left            =   2350
            MouseIcon       =   "frmSendMessage.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   6360
            Width           =   1335
         End
         Begin VB.PictureBox Picture3 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   80
            Picture         =   "frmSendMessage.frx":015E
            ScaleHeight     =   195
            ScaleWidth      =   5055
            TabIndex        =   15
            Top             =   5940
            Width           =   5050
            Begin VB.Image imgProgress 
               Height          =   195
               Left            =   0
               Picture         =   "frmSendMessage.frx":3090
               Stretch         =   -1  'True
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
            Caption         =   "&Close"
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
            Left            =   3760
            MouseIcon       =   "frmSendMessage.frx":448A
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   6360
            Width           =   1335
         End
         Begin VB.CheckBox chkAdministratorOptions 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Administrator Options for Administrators"
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
            Left            =   410
            TabIndex        =   5
            Top             =   4320
            Width           =   3660
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Height          =   1215
            Left            =   120
            TabIndex        =   13
            Top             =   4320
            Width           =   4935
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   120
               ScaleHeight     =   855
               ScaleWidth      =   4695
               TabIndex        =   14
               Top             =   240
               Width           =   4695
               Begin VB.OptionButton optUserLogoutNotification 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "User System Log Off Notification"
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
                  Left            =   240
                  TabIndex        =   7
                  Top             =   480
                  Width           =   3495
               End
               Begin VB.OptionButton optSystemLogoutNotification 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "System Log Out Nitification"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   240
                  TabIndex        =   6
                  Top             =   120
                  Width           =   3255
               End
               Begin VB.Image Image2 
                  Height          =   735
                  Left            =   3900
                  Picture         =   "frmSendMessage.frx":45DC
                  Top             =   0
                  Width           =   720
               End
            End
         End
         Begin VB.CommandButton cmdSendMessage 
            Caption         =   "Send &Message"
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
            Left            =   3480
            MouseIcon       =   "frmSendMessage.frx":61AE
            MousePointer    =   99  'Custom
            TabIndex        =   4
            Top             =   3720
            Width           =   1575
         End
         Begin VB.CheckBox chkSendAll 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Send to All Active Users"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   3720
            Width           =   2895
         End
         Begin VB.ComboBox cmbLoggedUsers 
            BackColor       =   &H00FBF4F4&
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   3240
            Width           =   3375
         End
         Begin VB.TextBox txtMessage 
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
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   1
            Top             =   360
            Width           =   4935
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Press F5 to Refresh"
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
            Left            =   3345
            TabIndex        =   17
            Top             =   5640
            Width           =   1710
         End
         Begin VB.Line Line5 
            X1              =   840
            X2              =   2280
            Y1              =   6720
            Y2              =   6720
         End
         Begin VB.Line Line4 
            X1              =   840
            X2              =   2280
            Y1              =   6600
            Y2              =   6600
         End
         Begin VB.Line Line3 
            X1              =   840
            X2              =   2280
            Y1              =   6480
            Y2              =   6480
         End
         Begin VB.Line Line2 
            X1              =   840
            X2              =   2280
            Y1              =   6360
            Y2              =   6360
         End
         Begin VB.Label lblSendingstatus 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sending Status..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   5640
            Width           =   1635
         End
         Begin VB.Image Image1 
            Height          =   585
            Left            =   120
            Picture         =   "frmSendMessage.frx":6300
            Top             =   6165
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Active Users"
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
            TabIndex        =   12
            Top             =   3240
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type Your Message Here..."
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
            TabIndex        =   11
            Top             =   0
            Width           =   2595
         End
      End
   End
End
Attribute VB_Name = "frmSendMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstusernamesformsg As ADODB.Recordset
Dim rstsending_msg As ADODB.Recordset
'Dim blnUser_failed As Boolean
Dim admin_options_enabled As Boolean

Private Sub chkAdministratorOptions_Click()
If chkAdministratorOptions.Value = 1 Then
    optSystemLogoutNotification.Enabled = True
    optUserLogoutNotification.Enabled = True
    optSystemLogoutNotification.SetFocus
    admin_options_enabled = True
ElseIf chkAdministratorOptions.Value = 0 Then
    optSystemLogoutNotification.Enabled = False
    optUserLogoutNotification.Enabled = False
    admin_options_enabled = False
    'optSystemLogoutNotification.SetFocus
End If
End Sub

Private Sub chkAdministratorOptions_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub chkSendAll_Click()
If chkSendAll.Value = 1 Then
    cmbLoggedUsers.Enabled = False
ElseIf chkSendAll.Value = 0 Then
    cmbLoggedUsers.Enabled = True
End If
End Sub

Private Sub chkSendAll_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub cmbLoggedUsers_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub cmdCancel_Click()
'If blnUser_failed = False Then
 'If user_send_msg_privilege = 0 Then
    'Unload Me
    'Exit Sub
' End If
   ' SEND_MESSAGE_GLOBAL_OFF
'End If
Unload Me
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub cmdRefresh_Click()
cmbLoggedUsers.Clear
Data_Retrive
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub cmdSendMessage_Click()

If txtMessage.Text = "" Then
    'MsgBox "Message is empty.", vbCritical
    frmMsgErrorMessage.Label1.Caption = "Message can not be empty..."
    frmMsgErrorMessage.Show 1
    txtMessage.SetFocus
    Exit Sub
End If
If chkSendAll.Value = 0 Then
    If cmbLoggedUsers = User Then
        'MsgBox "You can't send messages to your Account itself.", vbCritical
        frmMsgErrorMessage.Label1.Caption = "You can't send messages to your Account itself."
        frmMsgErrorMessage.Show 1
        cmbLoggedUsers.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
End If

Send_Msg_Pro
If user_logged_out = True Then
    frmUserSendingMessage.Show 1
    user_logged_out = False
    Unload Me
    Exit Sub
End If
If intsending_all = 1 Then
    'intsending_all = 0
    'MsgBox "Message to All Logged Users sent succeeded.", vbInformation
    imgProgress.Width = 5055
    lblSendingstatus.Caption = "Message was sent."
    frmSentMsgMessage.Show 1
    Unload Me
Else
    'MsgBox "Message to " & cmbLoggedUsers & " sent succeeded.", vbInformation
    imgProgress.Width = 5055
    lblSendingstatus.Caption = "Message was sent."
    frmSentMsgMessage.Show 1
    Unload Me
End If
End Sub

Private Sub cmdSendMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
If sendwithactiveuserstatus = 1 Then
    sendwithactiveuserstatus = 0
        If receiveruser <> "" Then
            cmbLoggedUsers = receiveruser
            receiveruser = ""
        End If
End If

If sender <> "" Then
    cmbLoggedUsers.Text = sender
    sender = ""
    'txtMessage.SetFocus
End If
txtMessage.SetFocus
'MsgBox sender
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub Form_Load()
'Set rstusernamesformsg = New ADODB.Recordset
    'rstusernamesformsg.CursorLocation = adUseClient
    'rstusernamesformsg.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
    
'Set rstsending_msg = New ADODB.Recordset
    'rstsending_msg.CursorLocation = adUseClient
    'rstsending_msg.Open "SELECT * FROM MESSAGE ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
    
   
'CHECK_USER_SENDING_MSG
'Add_Logged_User_names_to_Combo
'If blnUser_failed = True Then
   'Unload Me
'End If
Data_Retrive
'MsgBox sender
If sender_bypass = False Then
    SEND_MESSAGE_GLOBAL_ON
End If
Me.Picture = frmStyle.Picture
If intaccount_type = 1 Then
    chkAdministratorOptions.Enabled = True
ElseIf intaccount_type = 0 Then
    chkAdministratorOptions.Enabled = False
    chkSendAll.Enabled = False
End If

optSystemLogoutNotification.Enabled = False
optUserLogoutNotification.Enabled = False

admin_options_enabled = False
If user_send_msg_privilege = 0 Then
   cmbLoggedUsers.Enabled = False
   txtMessage.Enabled = False
   optSystemLogoutNotification.Enabled = False
   optUserLogoutNotification.Enabled = False
   chkAdministratorOptions.Enabled = False
   chkSendAll.Enabled = False
   cmdSendMessage.Enabled = False
End If
End Sub
Public Sub Add_Logged_User_names_to_Combo()
'MsgBox sender
On Error Resume Next
If rstusernamesformsg.RecordCount > 0 Then
    Do While Not rstusernamesformsg.EOF
            If rstusernamesformsg("LOGGED_STATUS") = 1 Then
                'If rstusernamesformsg("USER_NAME") <> user Then
                    cmbLoggedUsers.AddItem rstusernamesformsg("USER_NAME")
                       
                'End If
            End If
            rstusernamesformsg.MoveNext
    Loop
End If

On Error Resume Next
cmbLoggedUsers.Text = cmbLoggedUsers.List(0)
'MsgBox cmbLoggedUsers.ListCount
If cmbLoggedUsers.ListCount = 1 Then
    chkSendAll.Enabled = False
End If
Exit Sub

Err:
MsgBox "No other Users Logged On to the system.", vbExclamation
cmdSendMessage.Enabled = False
End Sub
Public Sub SEND_MESSAGE_GLOBAL_ON()
On Error Resume Next
     rstsending_msg("MESSAGE_SENDING_GLOBAL") = "1"
     rstsending_msg("USER") = User
     rstsending_msg.Update

End Sub

Public Sub SEND_MESSAGE_GLOBAL_OFF()
On Error Resume Next
     rstsending_msg("MESSAGE_SENDING_GLOBAL") = "0"
     rstsending_msg("USER") = "User"
     rstsending_msg.Update
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If blnUser_failed = False Then
    'If user_send_msg_privilege = 0 Then
       ' Exit Sub
     'End If
 If sender_bypass = False Then
    SEND_MESSAGE_GLOBAL_OFF
 End If
 
 sender_bypass = False
  'End If
 On Error Resume Next
 rstusernamesformsg.Close
 Set rstusernamesformsg = Nothing
 
 rstsending_msg.Close
 Set rstsending_msg = Nothing
 
End Sub

Public Sub Send_Msg_Pro()
Dim receiver As String
On Error Resume Next
lblSendingstatus.Caption = "Sending Message..."
If chkSendAll.Value = 1 Then
    intsending_all = 1
    Set rstreceiver_all = New ADODB.Recordset
    rstreceiver_all.CursorLocation = adUseClient
    rstreceiver_all.Open "SELECT * FROM ACCOUNT_SET WHERE LOGGED_STATUS = '1' ", dbcon, adOpenStatic, adLockOptimistic
        Do While Not rstreceiver_all.EOF
            If rstreceiver_all("LOGGED_STATUS") = 1 Then
                If rstreceiver_all("USER_NAME") <> User Then
                    If admin_options_enabled = False Then
                        rstreceiver_all("INCOMMING_MSG") = txtMessage
                        rstreceiver_all("SENDER") = User
                        rstreceiver_all("MSG_TYPE") = "PUBLIC"
                        rstreceiver_all("MSG_SENT_TIME") = Time
                        rstreceiver_all("MSG_SENT_DATE") = Format(Date, "DD/mm/yyyy")
                        rstreceiver_all.Update
                    ElseIf admin_options_enabled = True Then
                        If optSystemLogoutNotification = True Then
                            rstreceiver_all("INCOMMING_MSG") = txtMessage
                            rstreceiver_all("SENDER") = User
                            rstreceiver_all("MSG_TYPE") = "SYSTEM LOG OUT NOTICE"
                            rstreceiver_all("SYSTEM_LOG_OUT_NTY") = "1"
                            rstreceiver_all("MSG_SENT_TIME") = Time
                            rstreceiver_all("MSG_SENT_DATE") = Format(Date, "DD/mm/yyyy")
                            rstreceiver_all.Update
                        ElseIf optUserLogoutNotification = True Then
                            rstreceiver_all("INCOMMING_MSG") = txtMessage
                            rstreceiver_all("SENDER") = User
                            rstreceiver_all("MSG_TYPE") = "USER LOG OFF NOTICE"
                            rstreceiver_all("USER_LOG_OFF_NTY") = "1"
                            rstreceiver_all("MSG_SENT_TIME") = Time
                            rstreceiver_all("MSG_SENT_DATE") = Format(Date, "DD/mm/yyyy")
                            rstreceiver_all.Update
                        End If
                    End If
                End If
            End If
            rstreceiver_all.MoveNext
        Loop

    rstreceiver_all.Close
    Set rstreceiver_all = Nothing
Else
    Set rstreceiver = New ADODB.Recordset
        rstreceiver.CursorLocation = adUseClient
        rstreceiver.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & cmbLoggedUsers & "'", dbcon, adOpenStatic, adLockOptimistic
            If rstreceiver("LOGGED_STATUS") = "0" Then
                'MsgBox "The User: " & cmbLoggedUsers & " has logged out.", vbExclamation
                user_logged_out = True
                Exit Sub
            End If
            If admin_options_enabled = False Then
                rstreceiver("INCOMMING_MSG") = txtMessage
                rstreceiver("SENDER") = User
                rstreceiver("MSG_TYPE") = "PRIVATE"
                rstreceiver("MSG_SENT_TIME") = Time
                rstreceiver("MSG_SENT_DATE") = Format(Date, "DD/mm/yyyy")
                rstreceiver.Update
            ElseIf admin_options_enabled = True Then
                If optSystemLogoutNotification = True Then
                    rstreceiver("INCOMMING_MSG") = txtMessage
                    rstreceiver("SENDER") = User
                    rstreceiver("MSG_TYPE") = "SYSTEM LOG OUT NOTICE"
                    rstreceiver("SYSTEM_LOG_OUT_NTY") = "1"
                    rstreceiver("MSG_SENT_TIME") = Time
                    rstreceiver("MSG_SENT_DATE") = Format(Date, "DD/mm/yyyy")
                    rstreceiver.Update
                ElseIf optUserLogoutNotification = True Then
                    rstreceiver("INCOMMING_MSG") = txtMessage
                    rstreceiver("SENDER") = User
                    rstreceiver("MSG_TYPE") = "USER LOG OFF NOTICE"
                    rstreceiver("USER_LOG_OFF_NTY") = "1"
                    rstreceiver("MSG_SENT_TIME") = Time
                    rstreceiver("MSG_SENT_DATE") = Format(Date, "DD/mm/yyyy")
                    rstreceiver.Update
                End If
            End If

    rstreceiver.Close
    Set rstreceiver = Nothing
End If
    
End Sub

Public Sub Data_Retrive()
On Error Resume Next
Set rstusernamesformsg = New ADODB.Recordset
    rstusernamesformsg.CursorLocation = adUseClient
    rstusernamesformsg.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
    
Set rstsending_msg = New ADODB.Recordset
    rstsending_msg.CursorLocation = adUseClient
    rstsending_msg.Open "SELECT * FROM MESSAGE ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
    
   
'CHECK_USER_SENDING_MSG
Add_Logged_User_names_to_Combo

End Sub

Private Sub optSystemLogoutNotification_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub optUserLogoutNotification_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub
