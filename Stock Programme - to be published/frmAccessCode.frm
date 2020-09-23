VERSION 5.00
Begin VB.Form frmAccessCode 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Access Code..."
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   80
      ScaleHeight     =   1215
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtAccessCode 
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
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "•"
         TabIndex        =   1
         ToolTipText     =   "Access Code Area"
         Top             =   360
         Width           =   4335
      End
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
         Left            =   2160
         MouseIcon       =   "frmAccessCode.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   720
         Width           =   1095
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
         Height          =   375
         Left            =   3360
         MouseIcon       =   "frmAccessCode.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Please type the Access Code for this action."
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
         TabIndex        =   4
         Top             =   0
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   465
         Left            =   120
         Picture         =   "frmAccessCode.frx":02A4
         Top             =   675
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmAccessCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intProceed_Active_User As Integer
Private Sub cmdCancel_Click()
sender_bypass = False: sender_clear = False
forcelogoff = False: forcelogout = False
blnclearuser = False: proceedforce = False
blnDataDelete = False: clearuserok = False
Unload Me
End Sub

Private Sub cmdOk_Click()
If sender_bypass = True Then
    If txtAccessCode = "stock01" Then
        Unload Me: frmSendMessage.Show 1
    Else
        MsgBox "Invalid Access Code.", vbCritical
        txtAccessCode = "": txtAccessCode.SetFocus
    End If
End If
If sender_clear = True Then
    If txtAccessCode = "stock02" Then
        Set rstmessagereset = New ADODB.Recordset
        rstmessagereset.CursorLocation = adUseClient
        rstmessagereset.Open "SELECT * FROM MESSAGE ORDER BY ID", dbcon, adOpenStatic, adLockOptimistic
            rstmessagereset("MESSAGE_SENDING_GLOBAL") = "0"
            rstmessagereset("USER") = "User"
            rstmessagereset.Update
            sender_clear = False
            rstmessagereset.Close
            Set rstmessagereset = Nothing
            Unload Me
    Else
        On Error Resume Next
        MsgBox "Invalid Access Code.", vbCritical
        txtAccessCode = "": txtAccessCode.SetFocus
    End If
End If
    
If forcelogout = True Then
   If txtAccessCode = "stock03" Then
        forcelogout = False: proceedforce = True: Unload Me
   Else
        On Error Resume Next
        MsgBox "Invalid Access Code.", vbCritical
        txtAccessCode = "": txtAccessCode.SetFocus
   End If
End If
If forcelogoff = True Then
   If txtAccessCode = "stock04" Then
        forcelogoff = False: proceedforce = True: Unload Me
   Else
        On Error Resume Next
        MsgBox "Invalid Access Code.", vbCritical
        txtAccessCode = "": txtAccessCode.SetFocus
   End If
End If
If blnDataDelete = True Then
    Data_Delete_Pro
End If

If blnclearuser = True Then
    If txtAccessCode = "stock05" Then
        blnclearuser = False: clearuserok = True: Unload Me
    Else
        On Error Resume Next
        MsgBox "Invalid Access Code.", vbCritical
        txtAccessCode = "": txtAccessCode.SetFocus
    End If
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtAccessCode.SetFocus
End Sub

Private Sub Form_Load()
Me.Picture = frmStyle.Picture
cmdOk.Enabled = False
End Sub

Public Sub Data_Delete_Pro()
Dim Msg_Del_Confirm As String
Msg_Del_Confirm = "Are you sure you need to delete all records in " & frmDeleteAllRecords.cmbSelectTable & "?"

If txtAccessCode = "stock2009" Then
    Select Case frmDeleteAllRecords.cmbSelectTable.ListIndex
        Case 0: Unload Me
                If MsgBox(Msg_Del_Confirm, vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                    Unload Me: Call Record_Deletion(1)
                Else: Unload Me: End If
        Case 1: Unload Me
                If MsgBox(Msg_Del_Confirm, vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                    Unload Me: Call Record_Deletion(2)
                Else: Unload Me: End If
        Case 2: Unload Me
                If MsgBox(Msg_Del_Confirm, vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                    Unload Me: Call Record_Deletion(3)
                Else: Unload Me: End If
        Case 3: Unload Me
                If MsgBox(Msg_Del_Confirm, vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                    Unload Me: Call Record_Deletion(4)
                Else: Unload Me: End If
        Case 4: Unload Me
                If MsgBox(Msg_Del_Confirm, vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                    Unload Me: Call Record_Deletion(5)
                Else: Unload Me: End If
        Case 5: Unload Me
                If MsgBox(Msg_Del_Confirm, vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                    Unload Me: Call Record_Deletion(6)
                Else: Unload Me: End If
        Case 6: Unload Me
                If MsgBox(Msg_Del_Confirm, vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                    Unload Me: Call Record_Deletion(7)
                Else: Unload Me: End If
        Case 7: Unload Me
                If MsgBox(Msg_Del_Confirm, vbYesNo + vbInformation + vbDefaultButton2) = vbYes Then
                    Get_Logged_User_Count
                        If intProceed_Active_User = 0 Then: Unload Me: frmSendMessage.Show 1: Exit Sub
                    Unload Me: Call Record_Deletion(8)
                Else: Unload Me: End If
        Case 8: Unload Me
                If MsgBox(Msg_Del_Confirm, vbYesNo + vbInformation) = vbYes Then
                   If User <> "Administrator" Then: MsgBox "This operation can only be done by User Name: Administrator.", vbCritical: Unload Me: Exit Sub
                   Get_Logged_User_Count
                        If intProceed_Active_User = 0 Then: Unload Me: frmSendMessage.Show 1: Exit Sub
                   Unload Me: Call Record_Deletion(9)
                Else: Unload Me: End If
    End Select
Else
      MsgBox "Invalid Access Code.", vbCritical
      txtAccessCode = "": txtAccessCode.SetFocus
End If
End Sub

Private Sub txtAccessCode_Change()
If txtAccessCode <> "" Then
    cmdOk.Enabled = True
Else
    cmdOk.Enabled = False
End If
   
End Sub

Public Sub Get_Logged_User_Count()
Set rstloggedusercount = New ADODB.Recordset
    rstloggedusercount.CursorLocation = adUseClient
    rstloggedusercount.Open "SELECT * FROM ACCOUNT_SET WHERE LOGGED_STATUS = '" & 1 & "' ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly
    
If rstloggedusercount.RecordCount > 0 Then
        If rstloggedusercount.RecordCount = 1 And rstloggedusercount("USER_NAME") = User Then
           intProceed_Active_User = 1
        Else
            MsgBox "Some users are still using the system." & vbCrLf & "Please log out the current Active Users.", vbExclamation
            intProceed_Active_User = 0
        End If
End If
Exit Sub
Err:
MsgBox Err.Description & " _" & Err.Number, vbCritical
End Sub
