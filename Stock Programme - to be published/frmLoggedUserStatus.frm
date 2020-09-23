VERSION 5.00
Begin VB.Form frmLoggedUserStatus 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Active User Status..."
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   Icon            =   "frmLoggedUserStatus.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3550
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7680
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   40
         ScaleHeight     =   3375
         ScaleWidth      =   7500
         TabIndex        =   4
         Top             =   120
         Width           =   7500
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh Active User"
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
            Left            =   3960
            MouseIcon       =   "frmLoggedUserStatus.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   2880
            Width           =   2055
         End
         Begin VB.ListBox lstActiveUserTime 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   5520
            TabIndex        =   11
            Top             =   840
            Width           =   1935
         End
         Begin VB.ListBox lstActiveUserDate 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   3480
            TabIndex        =   10
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton cmdOk 
            Cancel          =   -1  'True
            Caption         =   "OK"
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
            Left            =   6120
            MouseIcon       =   "frmLoggedUserStatus.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   2880
            Width           =   1335
         End
         Begin VB.ListBox lstLoggedUserDisplay 
            Appearance      =   0  'Flat
            BackColor       =   &H00FAEDEF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   120
            TabIndex        =   1
            Top             =   840
            Width           =   3375
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Press F5 to Refresh Active User"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   4440
            TabIndex        =   12
            Top             =   2490
            Width           =   3015
         End
         Begin VB.Image Image3 
            Height          =   495
            Left            =   3420
            Picture         =   "frmLoggedUserStatus.frx":02B0
            Top             =   2835
            Width           =   510
         End
         Begin VB.Image Image5 
            Height          =   660
            Left            =   5520
            Picture         =   "frmLoggedUserStatus.frx":105A
            Top             =   120
            Width           =   645
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Logged Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   4250
            TabIndex        =   9
            Top             =   480
            Width           =   1110
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Logged Time"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   6240
            TabIndex        =   8
            Top             =   480
            Width           =   1110
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Active User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   900
            TabIndex        =   7
            Top             =   480
            Width           =   1005
         End
         Begin VB.Image Image2 
            Height          =   645
            Left            =   3480
            Picture         =   "frmLoggedUserStatus.frx":274C
            Top             =   120
            Width           =   705
         End
         Begin VB.Line Line6 
            X1              =   120
            X2              =   3840
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Line Line5 
            X1              =   120
            X2              =   3840
            Y1              =   3120
            Y2              =   3120
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   3840
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   3840
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   7440
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label lblCurrentUser 
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
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   1365
            TabIndex        =   6
            Top             =   2490
            Width           =   75
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current User:"
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
            Left            =   120
            TabIndex        =   5
            Top             =   2520
            Width           =   1170
         End
         Begin VB.Image Image1 
            Height          =   690
            Left            =   120
            Picture         =   "frmLoggedUserStatus.frx":3FBE
            Top             =   120
            Width           =   675
         End
      End
   End
End
Attribute VB_Name = "frmLoggedUserStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub cmdOk_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Public Sub cmdRefresh_Click()
lstLoggedUserDisplay.Clear
lstActiveUserTime.Clear
lstActiveUserDate.Clear
intlogstatusview = 0
Form_Load
End Sub

Private Sub cmdRefresh_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim rstUserlogstatusrestore As ADODB.Recordset
Me.Picture = frmStyle.Picture
Set rstUserlogstatusrestore = New ADODB.Recordset
    rstUserlogstatusrestore.CursorLocation = adUseClient
    rstUserlogstatusrestore.Open "SELECT * FROM ACCOUNT_SET WHERE LOGGED_STATUS = '" & 1 & "' ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly
    
If rstUserlogstatusrestore.RecordCount > 0 Then
 Dim intcounter As Integer
        If rstUserlogstatusrestore.RecordCount = 1 And rstUserlogstatusrestore("USER_NAME") = User Then
            If intlogstatusview = 1 Then
                MsgBox "No other users have logged in the system." & vbCrLf & "You can use the Restore Database action now.", vbInformation
            End If
              lstLoggedUserDisplay.AddItem rstUserlogstatusrestore("USER_NAME")
              lstActiveUserDate.AddItem rstUserlogstatusrestore("LOGGED_STATUS_DATE")
              lstActiveUserTime.AddItem rstUserlogstatusrestore("LOGGED_STATUS_TIME")
        Else
             If intlogstatusview = 1 Then
                MsgBox rstUserlogstatusrestore.RecordCount & " users are active on the system.", vbInformation
             End If
             
             intcounter = 1
                Do While Not rstUserlogstatusrestore.EOF
                    lstLoggedUserDisplay.AddItem rstUserlogstatusrestore("USER_NAME")
                    lstActiveUserDate.AddItem intcounter & " - " & rstUserlogstatusrestore("LOGGED_STATUS_DATE")
                    lstActiveUserTime.AddItem intcounter & " - " & rstUserlogstatusrestore("LOGGED_STATUS_TIME")
                    rstUserlogstatusrestore.MoveNext
                    intcounter = intcounter + 1
                Loop
        End If
End If
lstLoggedUserDisplay = lstLoggedUserDisplay.List(0)
lstActiveUserDate = lstActiveUserDate.List(lstLoggedUserDisplay.ListIndex)
lstActiveUserTime = lstActiveUserTime.List(lstLoggedUserDisplay.ListIndex)
lblCurrentUser = User
rstUserlogstatusrestore.Close
Set rstUserlogstatusrestore = Nothing
'Exit Sub
'Err:
'MsgBox Err.Description & " -" & Err.Number, vbCritical
'rstuserlogstatusrestore.Close
'Set rstuserlogstatusrestore = Nothing
'Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = False
End Sub

Private Sub lstActiveUserDate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lstLoggedUserDisplay = lstLoggedUserDisplay.List(0)
End Sub

Private Sub lstActiveUserTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'lstLoggedUserDisplay = lstLoggedUserDisplay.List(0)
End Sub

Private Sub lstLoggedUserDisplay_Click()
'sender = lstLoggedUserDisplay.List(lstLoggedUserDisplay.ListIndex)
'strinstantuser = lstLoggedUserDisplay.List(lstLoggedUserDisplay.ListIndex)
'strclearuser = lstLoggedUserDisplay.List(lstLoggedUserDisplay.ListIndex)
lstActiveUserDate = lstActiveUserDate.List(lstLoggedUserDisplay.ListIndex)
'lstActiveUserDate.Refresh
lstActiveUserTime = lstActiveUserTime.List(lstLoggedUserDisplay.ListIndex)
'lstActiveUserTime.Refresh
'lstLoggedUserDisplay.Refresh
End Sub

Private Sub lstLoggedUserDisplay_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF5 Then
        cmdRefresh_Click
    End If
End Sub

Private Sub lstLoggedUserDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim intc As Integer
receiveruser = lstLoggedUserDisplay.List(lstLoggedUserDisplay.ListIndex)
strinstantuser = lstLoggedUserDisplay.List(lstLoggedUserDisplay.ListIndex)
strclearuser = lstLoggedUserDisplay.List(lstLoggedUserDisplay.ListIndex)
'intc = lstLoggedUserDisplay.ListIndex
lstActiveUserDate = lstActiveUserDate.List(lstLoggedUserDisplay.ListIndex)
'lstActiveUserDate.Refresh
lstActiveUserTime = lstActiveUserTime.List(lstLoggedUserDisplay.ListIndex)
If Button = vbRightButton Then
    If intaccount_type = 0 Then
        Exit Sub
    End If
    'MsgBox lstLoggedUserDisplay.ListCount
    'If lstLoggedUserDisplay.ListCount = 1 Then
        'Exit Sub
    'End If
    'If lstLoggedUserDisplay.Selected(intc) = True Then
   'MsgBox sender
   sendwithactiveuserstatus = 1
   PopupMenu frmMain.mnuPopupActiveUser, , , , frmMain.mnuSendaMessage
   cmdRefresh_Click
   'Timer1.Enabled = True
End If
'lstActiveUserTime.Refresh
'lstLoggedUserDisplay.Refresh
End Sub

Public Sub Timer1_Timer()
cmdRefresh_Click
Timer1.Enabled = False
End Sub
