VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmUserLogStatusReports 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Log Status Reports"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   Icon            =   "frmUserLogStatusReports.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4540
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2840
         Left            =   40
         ScaleHeight     =   2835
         ScaleWidth      =   4395
         TabIndex        =   8
         Top             =   120
         Width           =   4400
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
            Left            =   3240
            MouseIcon       =   "frmUserLogStatusReports.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   2400
            Width           =   1095
         End
         Begin VB.CommandButton cmdPreview 
            Caption         =   "&View Report"
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
            Left            =   1800
            MouseIcon       =   "frmUserLogStatusReports.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "Re&fresh"
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
            Left            =   3000
            MouseIcon       =   "frmUserLogStatusReports.frx":02B0
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   480
            Width           =   1335
         End
         Begin VB.CheckBox chkDateRange 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enable Date Range"
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
            TabIndex        =   3
            Top             =   960
            Width           =   2655
         End
         Begin VB.ComboBox cmblogUsers 
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
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   480
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker dtpFrom 
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   1680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   48234499
            CurrentDate     =   40027
         End
         Begin MSComCtl2.DTPicker dtpTo 
            Height          =   375
            Left            =   1920
            TabIndex        =   5
            Top             =   1680
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   48234499
            CurrentDate     =   40027
         End
         Begin VB.Line Line5 
            X1              =   720
            X2              =   1680
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Line Line4 
            X1              =   840
            X2              =   1680
            Y1              =   2640
            Y2              =   2640
         End
         Begin VB.Line Line3 
            X1              =   840
            X2              =   1680
            Y1              =   2520
            Y2              =   2520
         End
         Begin VB.Line Line2 
            X1              =   840
            X2              =   1680
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Image Image2 
            Height          =   660
            Left            =   120
            Picture         =   "frmUserLogStatusReports.frx":0402
            Top             =   2160
            Width           =   675
         End
         Begin VB.Image Image1 
            Height          =   780
            Left            =   3720
            Picture         =   "frmUserLogStatusReports.frx":1BA4
            Top             =   1080
            Width           =   630
         End
         Begin VB.Line Line1 
            X1              =   600
            X2              =   4320
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "To:"
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
            Left            =   1920
            TabIndex        =   11
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "From:"
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
            TabIndex        =   10
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Report Status"
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
            TabIndex        =   9
            Top             =   120
            Width           =   1785
         End
      End
   End
End
Attribute VB_Name = "frmUserLogStatusReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstlogusers As ADODB.Recordset

Private Sub chkDateRange_Click()
If chkDateRange.Value = 1 Then
    dtpFrom.Enabled = True
    dtpTo.Enabled = True
Else
    dtpFrom.Enabled = False
    dtpTo.Enabled = False
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdPreview_Click()
If Check_For_Privilege(4) = True Then: Exit Sub
Report_Generate_Pro
End Sub

Private Sub cmdRefresh_Click()
Add_User_Names_to_Combo
End Sub

Private Sub Form_Load()
Me.Picture = frmStyle.Picture
Add_User_Names_to_Combo
dtpFrom.Value = Date
dtpTo.Value = Date
'dtpFrom = Format(dtpFrom.Value, "dd/mm/yyyy")
'dtpTo = Format(dtpTo.Value, "dd/mm/yyyy")
End Sub
Public Sub Add_User_Names_to_Combo()
On Error GoTo Err
Set rstlogusers = New ADODB.Recordset
    rstlogusers.CursorLocation = adUseClient
    rstlogusers.Open "SELECT * FROM ACCOUNT_SET ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly
cmblogUsers.Clear
cmblogUsers.AddItem "----- All Users -----"
If rstlogusers.RecordCount > 0 Then
    Do While Not rstlogusers.EOF
        cmblogUsers.AddItem rstlogusers("USER_NAME")
    rstlogusers.MoveNext
    Loop
End If

rstlogusers.Close
Set rstlogusers = Nothing

cmblogUsers = cmblogUsers.List(0)
Exit Sub
Err:
MsgBox Err.Description & "_" & Err.Number, vbCritical
End Sub

Public Sub Report_Generate_Pro()
On Error GoTo Err
If cmblogUsers = "----- All Users -----" Then
    If chkDateRange.Value = 1 Then
         If dtpFrom.Value > dtpTo.Value Then
                MsgBox "Invalid Date Range.." & vbCrLf & "Check the Date Range.", vbExclamation
                dtpFrom.SetFocus
                Exit Sub
        End If
        Set rstlogstatusallusersdate = New ADODB.Recordset
            rstlogstatusallusersdate.CursorLocation = adUseClient
            rstlogstatusallusersdate.Open "SELECT * FROM USER_LOG WHERE DATE >= " & "#" & dtpFrom & "#" & " AND DATE <= " & "#" & dtpTo & "#" & " ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly
            'MsgBox rstlogstatusallusersdate.RecordCount
        Set drLogStatus.DataSource = rstlogstatusallusersdate
            drLogStatus.Sections("Section4").Controls("Label1").Caption = "Log Status Report for - All Users"
            drLogStatus.Sections("Section4").Controls("Label13").Caption = "From:"
            drLogStatus.Sections("Section4").Controls("Label7").Caption = Format(dtpFrom, "dd/MM/yyyy")
            drLogStatus.Sections("Section4").Controls("Label12").Caption = "To:"
            drLogStatus.Sections("Section4").Controls("Label11").Caption = Format(dtpTo, "dd/MM/yyyy")
            drLogStatus.Show 1
        rstlogstatusallusersdate.Close
        Set rstlogstatusallusersdate = Nothing
    Else
        Set rstlogstatusallusers = New ADODB.Recordset
            rstlogstatusallusers.CursorLocation = adUseClient
            rstlogstatusallusers.Open "SELECT * FROM USER_LOG ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly
        Set drLogStatus.DataSource = rstlogstatusallusers
            drLogStatus.Sections("Section4").Controls("Label1").Caption = "Log Status Report for - All Users"
            drLogStatus.Sections("Section4").Controls("Label13").Caption = ""
            drLogStatus.Sections("Section4").Controls("Label7").Caption = ""
            drLogStatus.Sections("Section4").Controls("Label12").Caption = ""
            drLogStatus.Sections("Section4").Controls("Label11").Caption = ""
            drLogStatus.Show 1
        rstlogstatusallusers.Close
        Set rstlogstatusallusers = Nothing
    End If
Else
    If chkDateRange.Value = 1 Then
           If dtpFrom.Value > dtpTo.Value Then
                MsgBox "Invalid Date Range.", vbExclamation
                dtpFrom.SetFocus
                Exit Sub
           End If
        Set rstlogstatususerdate = New ADODB.Recordset
            rstlogstatususerdate.CursorLocation = adUseClient
            rstlogstatususerdate.Open "SELECT * FROM USER_LOG WHERE USER = '" & cmblogUsers & "' AND DATE >= " & "#" & dtpFrom & "#" & " AND DATE <= " & "#" & dtpTo & "#" & " ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly
        Set drLogStatus.DataSource = rstlogstatususerdate
            drLogStatus.Sections("Section4").Controls("Label1").Caption = "Log Status Report for - " & cmblogUsers
            drLogStatus.Sections("Section4").Controls("Label13").Caption = "From:"
            drLogStatus.Sections("Section4").Controls("Label7").Caption = Format(dtpFrom, "dd/MM/yyyy")
            drLogStatus.Sections("Section4").Controls("Label12").Caption = "To:"
            drLogStatus.Sections("Section4").Controls("Label11").Caption = Format(dtpTo, "dd/MM/yyyy")
            drLogStatus.Show 1
        rstlogstatususerdate.Close
        Set rstlogstatususerdate = Nothing
    Else

        Set rstlogstatususer = New ADODB.Recordset
            rstlogstatususer.CursorLocation = adUseClient
            rstlogstatususer.Open "SELECT * FROM USER_LOG WHERE USER = '" & cmblogUsers & "' ORDER BY ID", dbcon, adOpenStatic, adLockReadOnly
        Set drLogStatus.DataSource = rstlogstatususer
            drLogStatus.Sections("Section4").Controls("Label1").Caption = "Log Status Report for - " & cmblogUsers
            drLogStatus.Sections("Section4").Controls("Label13").Caption = ""
            drLogStatus.Sections("Section4").Controls("Label7").Caption = ""
            drLogStatus.Sections("Section4").Controls("Label12").Caption = ""
            drLogStatus.Sections("Section4").Controls("Label11").Caption = ""
            drLogStatus.Show 1
        rstlogstatususer.Close
        Set rstlogstatususer = Nothing
    End If
End If
Exit Sub

Err:
MsgBox Err.Description & " _ " & Err.Description & ".", vbCritical
End Sub
