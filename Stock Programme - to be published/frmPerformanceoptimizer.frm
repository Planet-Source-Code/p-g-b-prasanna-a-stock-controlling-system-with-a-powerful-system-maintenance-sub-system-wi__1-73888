VERSION 5.00
Begin VB.Form frmPerformanceoptimizer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Optimize Performance"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   Icon            =   "frmPerformanceoptimizer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3810
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3435
         Left            =   120
         ScaleHeight     =   3435
         ScaleWidth      =   3735
         TabIndex        =   4
         Top             =   240
         Width           =   3735
         Begin VB.ListBox lstSourse 
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
            Height          =   960
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   1
            Top             =   360
            Width           =   3495
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
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
            Left            =   1320
            MouseIcon       =   "frmPerformanceoptimizer.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   3000
            Width           =   1095
         End
         Begin VB.CommandButton cmdOk 
            Cancel          =   -1  'True
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
            Left            =   2520
            MouseIcon       =   "frmPerformanceoptimizer.frx":015E
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   3600
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Image Image1 
            Height          =   690
            Left            =   120
            Picture         =   "frmPerformanceoptimizer.frx":02B0
            Top             =   2760
            Width           =   645
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Important Tip:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmPerformanceoptimizer.frx":1AAA
            Height          =   975
            Left            =   120
            TabIndex        =   6
            Top             =   1800
            Width           =   3495
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Select Soruce"
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
            TabIndex        =   5
            Top             =   0
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "frmPerformanceoptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstperformancetstatus As ADODB.Recordset
Private Sub Load_Data()
On Error GoTo Err
   If Not IsNull(rstperformancetstatus("BOOKS_ISSUE_GRID")) Then
        If rstperformancetstatus("BOOKS_ISSUE_GRID") = "1" Then
            lstSourse.Selected(0) = True
        Else
            lstSourse.Selected(0) = False
        End If
   Else
        lstSourse.Selected(0) = True
   End If
   
   If Not IsNull(rstperformancetstatus("BOOKS_RECEIPT_GRID")) Then
        If rstperformancetstatus("BOOKS_RECEIPT_GRID") = "1" Then
            lstSourse.Selected(1) = True
        Else
            lstSourse.Selected(1) = False
        End If
   Else
        lstSourse.Selected(1) = True
   End If
Exit Sub
Err:
MsgBox Err.Description & " _ " & Err.Number & ".", vbCritical
'Unload Me
Form_Load
End Sub

Private Sub cmdOk_Click()
cmdSave_Click
Unload Me
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
For i = 0 To lstSourse.ListCount - 1
     If lstSourse.Selected(i) = True Then
            If lstSourse.List(i) = "Enable Grid for Books Issue" Then
                 rstperformancetstatus("BOOKS_ISSUE_GRID") = "1"
                 rstperformancetstatus.Update
                 enabledataviewbooksissue = 1
            ElseIf lstSourse.List(i) = "Enable Grid for Books Receipt" Then
                 rstperformancetstatus("BOOKS_RECEIPT_GRID") = "1"
                 rstperformancetstatus.Update
                 enabledataviewbooksreceipt = 1
            End If
                 
     ElseIf lstSourse.Selected(i) = False Then
            If lstSourse.List(i) = "Enable Grid for Books Issue" Then
                 rstperformancetstatus("BOOKS_ISSUE_GRID") = "0"
                 rstperformancetstatus.Update
                 enabledataviewbooksissue = 0
            ElseIf lstSourse.List(i) = "Enable Grid for Books Receipt" Then
                 rstperformancetstatus("BOOKS_RECEIPT_GRID") = "0"
                 rstperformancetstatus.Update
                 enabledataviewbooksreceipt = 0
            End If
     End If
Next i
lstSourse.Refresh
End Sub

Private Sub Form_Load()
lstSourse.AddItem "Enable Grid for Books Issue"
lstSourse.AddItem "Enable Grid for Books Receipt"
Set rstperformancetstatus = New ADODB.Recordset
     rstperformancetstatus.CursorLocation = adUseClient
     rstperformancetstatus.Open "SELECT * FROM ACCOUNT_SET WHERE USER_NAME = '" & User & "'", dbcon, adOpenStatic, adLockOptimistic
Me.Picture = frmStyle.Picture
Load_Data
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
 rstperformancetstatus.Close
 Set rstperformancetstatus = Nothing
End Sub
