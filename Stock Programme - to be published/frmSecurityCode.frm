VERSION 5.00
Begin VB.Form frmSecurityCode 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Security Code"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1300
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1100
         Left            =   40
         ScaleHeight     =   1095
         ScaleWidth      =   6255
         TabIndex        =   1
         Top             =   120
         Width           =   6255
         Begin VB.TextBox txtSecurityCode 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1920
            PasswordChar    =   "*"
            TabIndex        =   4
            Top             =   120
            Width           =   4215
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
            Left            =   3720
            MouseIcon       =   "frmSecurityCode.frx":0000
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   600
            Width           =   1215
         End
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
            Left            =   5040
            MouseIcon       =   "frmSecurityCode.frx":0152
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   600
            Width           =   1095
         End
         Begin VB.Image Image1 
            Height          =   720
            Left            =   360
            Picture         =   "frmSecurityCode.frx":02A4
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Security Code"
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
            TabIndex        =   5
            Top             =   120
            Width           =   1710
         End
      End
   End
End
Attribute VB_Name = "frmSecurityCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cmdCancel_Click()
'Unload Me
'End Sub

'Private Sub cmdOk_Click()
'If txtSecurityCode = "stockcontrol20082009" Then
'Select Case frmDeleteAllRecords.cmbSelectTable.ListIndex
'Case 0
'If MsgBox("Are you sure you need to delete all records in " & frmDeleteAllRecords.cmbSelectTable & "?", vbYesNo + vbInformation) = vbYes Then
'    frmDeleteAllRecords.Records_Deletion_Books_Issue
'End If
'Case 1
'If MsgBox("Are you sure you need to delete all records in " & frmDeleteAllRecords.cmbSelectTable & "?", vbYesNo + vbInformation) = vbYes Then
'    frmDeleteAllRecords.Records_Deletion_Books_Receipt
'End If
'Case 2
'If MsgBox("Are you sure you need to delete all records in " & frmDeleteAllRecords.cmbSelectTable & "?", vbYesNo + vbInformation) = vbYes Then
'    frmDeleteAllRecords.Records_Deletion_Current_Stock
'End If
'Case 3
'If MsgBox("Are you sure you need to delete all records in " & frmDeleteAllRecords.cmbSelectTable & "?", vbYesNo + vbInformation) = vbYes Then
 '   frmDeleteAllRecords.Records_Deletion_Details
'End If
'Case 4
'If MsgBox("Are you sure you need to delete all records in " & frmDeleteAllRecords.cmbSelectTable & "?", vbYesNo + vbInformation) = vbYes Then
'    frmDeleteAllRecords.Records_Deletion_Course
'End If
'Case 5
'If MsgBox("Are you sure you need to delete all records in " & frmDeleteAllRecords.cmbSelectTable & "?", vbYesNo + vbInformation) = vbYes Then
'    frmDeleteAllRecords.Records_Deletion_Module
'End If
'Case 6
'If MsgBox("Are you sure you need to delete all records in " & frmDeleteAllRecords.cmbSelectTable & "?", vbYesNo + vbInformation) = vbYes Then
'    frmDeleteAllRecords.Records_Deletion_Category
'End If
'End Select
'Else
'MsgBox "Invalid Security Code", vbCritical
'End If
'End Sub

'Private Sub Form_Activate()
'txtSecurityCode.SetFocus
'End Sub

'Private Sub Form_Load()
'Me.Picture = frmStyle.Picture
'End Sub
