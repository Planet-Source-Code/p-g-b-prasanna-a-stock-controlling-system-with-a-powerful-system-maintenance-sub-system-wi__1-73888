VERSION 5.00
Begin VB.Form frmClearDatabaseLocation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clear Databae Location..."
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "frmClearDatabaseLocation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmClearDatabaseLocation.frx":617A
   ScaleHeight     =   1545
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Database Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1050
         Left            =   120
         ScaleHeight     =   1050
         ScaleWidth      =   7125
         TabIndex        =   4
         Top             =   200
         Width           =   7120
         Begin VB.CommandButton cmdClear 
            Caption         =   "&Clear Database Location"
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
            MouseIcon       =   "frmClearDatabaseLocation.frx":8C8E
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   600
            Width           =   2415
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "C&lose"
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
            Left            =   6000
            MouseIcon       =   "frmClearDatabaseLocation.frx":8DE0
            MousePointer    =   99  'Custom
            TabIndex        =   3
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtDbLocation 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBF4F4&
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   6
               Charset         =   255
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   120
            Width           =   7095
         End
         Begin VB.Line Line5 
            X1              =   0
            X2              =   2760
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   2760
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   2760
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   2760
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   2880
            Picture         =   "frmClearDatabaseLocation.frx":8F32
            Top             =   480
            Width           =   540
         End
      End
   End
End
Attribute VB_Name = "frmClearDatabaseLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
On Error GoTo Err
If MsgBox("Are you sure you want to clear the Database Location?", vbQuestion + vbYesNo) = vbYes Then
    'use dll function to clear the database location.
    Clear_Db_Location
    txtDbLocation = ""
    End
End If
Exit Sub
Err:
MsgBox "db_cls_config.dll was not found..." & vbCrLf & "Process was aborted.", vbCritical
End
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub Form_Load()
On Error GoTo Err
Get_Db_Location
txtDbLocation = db_location
Exit Sub
Err:
MsgBox "db_cls_config.dll was not found..." & vbCrLf & "Process was aborted.", vbCritical
End
End Sub
