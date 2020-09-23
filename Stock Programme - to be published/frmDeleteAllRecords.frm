VERSION 5.00
Begin VB.Form frmDeleteAllRecords 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete  Database Records"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   Icon            =   "frmDeleteAllRecords.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
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
      Left            =   4050
      MouseIcon       =   "frmDeleteAllRecords.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5380
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1410
         Left            =   40
         ScaleHeight     =   1410
         ScaleWidth      =   5220
         TabIndex        =   4
         Top             =   240
         Width           =   5220
         Begin VB.PictureBox Picture2 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   120
            Picture         =   "frmDeleteAllRecords.frx":015E
            ScaleHeight     =   195
            ScaleWidth      =   5055
            TabIndex        =   5
            Top             =   1260
            Width           =   5050
            Begin VB.Image imgProgress 
               Height          =   195
               Left            =   0
               Picture         =   "frmDeleteAllRecords.frx":3090
               Stretch         =   -1  'True
               Top             =   0
               Width           =   15
            End
         End
         Begin VB.ComboBox cmbSelectTable 
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
            Top             =   120
            Width           =   5055
         End
         Begin VB.CommandButton cmdDeleteRecords 
            Caption         =   "Delete Records"
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
            Left            =   3360
            MouseIcon       =   "frmDeleteAllRecords.frx":448A
            MousePointer    =   99  'Custom
            TabIndex        =   2
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label lblDeleting 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   840
            TabIndex        =   7
            Top             =   960
            Width           =   60
         End
         Begin VB.Label lblReccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   840
            TabIndex        =   6
            Top             =   600
            Width           =   60
         End
         Begin VB.Image Image1 
            Height          =   510
            Left            =   120
            Picture         =   "frmDeleteAllRecords.frx":45DC
            Top             =   480
            Width           =   495
         End
      End
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   3960
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3960
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3960
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3960
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "frmDeleteAllRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDeleteRecords_Click()
'frmSecurityCode.Show 1
blnDataDelete = True
frmAccessCode.Show 1
End Sub

Private Sub Form_Load()
cmbSelectTable.AddItem "BOOKS ISSUE"
cmbSelectTable.AddItem "BOOKS RECEIPT"
cmbSelectTable.AddItem "CURRENT STOCK"
cmbSelectTable.AddItem "DETAILS"
cmbSelectTable.AddItem "COURSE"
cmbSelectTable.AddItem "MODULE"
cmbSelectTable.AddItem "CATEGORY"
cmbSelectTable.AddItem "LOG STATUS"
cmbSelectTable.AddItem "USER ACCOUNTS"
cmbSelectTable = cmbSelectTable.List(0)
Me.Picture = frmStyle.Picture
End Sub
    

