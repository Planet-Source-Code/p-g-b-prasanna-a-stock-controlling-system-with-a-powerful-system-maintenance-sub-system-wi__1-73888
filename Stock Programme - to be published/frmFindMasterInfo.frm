VERSION 5.00
Begin VB.Form frmFindMasterInfo 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find...?"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Parameter"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtFindValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   525
         Left            =   720
         Picture         =   "frmFindMasterInfo.frx":0000
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Module Code"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
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
      Left            =   1080
      MouseIcon       =   "frmFindMasterInfo.frx":0DEE
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1440
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
      Left            =   2280
      MouseIcon       =   "frmFindMasterInfo.frx":0F40
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   960
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   960
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   960
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   960
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmFindMasterInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
If txtFindValue.Text = "" Then
    MsgBox "Value is empty!", vbExclamation
    txtFindValue.SetFocus
    Exit Sub
End If
find_check = True
blnfind_status = True
Find_Val = txtFindValue.Text
frmMasterInformation.Find_for_Details
'Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtFindValue.SetFocus

End Sub

Private Sub Form_Load()
Me.Picture = frmStyle.Picture
Me.Top = frmMasterInformation.Top + (frmMasterInformation.Height - Me.Height) / 2
Me.Left = frmMasterInformation.Left + (frmMasterInformation.Width - Me.Width) / 2
End Sub
