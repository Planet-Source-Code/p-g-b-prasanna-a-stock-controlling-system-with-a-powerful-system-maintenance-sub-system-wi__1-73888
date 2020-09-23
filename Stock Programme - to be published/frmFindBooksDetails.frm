VERSION 5.00
Begin VB.Form frmFindBooksDetails 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find...?"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
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
      MouseIcon       =   "frmFindBooksDetails.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1680
      Width           =   975
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
      MouseIcon       =   "frmFindBooksDetails.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
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
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtFindValue 
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
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox cmbValue 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   525
         Left            =   720
         Picture         =   "frmFindBooksDetails.frx":02A4
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   960
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   960
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   960
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   960
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "frmFindBooksDetails"
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
If cmbValue = cmbValue.List(0) Then
    detailsfindparameter = 0
ElseIf cmbValue = cmbValue.List(1) Then
    detailsfindparameter = 1
ElseIf cmbValue = cmbValue.List(2) Then
    detailsfindparameter = 2
End If
blnfind_status = True
Find_Val = txtFindValue.Text
frmBookdetails.Find_for_Details
'Unload Me
End Sub

Private Sub Form_Activate()
txtFindValue.SetFocus

End Sub

Private Sub Form_Load()
cmbValue.AddItem "Module"
cmbValue.AddItem "Course"
cmbValue.AddItem "Category"
cmbValue.Text = cmbValue.List(0)
Me.Picture = frmStyle.Picture

Me.Top = frmBookdetails.Top + (frmBookdetails.Height - Me.Height) / 2
Me.Left = frmBookdetails.Left + (frmBookdetails.Width - Me.Width) / 2

End Sub
