VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFindBookReceipt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find...?"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   3825
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
      Left            =   2640
      MouseIcon       =   "frmFindBookReceipt.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4080
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
      Left            =   1440
      MouseIcon       =   "frmFindBookReceipt.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4080
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
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3100
         Left            =   30
         ScaleHeight     =   3105
         ScaleWidth      =   3525
         TabIndex        =   9
         Top             =   720
         Width           =   3530
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Value Status"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3050
            Left            =   40
            TabIndex        =   10
            Top             =   0
            Width           =   3375
            Begin VB.Frame Frame3 
               BackColor       =   &H00FFFFFF&
               Height          =   1455
               Left            =   120
               TabIndex        =   12
               Top             =   840
               Width           =   3135
               Begin MSComCtl2.DTPicker dtpFrom 
                  Height          =   375
                  Left            =   1200
                  TabIndex        =   4
                  Top             =   360
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   661
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CustomFormat    =   "dd/MM/yyyy"
                  Format          =   24117251
                  CurrentDate     =   39710
               End
               Begin MSComCtl2.DTPicker dtpTo 
                  Height          =   375
                  Left            =   1200
                  TabIndex        =   5
                  Top             =   840
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   661
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CustomFormat    =   "dd/MM/yyyy"
                  Format          =   24117251
                  CurrentDate     =   39710
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "From Date"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   14
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "To Date"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   13
                  Top             =   840
                  Width           =   855
               End
            End
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
               Left            =   1320
               TabIndex        =   6
               Top             =   2400
               Width           =   1695
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   670
               Left            =   120
               ScaleHeight     =   675
               ScaleWidth      =   3135
               TabIndex        =   11
               Top             =   240
               Width           =   3135
               Begin VB.OptionButton optDate 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Da&te"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   3
                  Top             =   360
                  Width           =   855
               End
               Begin VB.OptionButton optAll 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "&All"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   120
                  TabIndex        =   2
                  Top             =   0
                  Width           =   615
               End
            End
            Begin VB.Image Image2 
               Height          =   525
               Left            =   840
               Picture         =   "frmFindBookReceipt.frx":02A4
               Top             =   2400
               Width           =   495
            End
            Begin VB.Label Label2 
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
               Height          =   255
               Left            =   240
               TabIndex        =   15
               Top             =   2400
               Width           =   855
            End
         End
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
         TabIndex        =   16
         Top             =   360
         Width           =   960
      End
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   1320
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   1320
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   1320
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   1320
      Y1              =   4080
      Y2              =   4080
   End
End
Attribute VB_Name = "frmFindBookReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
If dtpFrom.Value > dtpTo.Value Then
    MsgBox "Invalid Date Range.", vbExclamation
    dtpFrom.SetFocus
    Exit Sub
End If

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
If optAll.Value = True Then
    blnenabledate = False
ElseIf optDate.Value = True Then
    blnenabledate = True
    fromdate = Format(dtpFrom.Value, "dd/mm/yyyy")
    todate = Format(dtpTo.Value, "dd/mm/yyyy")
    'MsgBox fromdate & vbCrLf & todate
End If
blnfind_status = True
Find_Val = txtFindValue.Text
frmBookReceipt.Find_for_Details
'Unload Me
End Sub

Private Sub Form_Activate()
optAll.Value = True
txtFindValue.SetFocus

End Sub

Private Sub Form_Load()
dtpFrom.Value = Date
dtpTo.Value = Date
'cmbValue.AddItem "Date"
cmbValue.AddItem "Module"
cmbValue.AddItem "Course"
cmbValue.AddItem "Category"
cmbValue.Text = cmbValue.List(0)
Me.Picture = frmStyle.Picture

Me.Top = frmBookReceipt.Top + (frmBookReceipt.Height - Me.Height) / 2
Me.Left = frmBookReceipt.Left + (frmBookReceipt.Width - Me.Width) / 2

End Sub

Private Sub optAll_Click()
Frame3.Enabled = False
dtpFrom.Enabled = False
dtpTo.Enabled = False
End Sub

Private Sub optDate_Click()
Frame3.Enabled = True
dtpFrom.Enabled = True
dtpTo.Enabled = True
End Sub

