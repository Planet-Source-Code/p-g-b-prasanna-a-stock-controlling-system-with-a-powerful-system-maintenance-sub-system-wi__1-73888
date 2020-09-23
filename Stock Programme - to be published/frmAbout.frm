VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   480
   ClientWidth     =   4875
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   80
      Width           =   4650
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4695
         Left            =   40
         ScaleHeight     =   4695
         ScaleWidth      =   4560
         TabIndex        =   1
         Top             =   120
         Width           =   4560
         Begin VB.PictureBox Picture4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   3360
            ScaleHeight     =   495
            ScaleWidth      =   1095
            TabIndex        =   4
            Top             =   4200
            Width           =   1095
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
               Left            =   120
               MouseIcon       =   "frmAbout.frx":000C
               MousePointer    =   99  'Custom
               TabIndex        =   5
               Top             =   45
               Width           =   975
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   0
            ScaleHeight     =   855
            ScaleWidth      =   4575
            TabIndex        =   3
            Top             =   0
            Width           =   4575
            Begin VB.Image Image2 
               Height          =   540
               Left            =   120
               Picture         =   "frmAbout.frx":015E
               Top             =   120
               Width           =   525
            End
            Begin VB.Image Image1 
               Height          =   690
               Left            =   840
               Picture         =   "frmAbout.frx":10D0
               Top             =   120
               Width           =   3660
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   2760
            ScaleHeight     =   495
            ScaleWidth      =   1815
            TabIndex        =   2
            Top             =   2520
            Width           =   1815
            Begin VB.Image Image3 
               Height          =   375
               Left            =   80
               Picture         =   "frmAbout.frx":949A
               Top             =   0
               Width           =   1605
            End
         End
         Begin VB.Label lblContact 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "pgbsoft@gmail.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            MouseIcon       =   "frmAbout.frx":B480
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   2760
            Width           =   1395
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Chaminda Wijegunawardhana who gave me the idea of Stock Controlling System."
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   3480
            Width           =   4335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "---------------------------"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   105
            TabIndex        =   13
            Top             =   3255
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Special Thanks To:"
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
            Left            =   105
            TabIndex        =   12
            Top             =   3120
            Width           =   1680
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "P.G. Bandula Prasanna."
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   100
            TabIndex        =   11
            Top             =   2520
            Width           =   2775
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "-------------------------------------------"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   100
            TabIndex        =   10
            Top             =   2295
            Width           =   2595
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Designed and programmed by:"
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
            Left            =   105
            TabIndex        =   9
            Top             =   2160
            Width           =   2580
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":B5D2
            Height          =   615
            Left            =   100
            TabIndex        =   8
            Top             =   1320
            Width           =   4365
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "All rights reserved.   BC Systems 2009."
            ForeColor       =   &H000000C0&
            Height          =   495
            Left            =   1665
            TabIndex        =   7
            Top             =   4200
            Width           =   1530
         End
         Begin VB.Line Line1 
            X1              =   100
            X2              =   4440
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label4"
            Height          =   45
            Left            =   105
            TabIndex        =   6
            Top             =   4080
            Width           =   4380
         End
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = frmStyle.Picture
End Sub

Private Sub lblContact_Click()
ShellExecute hwnd, "Open", "mailto:pgbsoft@gmail.com", vbNullString, vbNullString, SW_SHOW
End Sub

