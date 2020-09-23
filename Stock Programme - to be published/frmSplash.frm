VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7200
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   600
         ScaleHeight     =   1575
         ScaleWidth      =   5895
         TabIndex        =   1
         Top             =   480
         Width           =   5895
         Begin VB.Image Image2 
            Height          =   540
            Left            =   120
            Picture         =   "frmSplash.frx":0000
            Top             =   480
            Width           =   525
         End
         Begin VB.Image Image1 
            Height          =   1050
            Left            =   720
            Picture         =   "frmSplash.frx":0F72
            Top             =   480
            Width           =   4800
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designed for Windows XP/Windows Vista/Windows 7"
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
         Left            =   1785
         TabIndex        =   8
         Top             =   2055
         Width           =   4230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Specially Developed for"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   4080
         TabIndex        =   7
         Top             =   3135
         Width           =   2040
      End
      Begin VB.Image Image3 
         Height          =   540
         Left            =   3360
         Picture         =   "frmSplash.frx":11634
         Top             =   3360
         Width           =   2745
      End
      Begin VB.Label lblCopyright 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning:                                                                      This computer program is protected by copyright law."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   4200
         Width           =   4935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "P.G. Bandula Prasanna."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "---------------------------------------------"
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
         Height          =   150
         Left            =   720
         TabIndex        =   4
         Top             =   2775
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Designed and Developed by"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "All rights reserved.   BC Systems 2009."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   720
         TabIndex        =   2
         Top             =   3360
         Width           =   2250
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   5655
      Left            =   20
      Top             =   20
      Width           =   7400
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = frmStyle.Picture
End Sub
