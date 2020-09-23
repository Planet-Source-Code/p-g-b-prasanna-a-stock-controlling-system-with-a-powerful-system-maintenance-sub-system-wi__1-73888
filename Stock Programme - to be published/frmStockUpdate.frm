VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStockUpdate 
   Caption         =   "Books Stock Update"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdIssue 
      Caption         =   "&Update Book Issue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   3495
   End
   Begin VB.CommandButton cmdReceipt 
      Caption         =   "Update Book Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2880
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtNull 
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSComctlLib.StatusBar stbReceipt 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3330
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5159
            MinWidth        =   5159
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbIssues 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   3030
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5159
            MinWidth        =   5159
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prgbarUpdate 
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
      Max             =   1
   End
   Begin VB.Label lbl1 
      BorderStyle     =   1  'Fixed Single
      Height          =   15
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lbl2 
      BorderStyle     =   1  'Fixed Single
      Height          =   15
      Left            =   2880
      TabIndex        =   13
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lbl3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1725
      Left            =   3960
      TabIndex        =   12
      Top             =   240
      Width           =   15
   End
   Begin VB.Label lbl4 
      BorderStyle     =   1  'Fixed Single
      Height          =   15
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   3855
   End
   Begin VB.Label lbl5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1725
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   15
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Updating Book Issue/Receipt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label lblUpdate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   15
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   3855
   End
End
Attribute VB_Name = "frmStockUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
